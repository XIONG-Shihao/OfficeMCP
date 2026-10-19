import * as fs from "node:fs";
import path from "node:path";
import { execFile } from "node:child_process";
import { DatabaseSync } from "../infrastructure/sqlite.mjs";
import { DocxPool } from "../formats/docx/docx-pool.mjs";
import { sqliteOptions } from "../infrastructure/sqlite-options.mjs";
import { checkStateQuota, createQuotaSnapshot } from "../infrastructure/quota.mjs";
import { OfficeError, requireThat } from "../infrastructure/errors.mjs";
import { fingerprint, newId } from "../infrastructure/files.mjs";
import { TaskStore } from "./task-store.mjs";

/** Durable coordinator ownership, record storage and worker lifecycle. */
export class CoordinatorState {
  constructor(config, testing = {}) {
    this.now = testing.now ?? Date.now;
    this.checkpoint = testing.checkpoint ?? (() => {});
    this.activeRequests = 0;
    this.config = config;
    this.activeWorkers = new Set();
    this.quotaSnapshot = createQuotaSnapshot();
    this.docxPool = new DocxPool(this);
    this.taskStore = new TaskStore(this);
    this.jobs = this.taskStore.active;
    this.token = newId("owner");
    this.closed = false;
    try {
      this.ownerDb = new DatabaseSync(
        path.join(config.stateDirectory, "owner-lock.sqlite"),
        sqliteOptions(1024 * 1024),
      );
      this.ownerDb.exec(
        "PRAGMA journal_mode=DELETE; CREATE TABLE IF NOT EXISTS lease (singleton INTEGER PRIMARY KEY CHECK(singleton=1), pid INTEGER NOT NULL, token TEXT NOT NULL); BEGIN EXCLUSIVE;",
      );
      this.ownerDb
        .prepare("INSERT OR REPLACE INTO lease VALUES (1, ?, ?)")
        .run(process.pid, this.token);
    } catch {
      try {
        this.ownerDb?.close();
      } catch {
        // SQLite already owns the primary startup error.
      }
      throw new OfficeError(
        "STATE_IN_USE",
        "This private state directory already has a live server. Use one coordinator per state directory.",
      );
    }
    this.db = new DatabaseSync(
      path.join(config.stateDirectory, "state.sqlite"),
      sqliteOptions(64 * 1024 * 1024),
    );
    this.db.exec(
      "PRAGMA trusted_schema=OFF; PRAGMA journal_mode=WAL; PRAGMA synchronous=FULL; CREATE TABLE IF NOT EXISTS records (kind TEXT NOT NULL, id TEXT NOT NULL, value TEXT NOT NULL, PRIMARY KEY(kind,id));",
    );
    this.db.exec("BEGIN IMMEDIATE");
    try {
      const schema = this.get("meta", "schema", false);
      requireThat(
        !schema || [1, 2, 3].includes(schema.version),
        "UNSUPPORTED_STATE",
        "Unknown state schema version.",
      );
      const migrationHash = fingerprint({
        version: 2,
        change: "record lifecycle timestamps and explicit expiry tombstones",
      });
      if (schema?.version >= 2)
        requireThat(
          this.get("meta", "migration-2", false)?.checksum === migrationHash,
          "UNSUPPORTED_STATE",
          "State migration checksum mismatch.",
        );
      else {
        for (const row of this.db
          .prepare("SELECT kind,id,value FROM records")
          .all())
          this.put(row.kind, row.id, JSON.parse(row.value));
        this.checkpoint("migration-before-commit");
        this.put("meta", "migration-2", { checksum: migrationHash });
        this.put("meta", "schema", { version: 2 });
      }
      const removalHash = fingerprint({
        version: 3,
        change: "retire renderer layout records and visual publication requirements",
      });
      if (schema?.version === 3) {
        requireThat(this.get("meta", "migration-3", false)?.checksum === removalHash,
          "UNSUPPORTED_STATE", "State migration checksum mismatch.");
      } else {
        const requirements = new Set();
        for (const { id, validationRequirement, ...candidate } of this.rows("candidate")) {
          if (!validationRequirement) continue;
          requirements.add(validationRequirement.validationId);
          this.put("candidate", id, candidate);
        }
        let retiredValidations = 0;
        for (const { id, visualRequest, visualEvidence, ...validation } of this.rows("validation")) {
          if (!visualRequest && !visualEvidence && !requirements.has(id)) continue;
          // Old mixed structural/visual results cannot authorize publication. A new
          // semantic validation is required; preserve the prior outcome for audit.
          this.put("validation", id, { ...validation, state: "retired",
            result: { ...validation.result, passed: false },
            retirement: { reason: "rendering-removed", previousPassed: validation.result?.passed ?? false } });
          retiredValidations++;
        }
        const layouts = this.rows("docxLayout");
        const layoutPaths = new Set(layouts.map(layout => layout.path));
        for (const asset of this.rows("asset"))
          if (layoutPaths.has(asset.path)) this.delete("asset", asset.id);
        for (const layout of layouts) this.delete("docxLayout", layout.id);
        // Files are reclaimed by normal private-artifact cleanup, never by unlinking
        // arbitrary paths read from a legacy record inside a database transaction.
        this.checkpoint("rendering-removal-before-commit");
        this.put("meta", "migration-3", { checksum: removalHash,
          retiredRequirements: requirements.size, retiredValidations, retiredLayouts: layouts.length });
        this.put("meta", "schema", { version: 3 });
      }
      this.put("owner", "server", { pid: process.pid, token: this.token });
      this.db.exec("COMMIT");
    } catch (error) {
      this.db.exec("ROLLBACK");
      this.db.close();
      try {
        this.ownerDb.exec("ROLLBACK");
      } catch {
        // Releasing the SQLite connection below also releases the lease.
      }
      this.ownerDb.close();
      throw error;
    }
    for (const directory of ["snapshots", "candidates", "assets", "exports"])
      fs.mkdirSync(this.internal(directory), { recursive: true, mode: 0o700 });
  }

  internal(...parts) {
    return path.join(this.config.stateDirectory, ...parts);
  }

  put(kind, id, value) {
    this.db
      .prepare(
        "INSERT INTO records VALUES (?,?,?) ON CONFLICT(kind,id) DO UPDATE SET value=excluded.value",
      )
      .run(kind, id, JSON.stringify({ ...value, updatedAt: this.now() }));
  }

  get(kind, id, required = true) {
    const row = this.db
      .prepare("SELECT value FROM records WHERE kind=? AND id=?")
      .get(kind, id);
    requireThat(
      row || !required,
      "NOT_FOUND",
      "The requested handle or record is unavailable.",
    );
    if (!row) return null;
    const value = JSON.parse(row.value);
    if (
      kind === "operation" &&
      ["completed", "failed"].includes(value.status) &&
      this.now() - value.updatedAt >= this.config.operationRetentionMs
    ) {
      const expired = {
        status: "expired",
        fingerprint: value.fingerprint,
        name: value.name,
      };
      this.put(kind, id, expired);
      return expired;
    }
    return value;
  }

  rows(kind) {
    return this.db
      .prepare("SELECT id,value FROM records WHERE kind=?")
      .all(kind)
      .map((row) => ({ id: row.id, ...JSON.parse(row.value) }));
  }

  delete(kind, id) {
    this.db.prepare("DELETE FROM records WHERE kind=? AND id=?").run(kind, id);
  }

  killWorker(child) {
    if (!child.pid) return;
    if (process.platform === "win32")
      execFile("taskkill", ["/pid", String(child.pid), "/T", "/F"], () => {});
    else {
      try {
        process.kill(-child.pid, "SIGKILL");
      } catch {
        child.kill("SIGKILL");
      }
    }
  }

  checkQuota(extra = 0) {
    // Reclaim completed worst-case reservations before rejecting a fast sequence
    // of small jobs under a large per-file limit. Never drop in-flight budgets.
    if (this.quotaSnapshot.bytes + this.quotaSnapshot.reserved + extra > this.config.maxStateBytes &&
        this.activeWorkers.size === 0 && this.docxPool.activeRuns === 0)
      this.quotaSnapshot.checkedAt = Number.NEGATIVE_INFINITY;
    checkStateQuota(this.config, this.db, this.now, this.quotaSnapshot, extra);
  }

  close() {
    if (this.closed) return;
    this.closed = true;
    this.docxPool.close();
    this.documents?.closePublicationChecks();
    for (const job of this.jobs.values()) job.abort();
    for (const worker of this.activeWorkers) this.killWorker(worker);
    if (this.get("owner", "server", false)?.token === this.token)
      this.delete("owner", "server");
    this.db.close();
    try {
      this.ownerDb.exec("COMMIT");
    } catch {
      // Closing the connection is the authoritative lease release.
    }
    this.ownerDb.close();
  }
}
