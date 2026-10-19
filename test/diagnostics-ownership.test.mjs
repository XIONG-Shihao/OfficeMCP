import test from "./support/test.mjs";
import assert from "node:assert/strict";
import * as fs from "node:fs";
import path from "node:path";
import { tmpdir } from "node:os";
import { DatabaseSync } from "./support/raw-sqlite.mjs";
import { OfficeService } from "../src/application/service.mjs";
import { configFrom } from "../src/infrastructure/files.mjs";
import { safeEngineFailure } from "../src/infrastructure/diagnostics.mjs";

function fixture() {
  const root = fs.mkdtempSync(path.join(tmpdir(), "officemcp-diagnostics-"));
  const docs = path.join(root, "docs");
  fs.mkdirSync(docs);
  const configPath = path.join(root, "config.json");
  fs.writeFileSync(configPath, JSON.stringify({
    readRoots: [docs],
    writeRoots: { output: docs },
    stateDirectory: path.join(root, "state"),
  }));
  return { root, config: configFrom(configPath) };
}

test("state ownership uses an OS-backed SQLite lock, not PID existence", () => {
  const f = fixture();
  const first = new OfficeService(f.config);
  try {
    assert.throws(() => new OfficeService(f.config), { code: "STATE_IN_USE" });
  } finally {
    first.close();
  }
  const db = new DatabaseSync(path.join(f.config.stateDirectory, "state.sqlite"));
  db.prepare("INSERT OR REPLACE INTO records VALUES ('owner','server',?)")
    .run(JSON.stringify({ pid: process.pid, token: "stale-pid-reuse" }));
  db.close();
  const replacement = new OfficeService(f.config);
  replacement.close();
  fs.rmSync(f.root, { recursive: true, force: true });
});

test("unexpected failures receive private redacted diagnostics", async () => {
  const f = fixture();
  const service = new OfficeService(f.config);
  try {
    service.execute = async () => { throw new Error("failed at /Users/alice/private/offer.docx") };
    const result = await service.call("office_engine_status", {});
    assert.equal(result.payload.error.code, "INTERNAL_ERROR");
    assert.match(result.payload.error.diagnosticId, /^diag_/);
    assert.match(result.payload.error.message, new RegExp(result.payload.error.diagnosticId));
    const log = fs.readFileSync(path.join(f.config.stateDirectory, "diagnostics", "officemcp.jsonl"), "utf8");
    assert.match(log, new RegExp(result.payload.error.diagnosticId));
    assert.doesNotMatch(log, /alice|offer\.docx/);
    assert.deepEqual(safeEngineFailure({ code: "EIO", message: "/Users/alice/private.docx" }), {
      code: "ENGINE_FAILED",
      message: "SuperDoc failed (EIO). Inspect the coordinator diagnostic record.",
    });
  } finally {
    service.close();
    fs.rmSync(f.root, { recursive: true, force: true });
  }
});

test("quota checks use maintained reservations and SQL counts without parsing operation rows", () => {
  const f = fixture();
  const service = new OfficeService(f.config, { now: () => 1000 });
  try {
    // Plant a malformed row through the product's own bound statement.
    service.db.prepare("INSERT INTO records VALUES (?,?,?) ON CONFLICT(kind,id) DO UPDATE SET value=excluded.value").run('operation','broken-test-row','not-json');
    service.checkQuota(10);
    const checkedAt = service.quotaSnapshot.checkedAt;
    service.checkQuota(20);
    assert.equal(service.quotaSnapshot.checkedAt, checkedAt);
    assert.equal(service.quotaSnapshot.reserved, 30);
  } finally {
    service.db.prepare("DELETE FROM records WHERE kind=? AND id=?").run('operation','broken-test-row');
    service.close();
    fs.rmSync(f.root, { recursive: true, force: true });
  }
});

test("expired operation tombstones do not consume active receipt capacity", () => {
  const f = fixture();
  const service = new OfficeService(
    { ...f.config, maxActiveOperations: 2 },
    { now: () => 1000 },
  );
  try {
    service.put("operation", "expired-1", { status: "expired", fingerprint: "one" });
    service.put("operation", "expired-2", { status: "expired", fingerprint: "two" });
    service.put("operation", "active-1", { status: "completed", fingerprint: "three" });
    service.checkQuota();
    service.put("operation", "active-2", { status: "failed", fingerprint: "four" });
    assert.throws(() => service.checkQuota(), { code: "LIMIT_EXCEEDED" });
    assert.equal(service.get("operation", "expired-1").status, "expired");
  } finally {
    service.close();
    fs.rmSync(f.root, { recursive: true, force: true });
  }
});
