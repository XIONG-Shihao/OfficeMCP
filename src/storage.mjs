/** Frozen optional embedding surface: maintenance fixes only. The integrated app uses the MCP CLI. See docs/SDK.md. */
import { DatabaseSync } from "./infrastructure/sqlite.mjs";
import { digest } from "./infrastructure/files.mjs";
import { requireThat } from "./infrastructure/errors.mjs";
import { sqliteOptions } from "./infrastructure/sqlite-options.mjs";
/** Optional persistent object store. Atomic bytes + receipt in one SQLite transaction. */
export class SQLiteObjectStore {
  constructor(
    file,
    {
      maxObjectBytes = 20 * 1024 * 1024,
      maxTotalBytes = 256 * 1024 * 1024,
      maxObjects = 10000,
    } = {},
  ) {
    this.maxObjects = maxObjects;
    this.maxObjectBytes = maxObjectBytes;
    this.maxTotalBytes = maxTotalBytes;
    this.db = new DatabaseSync(
      file,
      sqliteOptions(maxObjectBytes + 1024 * 1024),
    );
    this.db.exec(
      "PRAGMA trusted_schema=OFF; PRAGMA journal_mode=WAL; PRAGMA synchronous=FULL; CREATE TABLE IF NOT EXISTS objects (key TEXT PRIMARY KEY, bytes BLOB NOT NULL, revision TEXT NOT NULL, operation_id TEXT UNIQUE NOT NULL);",
    );
  }
  async read(key) {
    const row = this.db
      .prepare("SELECT bytes,revision FROM objects WHERE key=?")
      .get(key);
    requireThat(row, "NOT_FOUND", "Storage key unavailable.");
    return { bytes: Buffer.from(row.bytes), revision: row.revision };
  }
  async putIfAbsent({ key, bytes, operationId }) {
    requireThat(
      typeof key === "string" &&
        key.length > 0 &&
        key.length <= 4096 &&
        typeof operationId === "string" &&
        operationId.length > 0 &&
        operationId.length <= 150 &&
        Buffer.isBuffer(bytes) &&
        bytes.length <= this.maxObjectBytes,
      "INVALID_ARGUMENT",
      "Invalid bounded storage publication.",
    );
    const revision = digest(bytes);
    this.db.exec("BEGIN IMMEDIATE");
    try {
      const previous = this.db
        .prepare("SELECT key,revision FROM objects WHERE operation_id=?")
        .get(operationId);
      if (previous) {
        requireThat(
          previous.key === key && previous.revision === revision,
          "OPERATION_ID_CONFLICT",
          "Storage operation ID conflict.",
        );
        this.db.exec("COMMIT");
        return { key, revision, operationId, published: true };
      }
      requireThat(
        !this.db.prepare("SELECT key FROM objects WHERE key=?").get(key),
        "DESTINATION_CONFLICT",
        "Storage key already exists.",
      );
      requireThat(
        this.db
          .prepare("SELECT COALESCE(SUM(length(bytes)),0) AS n FROM objects")
          .get().n +
          bytes.length <=
          this.maxTotalBytes,
        "LIMIT_EXCEEDED",
        "Storage quota reached.",
      );
      requireThat(
        this.db.prepare("SELECT COUNT(*) AS n FROM objects").get().n <
          this.maxObjects,
        "LIMIT_EXCEEDED",
        "Storage object count reached.",
      );
      this.db
        .prepare("INSERT INTO objects VALUES (?,?,?,?)")
        .run(key, bytes, revision, operationId);
      this.db.exec("COMMIT");
      return { key, revision, operationId, published: true };
    } catch (e) {
      if (this.db.isTransaction) this.db.exec("ROLLBACK");
      throw e;
    }
  }
  close() {
    this.db.close();
  }
}
