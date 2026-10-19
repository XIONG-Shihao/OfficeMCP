import * as fs from "node:fs";
import path from "node:path";
import { requireThat } from "./errors.mjs";

export function createQuotaSnapshot() {
  return { checkedAt: Number.NEGATIVE_INFINITY, bytes: 0, reserved: 0 };
}

function treeBytes(directory) {
  let bytes = 0;
  for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
    const file = path.join(directory, entry.name);
    bytes += entry.isDirectory() ? treeBytes(file) : fs.statSync(file).size;
  }
  return bytes;
}

export function checkStateQuota(config, db, now, snapshot, extra = 0) {
  const checkedAt = now();
  if (checkedAt - snapshot.checkedAt >= 1000) {
    snapshot.checkedAt = checkedAt;
    snapshot.bytes = treeBytes(config.stateDirectory);
    snapshot.reserved = 0;
  }
  requireThat(
    snapshot.bytes + snapshot.reserved + extra <= config.maxStateBytes,
    "LIMIT_EXCEEDED",
    "Private state quota reached. Close/discard unused documents or use a new state directory.",
  );
  // Reserve the declared worst case until the next real scan. Bursty batches
  // avoid repeated tree walks without weakening the quota.
  snapshot.reserved += extra;
  const operationCount = db
    .prepare(
      `SELECT COALESCE(SUM(CASE
        WHEN json_valid(value)=0 THEN 1
        WHEN json_extract(value,'$.status')='expired' THEN 0
        ELSE 1 END),0) AS count
       FROM records WHERE kind='operation'`,
    )
    .get().count;
  requireThat(
    operationCount < config.maxActiveOperations,
    "LIMIT_EXCEEDED",
    "Active operation receipt limit reached; run office_cleanup after retention elapses or use a new state directory.",
  );
}
