// One private-state database interface for both supported runtimes. Node has a
// built-in synchronous SQLite; Bun has its own, wrapped to the same bounded
// shape (see sqlite-bun.mjs).
export const DatabaseSync = process.versions.bun
  ? (await import("./sqlite-bun.mjs")).DatabaseSync
  : (await import("node:sqlite")).DatabaseSync;
export const sqliteEngine = process.versions.bun
  ? "bun:sqlite (bounded private-state adapter)"
  : "node:sqlite";
