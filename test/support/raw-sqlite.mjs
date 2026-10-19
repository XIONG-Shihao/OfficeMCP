// Unrestricted database access for tests that plant or inspect historical
// state. Product code uses src/infrastructure/sqlite.mjs.
let DatabaseSync;
if (process.versions.bun) {
  const { Database } = await import("bun:sqlite");
  DatabaseSync = class extends Database {
    constructor(file, options = {}) {
      super(file, options.readOnly ? { readonly: true } : { create: true, readwrite: true });
    }
  };
} else ({ DatabaseSync } = await import("node:sqlite"));
export { DatabaseSync };
