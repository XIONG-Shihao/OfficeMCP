// Private OfficeMCP state under Bun. Bun does not expose Node's sqlite3_limit or
// defensive options. Instead expose a finite, parameter-bound SQL interface;
// reject foreign schemas/views/triggers rather than silently ignore hardening.
import { Database } from 'bun:sqlite'
const normalize = sql => sql.replace(/\s+/g, ' ').trim().replace(/;$/, '')
const declaration = sql => normalize(sql).replace(/\s*([(),])\s*/g, '$1').toLowerCase()
const schema = {
  lease: 'CREATE TABLE IF NOT EXISTS lease (singleton INTEGER PRIMARY KEY CHECK(singleton=1), pid INTEGER NOT NULL, token TEXT NOT NULL)',
  records: 'CREATE TABLE IF NOT EXISTS records (kind TEXT NOT NULL, id TEXT NOT NULL, value TEXT NOT NULL, PRIMARY KEY(kind,id))',
  objects: 'CREATE TABLE IF NOT EXISTS objects (key TEXT PRIMARY KEY, bytes BLOB NOT NULL, revision TEXT NOT NULL, operation_id TEXT UNIQUE NOT NULL)'
}
const statements = new Set([
  'INSERT OR REPLACE INTO lease VALUES (1, ?, ?)',
  'SELECT kind,id,value FROM records',
  'INSERT INTO records VALUES (?,?,?) ON CONFLICT(kind,id) DO UPDATE SET value=excluded.value',
  'SELECT value FROM records WHERE kind=? AND id=?',
  'SELECT id,value FROM records WHERE kind=?',
  'DELETE FROM records WHERE kind=? AND id=?',
  'SELECT kind,COUNT(*) AS count FROM records GROUP BY kind',
  "SELECT COALESCE(SUM(CASE WHEN json_valid(value)=0 THEN 1 WHEN json_extract(value,'$.status')='expired' THEN 0 ELSE 1 END),0) AS count FROM records WHERE kind='operation'",
  'SELECT bytes,revision FROM objects WHERE key=?',
  'SELECT key,revision FROM objects WHERE operation_id=?',
  'SELECT key FROM objects WHERE key=?',
  'SELECT COALESCE(SUM(length(bytes)),0) AS n FROM objects',
  'SELECT COUNT(*) AS n FROM objects',
  'INSERT INTO objects VALUES (?,?,?,?)'
])
const commands = new Set(['BEGIN IMMEDIATE', 'BEGIN EXCLUSIVE', 'COMMIT', 'ROLLBACK',
  'PRAGMA trusted_schema=OFF', 'PRAGMA journal_mode=WAL', 'PRAGMA journal_mode=DELETE',
  'PRAGMA synchronous=FULL', ...Object.values(schema)])
function refusal(message) { const error = Error(message); error.code = 'UNSUPPORTED_STATE'; return error }
export class DatabaseSync {
  #db
  #maxLength
  constructor(file, options) {
    if (!Number.isSafeInteger(options?.limits?.length) || options.limits.length <= 0 || options.allowExtension !== false) throw refusal('Invalid private database bounds')
    this.#maxLength = options.limits.length
    this.#db = new Database(file, {create: true, strict: true})
    try {
      this.#db.exec('PRAGMA trusted_schema=OFF; PRAGMA busy_timeout=2000; PRAGMA cell_size_check=ON;')
      // Every supported database has one of these app-owned schemas. Never
      // execute stored triggers or read an attacker-provided view as a table.
      for (const row of this.#db.query('SELECT type,name,tbl_name,sql FROM sqlite_schema').all()) {
        if (row.type === 'index' && row.name.startsWith('sqlite_autoindex_') && row.sql === null && schema[row.tbl_name]) continue
        // State written by earlier releases declares the same tables with
        // different spacing. Compare the declarations, not their layout.
        const expected = schema[row.name]?.replace('IF NOT EXISTS ', '')
        if (row.type !== 'table' || !expected || declaration(row.sql ?? '') !== declaration(expected)) throw refusal('Unknown private database schema')
      }
    } catch (error) { this.#db.close(true); throw error }
  }
  get isTransaction() { return this.#db.inTransaction }
  exec(sql) {
    const list = sql.split(';').map(normalize).filter(Boolean)
    if (!list.length || list.some(command => !commands.has(command))) throw refusal('Unsupported private database command')
    this.#db.exec(list.join(';'))
  }
  prepare(sql) {
    if (!statements.has(normalize(sql))) throw refusal('Unsupported private database statement')
    const query = this.#db.query(sql)
    const execute = (method, args) => {
      if (args.length > 32 || args.reduce((total, value) => total + (typeof value === 'string' ? Buffer.byteLength(value) : ArrayBuffer.isView(value) ? value.byteLength : 8), 0) > this.#maxLength) throw refusal('Private database binding limit exceeded')
      const result = query[method](...args)
      for (const row of method === 'all' ? result : method === 'get' && result ? [result] : []) {
        if (Object.values(row).reduce((total, value) => total + (typeof value === 'string' ? Buffer.byteLength(value) : ArrayBuffer.isView(value) ? value.byteLength : 8), 0) > this.#maxLength) throw refusal('Private database row limit exceeded')
      }
      return result
    }
    return {run: (...args) => execute('run', args), get: (...args) => execute('get', args), all: (...args) => execute('all', args)}
  }
  // Bun's close(false) can retain prepared statements and the exclusive lease.
  close() { this.#db.close(true) }
}
