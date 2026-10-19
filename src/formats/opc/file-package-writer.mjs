/** Copy untouched compressed members; inflate only explicitly changed parts. */
import fs from 'node:fs';
import { zipSync } from 'fflate';
import { openFilePackage } from './file-package.mjs';
import { requireThat } from '../../infrastructure/errors.mjs';
import { fileRevision } from '../../infrastructure/files.mjs';
export function rewriteFilePackage(input, output, replacements, { maxExpandedBytes, maxFileBytes } = {}) {
  requireThat(!fs.existsSync(output), 'DESTINATION_CONFLICT', 'Private package output exists.');
  const archive = openFilePackage(input, maxExpandedBytes, maxFileBytes);
  const source = fs.openSync(input, fs.constants.O_RDONLY | (fs.constants.O_NOFOLLOW ?? 0));
  let target;
  try {
    const entries = archive.entries(), names = new Set(entries.map(entry => entry.name));
    requireThat([...replacements.keys()].every(name => names.has(name)), 'UNSUPPORTED_FEATURE', 'This adapter replaces existing members only.');
    target = fs.openSync(output, fs.constants.O_WRONLY | fs.constants.O_CREAT | fs.constants.O_EXCL, 0o600);
    const directory = [], chunk = Buffer.alloc(1024 * 1024); let at = 0, expanded = 0;
    const write = bytes => { let offset = 0; while (offset < bytes.length) offset += fs.writeSync(target, bytes, offset, bytes.length - offset); at += bytes.length; };
    for (const entry of entries) {
      const offset = at, central = Buffer.from(entry.central); let header, payload, compressed = entry.compressed, size = entry.expanded, crc = entry.crc;
      if (replacements.has(entry.name)) {
        const bytes = replacements.get(entry.name), zip = Buffer.from(zipSync({ [entry.name]: bytes }, { level: 6, mtime: new Date(1980, 0, 1) }));
        const end = zip.length - 22, cd = zip.readUInt32LE(end + 16); crc = zip.readUInt32LE(cd + 16); compressed = zip.readUInt32LE(cd + 20); size = zip.readUInt32LE(cd + 24);
        const data = 30 + zip.readUInt16LE(26) + zip.readUInt16LE(28); header = Buffer.from(zip.subarray(0, data)); payload = zip.subarray(data, data + compressed);
        // Keep the original central metadata but use the new member's method.
        central.writeUInt16LE(zip.readUInt16LE(cd + 10), 10);
      } else {
        header = Buffer.alloc(entry.data - entry.local); fs.readSync(source, header, 0, header.length, entry.local);
      }
      header.writeUInt16LE(header.readUInt16LE(6) & ~8, 6); header.writeUInt32LE(crc, 14); header.writeUInt32LE(compressed, 18); header.writeUInt32LE(size, 22);
      central.writeUInt16LE(central.readUInt16LE(8) & ~8, 8); central.writeUInt32LE(crc, 16); central.writeUInt32LE(compressed, 20); central.writeUInt32LE(size, 24); central.writeUInt32LE(offset, 42);
      write(header);
      if (payload) write(payload);
      else for (let copied = 0; copied < compressed;) {
        const n = fs.readSync(source, chunk, 0, Math.min(chunk.length, compressed - copied), entry.data + copied);
        requireThat(n > 0, 'INTEGRITY_FAILED', 'Source package changed while copying.'); write(chunk.subarray(0, n)); copied += n;
      }
      expanded += size; directory.push(central);
      requireThat(expanded <= maxExpandedBytes && at <= maxFileBytes, 'LIMIT_EXCEEDED', 'Changed package exceeds configured limits.');
    }
    const start = at; for (const header of directory) write(header);
    const end = Buffer.alloc(22); end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(entries.length, 8); end.writeUInt16LE(entries.length, 10); end.writeUInt32LE(at - start, 12); end.writeUInt32LE(start, 16); write(end);
    requireThat(at <= maxFileBytes, 'LIMIT_EXCEEDED', 'Changed package exceeds the file budget.');
    fs.fsyncSync(target); archive.assertUnchanged(); fs.closeSync(target); target = undefined;
    return { revision: fileRevision(output, maxFileBytes), changedParts: [...replacements.keys()] };
  } catch (error) { if (target !== undefined) fs.closeSync(target); fs.rmSync(output, { force: true }); throw error; }
  finally { fs.closeSync(source); archive.close(); }
}
