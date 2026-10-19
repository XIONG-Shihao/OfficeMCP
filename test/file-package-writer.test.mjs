import test from './support/test.mjs';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { zipSync, unzipSync, strToU8 } from 'fflate';
import { rewriteFilePackage } from '../src/formats/opc/file-package-writer.mjs';
const limits = { maxFileBytes: 2 * 1024 ** 2, maxExpandedBytes: 3 * 1024 ** 2 };
test('file package writer changes only requested existing members and keeps binary content exact', () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'opc-writer-'));
  try {
    const parts = { '[Content_Types].xml': strToU8('<Types/>'), 'word/document.xml': strToU8('<doc>Old</doc>'), 'word/media/image.png': Uint8Array.from({ length: 100000 }, (_, i) => i % 256), 'word/中文.xml': strToU8('<data>保留</data>') };
    const input = path.join(root, 'in.docx'), output = path.join(root, 'out.docx'); fs.writeFileSync(input, zipSync(parts));
    const before = fs.readFileSync(input), changed = strToU8('<doc>Updated</doc>');
    rewriteFilePackage(input, output, new Map([['word/document.xml', changed]]), limits);
    assert.deepEqual(unzipSync(fs.readFileSync(output)), { ...parts, 'word/document.xml': changed });
    assert.deepEqual(fs.readFileSync(input), before);
    assert.throws(() => rewriteFilePackage(input, output, new Map(), limits), /exists/);
    assert.throws(() => rewriteFilePackage(input, path.join(root, 'unknown.docx'), new Map([['new.xml', changed]]), limits), /existing members/);
    assert(!fs.existsSync(path.join(root, 'unknown.docx')));
    assert.throws(() => rewriteFilePackage(input, path.join(root, 'large.docx'), new Map([['word/document.xml', new Uint8Array(4 * 1024 ** 2)]]), limits), /limits/);
    assert(!fs.existsSync(path.join(root, 'large.docx')));
  } finally { fs.rmSync(root, { recursive: true, force: true }); }
});
