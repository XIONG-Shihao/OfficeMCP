import test from './support/test.mjs';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { productFixture } from '../scripts/test/docx-product-fixture.mjs';
const oracle = fileURLToPath(new URL('../scripts/test/docx_corpus_evidence.py', import.meta.url));
test('independent corpus oracle rejects lost formatting, comment text/anchors and media bytes', async () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'docx-corpus-oracle-'));
  try {
    const original = await productFixture(), source = path.join(root, 'input.docx'), spec = path.join(root, 'copy.json');
    fs.writeFileSync(source, original); fs.writeFileSync(spec, JSON.stringify({ kind: 'copy' }));
    const verify = bytes => {
      const out = path.join(root, 'out.docx'); fs.writeFileSync(out, bytes);
      return JSON.parse(execFileSync(process.env.OFFICEMCP_PYTHON ?? 'python3', [oracle, 'verify', source, out, spec], { encoding: 'utf8' }));
    };
    assert.equal(verify(original).passed, true);
    const mutateXml = (file, transform) => {
      const parts = unzipSync(original), before = strFromU8(parts[file]), after = transform(before);
      assert.notEqual(after, before, 'Adversarial mutation did not change the fixture'); parts[file] = strToU8(after);
      return verify(zipSync(parts));
    };
    assert.equal(mutateXml('word/document.xml', xml => xml.replace(/<w:b\b[^>]*\/>/, '')).passed, false);
    assert.equal(mutateXml('word/comments.xml', xml => xml.replace('Approved by legal', 'Forged review note')).passed, false);
    assert.equal(mutateXml('word/document.xml', xml => xml.replace(/<w:commentRangeStart\b[^>]*\/>/, '')).passed, false);
    const parts = unzipSync(original), image = Object.keys(parts).find(n => n.startsWith('word/media/') && !n.endsWith('/') && parts[n].length > 0);
    assert(image); parts[image][0] ^= 1; const media = verify(zipSync(parts));
    assert.equal(media.passed, false); assert(media.mismatches.includes('binaries'));
  } finally { fs.rmSync(root, { recursive: true, force: true }); }
});
