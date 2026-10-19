import test from './support/test.mjs';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { PDFDocument, StandardFonts } from 'pdf-lib';
import { createServiceFixture } from './service-fixture.mjs';
import { inspectLayoutPdf, compareLayoutPixels } from '../src/formats/docx/docx-layout-provider.mjs';
test('appearance comparison detects changed text and missing pages in independent PDF artifacts', async () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'office-appearance-'));
  try {
    const make = async (file, text, pages = 1) => {
      const doc = await PDFDocument.create(), font = await doc.embedFont(StandardFonts.Helvetica);
      for (let i = 0; i < pages; i++) doc.addPage([300, 400]).drawText(text, { x: 30, y: 350, font, size: 12 });
      fs.writeFileSync(file, await doc.save());
    };
    const a = path.join(root, 'a.pdf'), b = path.join(root, 'b.pdf');
    await make(a, 'Original clause.'); await make(b, 'Original clause.');
    assert.equal(compareLayoutPixels(a, b, { maxBytes: 1000000 }).identical, true);
    await make(b, 'Changed clause.'); assert.equal(compareLayoutPixels(a, b, { maxBytes: 1000000 }).identical, false);
    await make(b, 'Original clause.', 2); assert.equal(compareLayoutPixels(a, b, { maxBytes: 1000000 }).pages.at(-1).reason, 'missing-page');
    assert.equal(inspectLayoutPdf(a, 1000000).pageCount, 1);
  } finally { fs.rmSync(root, { recursive: true, force: true }); }
});
test('layout unavailable is explicit and does not alter a document', async () => {
  const f = createServiceFixture({ toolset: 'docx-agent' });
  try {
    const created = await f.call('docx_create', { operationId: 'new', document: { format: 'docx', blocks: [{ kind: 'paragraph', text: 'Original.' }] } });
    const result = await f.service.call('docx_layout', { view: created.view });
    assert(result.isError); assert.equal(result.payload.error.code, 'ENGINE_UNAVAILABLE');
  } finally { f.close(); }
});
const libreOffice = process.env.OFFICEMCP_LIBREOFFICE;
test('page query resolves an actual SDK target and rejects stale or ambiguous page text', { skip: !libreOffice }, async () => {
  const f = createServiceFixture({ toolset: 'docx-agent', engines: { libreOffice }, limits: { workerTimeoutMs: 120000 } });
  try {
    let view = (await f.call('docx_create', { operationId: 'new', document: { format: 'docx', blocks: [{ kind: 'paragraph', text: 'Unique payment deadline is 30 days.' }, { kind: 'paragraph', text: 'Duplicate phrase.' }] } })).view;
    const layout = await f.call('docx_layout', { view, provider: 'libreoffice' });
    assert.equal(layout.data.pageCount, 1);
    const select = { type: 'page', number: 1, layoutId: layout.data.layoutId, pattern: 'Unique payment deadline is 30 days.' };
    const query = await f.call('docx_query', { view, select, require: 'exactlyOne' }); assert.equal(query.data.items.length, 1);
    view = (await f.call('docx_apply', { view, operationId: 'edit', action: 'replace_text', selectorTarget: query.data.items[0].target, arguments: { edits: [{ find: '30', replace: '45' }], caseSensitive: true, changeMode: 'direct' } })).view;
    const stale = await f.service.call('docx_query', { view, select }); assert(stale.isError); assert.equal(stale.payload.error.code, 'TARGET_CONFLICT');
    const saved = await f.call('docx_save', { view, operationId: 'save', destination: { rootId: 'output', relativePath: 'page-edited.docx' } });
    const reopened = (await f.call('docx_open', { path: saved.data.path })).view;
    const text = await f.call('docx_query', { view: reopened, select: { type: 'text', pattern: '45 days.' }, require: 'exactlyOne' }); assert.equal(text.data.items.length, 1);
  } finally { f.close(); }
});
