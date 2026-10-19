import test from './support/test.mjs';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { Document, Packer, Paragraph, Table, TableRow, TableCell, TextRun, HeadingLevel } from 'docx';
import { unzipSync, strFromU8 } from 'fflate';
import { createServiceFixture } from './service-fixture.mjs';
import { inspectPackage } from '../scripts/test/docx-product-evidence.mjs';
import { DOMParser } from '@xmldom/xmldom';
import { commentedDocument } from './comment-fixture.mjs';
const cell = (text, properties = {}) => new TableCell({ children: [new Paragraph({ children: [new TextRun({ text, bold: true })] })], ...properties });
const row = values => new TableRow({ children: values.map(value => cell(value)) });
const workflow = { type: 'update_table_by_key', keyColumn: 'Service', updates: [{ key: 'Support', column: 'Price USD', expectedText: '125', text: '150' }] };
test('keyed edit preserves an unrelated merged summary row and formatted chosen cell', async () => {
  const f = createServiceFixture({ toolset: 'docx-agent' });
  try {
    const source = path.join(f.docs, 'merged.docx');
    fs.writeFileSync(source, await Packer.toBuffer(new Document({ sections: [{ children: [new Table({ rows: [row(['Service', 'Price USD']), row(['Support', '125']), row(['Training', '300']), new TableRow({ children: [cell('Summary remains', { columnSpan: 2 })] })] }), new Paragraph('Outside remains')] }] })));
    const original = fs.readFileSync(source), view = (await f.call('docx_open', { path: source })).view;
    const edited = await f.call('docx_apply', { view, operationId: 'keyed', select: { type: 'table', header: 'Price USD' }, workflow, readback: {} });
    assert.equal(edited.data.workflowVerification.method, 'guarded-cell-text');
    const saved = await f.call('docx_save', { view: edited.view, operationId: 'save', destination: { rootId: 'output', relativePath: 'saved.docx' } });
    const xml = strFromU8(unzipSync(fs.readFileSync(saved.data.path))['word/document.xml']);
    assert(inspectPackage(fs.readFileSync(saved.data.path)).tables.some(rows => rows.some(row => row.includes('150')))); assert(xml.includes('Summary remains')); assert(xml.includes('Outside remains')); assert(xml.includes('gridSpan'));
    const tree = new DOMParser().parseFromString(xml, 'application/xml'), W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const price = Array.from(tree.getElementsByTagNameNS(W, 'p')).find(p => p.textContent === '150'); assert(price);
    assert(Array.from(price.getElementsByTagNameNS(W, 'r')).filter(r => r.getElementsByTagNameNS(W, 't').length).every(r => r.getElementsByTagNameNS(W, 'b').length > 0)); assert.deepEqual(fs.readFileSync(source), original);
  } finally { f.close(); }
});
test('section replacement preserves existing comments outside its selected range', async () => {
  const f = createServiceFixture({ toolset: 'docx-agent' });
  try {
    const source = path.join(f.docs, 'comment-outside.docx');
    fs.writeFileSync(source, await commentedDocument({ anchor: 'Untouched reviewer paragraph.', comment: 'Keep reviewer context.', after: [new Paragraph({ text: 'Methodology', heading: HeadingLevel.HEADING_1 }), new Paragraph('Old method.'), new Paragraph({ text: 'Conclusion', heading: HeadingLevel.HEADING_1 }), new Paragraph('Unchanged ending.')] }));
    const original = fs.readFileSync(source), view = (await f.call('docx_open', { path: source })).view;
    const edited = await f.call('docx_apply', { view, operationId: 'section', workflow: { type: 'replace_section', headingPath: ['Methodology'], subsections: 'preserve', expectedText: 'Old method.', paragraphs: ['Updated method.'] } });
    const saved = await f.call('docx_save', { view: edited.view, operationId: 'save', destination: { rootId: 'output', relativePath: 'saved.docx' } });
    const reopened = (await f.call('docx_open', { path: saved.data.path })).view;
    const comments = await f.call('docx_read', { view: reopened, representation: 'domain', domain: 'comments' });
    assert(comments.data.items.some(c => c.text === 'Keep reviewer context.' && c.status === 'open'));
    assert.deepEqual(fs.readFileSync(source), original);
    const before = unzipSync(original), after = unzipSync(fs.readFileSync(saved.data.path));
    assert.equal(strFromU8(after['word/comments.xml']), strFromU8(before['word/comments.xml']));
  } finally { f.close(); }
});
