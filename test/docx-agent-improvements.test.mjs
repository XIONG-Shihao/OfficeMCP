import test from './support/test.mjs';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { createServiceFixture } from './service-fixture.mjs';
import { docxAgentTools } from '../src/formats/docx/docx-agent-contract.mjs';
import { digest } from '../src/infrastructure/files.mjs';
import { unzipSync, strFromU8 } from 'fflate';
import { Paragraph, HeadingLevel } from 'docx';
import { inspectPackage } from '../scripts/test/docx-product-evidence.mjs';
import { commentedDocument } from './comment-fixture.mjs';

const setup = () => createServiceFixture({ toolset: 'docx-agent', limits: { workerTimeoutMs: 30000 }, contractVersion: 'docx/6' });
const paragraph = text => ({ kind: 'paragraph', text });
const heading = (text, level = 1) => ({ kind: 'heading', level, text });
const priceTable = { kind: 'table', rows: [['Service', 'Price USD'], ['Support', '125'], ['Training', '300']] };
const create = async (f, blocks, id = 'create') => (await f.call('docx_create', { operationId: id, document: { format: 'docx', blocks } })).view;
const text = async (f, view) => (await f.call('docx_read', { view, representation: 'text', limit: 100 })).data.items.map(i => i.text).join('\n');
const keyEdit = { type: 'update_table_by_key', keyColumn: 'Service', updates: [{ key: 'Support', column: 'Price USD', expectedText: '125', text: '150' }] };
const sectionEdit = { type: 'replace_section', headingPath: ['Methodology'], subsections: 'preserve', expectedText: 'Old method.', paragraphs: ['New method.', 'Delivery after approval.'] };
async function refused(f, args, code) {
  const before = f.service.rows('candidate').length;
  const response = await f.service.call('docx_apply', args);
  assert(response.isError, JSON.stringify(response.payload));
  assert.equal(response.payload.error.code, code, JSON.stringify(response.payload));
  assert.equal(f.service.rows('candidate').length, before, 'refusal must not create a candidate');
}

test('capability help needs no handle/worker and includes schema-valid workflow examples', async () => {
  const f = setup();
  try {
    f.service.worker = () => { throw Error('Help must not start an engine'); };
    const help = await f.call('docx_inspect', { actions: ['replace_text'], workflows: ['update_table_by_key', 'replace_section', 'fill_template', 'update_table', 'normalize_document_style'] });
    assert.equal(help.data.documentInspected, false);
    assert.equal(help.view, undefined);
    assert.equal(f.service.rows('document').length, 0);
    assert.equal(Object.keys(help.data.actionDescriptors).length, 1);
    assert.equal(help.data.availableWorkflows.length, 5);
    assert.equal(help.data.workflowLimits.update_table_by_key.unrelatedMerges, "preserved");
    assert.equal(help.data.workflowLimits.replace_section.outsideReviewContent, "preserved");
    assert.equal(help.data.workflowLimits.page.staleAfterEdit, true);
    for (const descriptor of Object.values(help.data.workflowDescriptors)) docxAgentTools.docx_apply.input.parse(descriptor.example.arguments);
    const invalid = await f.service.call('docx_inspect', { includeDomains: ['blocks'] });
    assert(invalid.isError);
  } finally { f.close(); }
});

test('bounded readback survives preview, replay, receipt retrieval, worker reload and fresh-target reuse', async () => {
  const f = setup();
  try {
    let view = await create(f, [paragraph('Payment within 30 days.'), paragraph('Keep me.')]);
    const request = { view, operationId: 'payment', action: 'replace_text', select: { type: 'text', pattern: '30 days' }, arguments: { edits: [{ find: '30 days', replace: '45 days' }] }, readback: { maxItems: 3, maxChars: 100 } };
    const preview = await f.call('docx_apply', { ...request, mode: 'preview' });
    assert.deepEqual(preview.view, view);
    assert.equal(preview.data.readback.reusableTargets, false);
    assert(preview.data.readback.items.every(i => !i.target && !i.nativeTarget));
    assert((await text(f, view)).includes('30 days'));
    const applied = await f.call('docx_apply', request);
    const changed = applied.data.readback.items.find(i => i.after?.text.includes('45 days'));
    assert(changed.target && changed.before.text.includes('30 days'));
    assert(!JSON.stringify(applied.data.readback).includes('nativeTarget'));
    assert.deepEqual((await f.call('docx_apply', request)).data.readback, applied.data.readback);
    assert.deepEqual((await f.call('docx_operation_get', { operationId: request.operationId })).data.readback, applied.data.readback);
    view = applied.view;
    await f.service.docxPool.evictIdle(view.documentId);
    const formatted = await f.call('docx_apply', { view, operationId: 'format-readback-target', action: 'format_text', selectorTarget: changed.target, arguments: { bold: true }, readback: {} });
    assert(formatted.data.readback.notChecked.includes('formatting'));
    const formattedXml = strFromU8(unzipSync(fs.readFileSync(f.service.get('candidate', formatted.view.candidateId).path))['word/document.xml']);
    assert(/<w:b(?:\s|\/|>)/.test(formattedXml), 'fresh readback target must format the saved candidate after worker reload');
    assert((await text(f, formatted.view)).includes('Keep me.'));
    await refused(f, { view: formatted.view, operationId: 'stale-target', action: 'format_text', selectorTarget: changed.target, arguments: { italic: true } }, 'TARGET_CONFLICT');
    const long = await f.call('docx_apply', { view: formatted.view, operationId: 'long-readback', action: 'insert_paragraphs', arguments: { texts: ['x'.repeat(500), 'y'.repeat(500)], placement: { at: 'document_end' } }, readback: { maxItems: 1, maxChars: 100 } });
    assert.equal(long.data.readback.items.length, 1);
    assert.equal(long.data.readback.textChars, 100);
    assert.equal(long.data.readback.truncated, true);
    assert(long.data.readback.total >= 2);
    await refused(f, { view: long.view, operationId: 'rollback', steps: [
      { action: 'replace_text', arguments: { edits: [{ find: 'Keep me.', replace: 'Broken' }] } },
      { action: 'write_table_cells', select: { type: 'table', header: 'Missing table' }, arguments: { cells: [{ row: 0, column: 0, expectedText: 'missing', text: 'changed' }] } },
    ], readback: {} }, 'TARGET_CONFLICT');
    assert((await text(f, long.view)).includes('Keep me.'));
  } finally { f.close(); }
});

test('keyed workflow resolves semantic keys, preserves content, saves and provides a fresh table target', async () => {
  const f = setup();
  try {
    const sourceView = await create(f, [heading('Pricing'), priceTable, paragraph('Untouched sentinel.')]);
    const source = await f.call('docx_save', { view: sourceView, operationId: 'source-save', destination: { rootId: 'output', relativePath: 'source.docx' } });
    const hash = digest(fs.readFileSync(source.data.path));
    let view = (await f.call('docx_open', { path: source.data.path })).view;
    const applied = await f.call('docx_apply', { view, operationId: 'keyed', select: { type: 'table', header: 'Price USD' }, workflow: keyEdit, readback: {} });
    assert.equal(applied.data.workflowVerification.changedCells, 1);
    const change = applied.data.readback.items.find(i => i.kind === 'tableCell');
    assert.equal(change.before.text, '125'); assert.equal(change.after.text, '150');
    assert.equal(change.row, 1); assert.equal(change.column, 1); assert.equal(change.targetKind, 'table');
    view = applied.view;
    await f.service.docxPool.evictIdle(view.documentId);
    const next = await f.call('docx_apply', { view, operationId: 'keyed-followup', selectorTarget: change.target, workflow: { ...keyEdit, updates: [{ key: 'Training', column: 'Price USD', expectedText: '300', text: '350' }] }, readback: {} });
    const saved = await f.call('docx_save', { view: next.view, operationId: 'save', destination: { rootId: 'output', relativePath: 'updated.docx' } });
    const xml = strFromU8(unzipSync(fs.readFileSync(saved.data.path))['word/document.xml']);
    const inspected = inspectPackage(fs.readFileSync(saved.data.path));
    assert(inspected.tables.some(rows => rows.some(row => row.includes('150')) && rows.some(row => row.includes('350'))));
    assert(xml.includes('Untouched sentinel.'));
    assert.equal(digest(fs.readFileSync(source.data.path)), hash);
  } finally { f.close(); }
});

test('keyed workflow refuses ambiguous headers/keys, missing keys, stale values and duplicate writes atomically', async () => {
  const f = setup();
  try {
    const view = await create(f, [priceTable]);
    for (const [id, updates, code] of [
      ['missing-key', [{ key: 'Missing', column: 'Price USD', expectedText: '125', text: '150' }], 'TARGET_CONFLICT'],
      ['missing-column', [{ key: 'Support', column: 'No column', expectedText: '125', text: '150' }], 'TARGET_CONFLICT'],
      ['stale-value', [{ key: 'Support', column: 'Price USD', expectedText: '124', text: '150' }], 'TARGET_CONFLICT'],
      ['later-stale', [keyEdit.updates[0], { key: 'Training', column: 'Price USD', expectedText: 'wrong', text: '350' }], 'TARGET_CONFLICT'],
      ['duplicate-write', [keyEdit.updates[0], keyEdit.updates[0]], 'INVALID_ARGUMENT'],
    ]) await refused(f, { view, operationId: id, select: { type: 'table', header: 'Price USD' }, workflow: { ...keyEdit, updates } }, code);
    assert((await text(f, view)).includes('125'));
    for (const [id, rows] of [
      ['duplicate-header', [['Service', 'Price USD', 'Price USD'], ['Support', '125', '200']]],
      ['duplicate-key', [['Service', 'Price USD'], ['Support', '125'], ['Support', '300']]],
    ]) {
      const duplicate = await create(f, [{ kind: 'table', rows }], id);
      await refused(f, { view: duplicate, operationId: id + '-edit', select: { type: 'table', header: 'Price USD' }, workflow: keyEdit }, 'TARGET_CONFLICT');
    }
  } finally { f.close(); }
});

test('section workflow preserves or replaces subsections explicitly and verifies boundaries through save/reopen', async () => {
  const f = setup();
  try {
    let view = await create(f, [paragraph('Prefix.'), heading('Methodology'), paragraph('Old method.'), heading('Details', 2), paragraph('Keep details.'), heading('Pricing'), priceTable, paragraph('Suffix.')]);
    const preview = await f.call('docx_apply', { view, mode: 'preview', workflow: sectionEdit, readback: {} });
    assert(preview.data.workflowVerification.passed);
    assert((await text(f, view)).includes('Old method.'));
    const applied = await f.call('docx_apply', { view, operationId: 'section-preserve', workflow: sectionEdit, readback: {} });
    const after = await text(f, applied.view);
    assert(after.includes('New method.') && after.includes('Keep details.') && after.includes('Suffix.'));
    assert(!after.includes('Old method.'));
    const item = applied.data.readback.items.find(i => i.change === 'inserted'); assert(item.target);
    await f.service.docxPool.evictIdle(applied.view.documentId);
    const formatted = await f.call('docx_apply', { view: applied.view, operationId: 'section-target-reuse', action: 'format_text', selectorTarget: item.target, arguments: { italic: true } });
    const replaced = await f.call('docx_apply', { view: formatted.view, operationId: 'section-replace-subsections', workflow: { ...sectionEdit, subsections: 'replace', expectedText: 'New method.\nDelivery after approval.\nDetails\nKeep details.', paragraphs: ['Complete replacement.'] } });
    assert.equal(replaced.data.workflowVerification.removedBlocks, 4);
    const saved = await f.call('docx_save', { view: replaced.view, operationId: 'section-save', destination: { rootId: 'output', relativePath: 'section.docx' } });
    const reopened = await f.call('docx_open', { path: saved.data.path });
    const persisted = await text(f, reopened.view);
    assert(persisted.includes('Complete replacement.') && persisted.includes('Prefix.') && persisted.includes('Suffix.'));
    assert(!persisted.includes('Keep details.'));
    const xml = strFromU8(unzipSync(fs.readFileSync(saved.data.path))['word/document.xml']);
    assert(xml.includes('Support') && xml.includes('125') && xml.includes('Training') && xml.includes('300'));
  } finally { f.close(); }
});

test('section workflow refuses ambiguous/rich/stale sections and supports empty and end-of-document sections', async () => {
  const f = setup();
  try {
    for (const [id, blocks, workflow, code] of [
      ['duplicate-heading', [heading('Methodology'), paragraph('Old method.'), heading('Methodology'), paragraph('Old method.')], sectionEdit, 'TARGET_CONFLICT'],
      ['stale-section', [heading('Methodology'), paragraph('Changed already.')], sectionEdit, 'TARGET_CONFLICT'],
      ['table-section', [heading('Methodology'), priceTable], { ...sectionEdit, expectedText: '' }, 'TARGET_CONFLICT'],
    ]) {
      const view = await create(f, blocks, id);
      await refused(f, { view, operationId: id + '-edit', workflow }, code);
    }
    for (const [id, blocks, expectedText] of [
      ['empty-section', [heading('Methodology'), heading('Next'), paragraph('Keep next.')], ''],
      ['last-section', [paragraph('Prefix.'), heading('Methodology'), paragraph('Old method.')], 'Old method.'],
    ]) {
      const view = await create(f, blocks, id);
      const edited = await f.call('docx_apply', { view, operationId: id + '-edit', workflow: { ...sectionEdit, expectedText }, readback: {} });
      assert(edited.data.workflowVerification.passed);
      assert((await text(f, edited.view)).includes('New method.'));
    }
  } finally { f.close(); }
});

test('keyed workflow refuses merged, nested and multi-paragraph target cells without publication', async () => {
  const { Document, Packer, Paragraph, Table, TableRow, TableCell } = await import('docx');
  const f = setup();
  try {
    const cell = (text, extra = {}) => new TableCell({ children: [new Paragraph(text)], ...extra });
    const header = () => new TableRow({ children: [cell('Service'), cell('Price USD')] });
    const nested = new Table({ rows: [new TableRow({ children: [cell('Nested')] })] });
    for (const [name, rows] of [
      ['merged', [new TableRow({ children: [cell('Service Price USD', { columnSpan: 2 })] }), new TableRow({ children: [cell('Support'), cell('125')] })]],
      ['nested', [header(), new TableRow({ children: [cell('Support'), new TableCell({ children: [nested] })] })]],
      ['complex-cell', [header(), new TableRow({ children: [cell('Support'), new TableCell({ children: [new Paragraph('125'), new Paragraph('Extra')] })] })]],
    ]) {
      const source = path.join(f.docs, `${name}.docx`);
      fs.writeFileSync(source, await Packer.toBuffer(new Document({ sections: [{ children: [new Table({ rows })] }] })));
      const before = digest(fs.readFileSync(source));
      const view = (await f.call('docx_open', { path: source })).view;
      await refused(f, { view, operationId: name, select: { type: 'table', header: 'Price USD' }, workflow: keyEdit }, 'UNSUPPORTED_FEATURE');
      assert.equal(digest(fs.readFileSync(source)), before);
    }
  } finally { f.close(); }
});

test('section workflow refuses review-bearing documents without deleting comments', async () => {
  const f = setup();
  try {
    const file = path.join(f.docs, 'existing-comments.docx');
    fs.writeFileSync(file, await commentedDocument({ anchor: 'Old method.', comment: 'Preserve reviewer context.', before: [new Paragraph({ text: 'Methodology', heading: HeadingLevel.HEADING_1 })] }));
    const commented = await f.call('docx_open', { path: file });
    await refused(f, { view: commented.view, operationId: 'review-section', workflow: sectionEdit }, 'UNSUPPORTED_FEATURE');
    const comments = await f.call('docx_read', { view: commented.view, representation: 'domain', domain: 'comments' });
    assert(JSON.stringify(comments.data).includes('Preserve reviewer context.'));
  } finally { f.close(); }
});

test('explicit postconditions extend workflow guards without removing or overriding them', async () => {
  const f = setup();
  try {
    const view = await create(f, [paragraph('Client {{CLIENT}}.')]);
    const workflow = { type: 'fill_template', bindings: [{ placeholder: '{{CLIENT}}', value: 'Meridian', expectedCount: 1 }] };
    await refused(f, { view, operationId: 'override-template-guard', workflow, expect: { beforeOccurrences: { '{{CLIENT}}': 0 } } }, 'INVALID_ARGUMENT');
    const edited = await f.call('docx_apply', { view, operationId: 'extend-template-guard', workflow,
      expect: { changedOccurrences: 1, remainingOccurrences: { '{{CLIENT}}': 0 } }, readback: {} });
    assert.equal(edited.data.semanticDelta.changedOccurrences, 1);
    assert(edited.data.readback.items.some(item => item.after?.text === 'Client Meridian.'));
  } finally { f.close(); }
});
