import test from './support/test.mjs';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { createServiceFixture } from './service-fixture.mjs';
import { OfficeService } from '../src/application/service.mjs';
import { configFrom, digest } from '../src/infrastructure/files.mjs';
import { decodeCsvBytes, encodeCsvText } from '#document-core/xlsx/csv-codec.mjs';
import { readCsvWorkbook, editCsvWorkbook, createCsvWorkbook } from '#document-core/xlsx/csv-document.mjs';

const text = 'id,price,note\r\n"001",125,"comma, quote "" and\n中文🙂"\r\n002,,=literal\r\n';
const open = async (f, file) => { const result = await f.call('office_open', { path: file }); assert.equal(result.document.format, 'csv'); return { documentId: result.document.id, revision: result.document.sourceRevision }; };

test('CSV Unicode encodings, exact quoting and late failure preserve all original bytes', () => {
  for (const content of ['a;b;c', 'a\tb\tc']) {
    const generated = createCsvWorkbook([[content, 'd']]);
    assert.equal(readCsvWorkbook(generated).columns, 2, 'New CSV header quotes competing separator characters');
  }
  for (const [encoding, bom] of [['utf-8', false], ['utf-8', true], ['utf-16le', true], ['utf-16be', true]]) {
    const bytes = encodeCsvText(text, encoding, bom), before = readCsvWorkbook(bytes);
    assert.equal(before.rows, 3); assert.equal(before.columns, 3);
    assert.deepEqual(decodeCsvBytes(bytes), { text, encoding, bom });
    const edited = editCsvWorkbook(bytes, [{ sheet: 'CSV', cell: 'B2', expectedValue: '125', value: '150' }], before.revision);
    assert.deepEqual(Buffer.from(edited.bytes), Buffer.from(encodeCsvText(text.replace('125', '150'), encoding, bom)));
    assert.throws(() => editCsvWorkbook(bytes, [
      { sheet: 'CSV', cell: 'B2', expectedValue: '125', value: '150' },
      { sheet: 'CSV', cell: 'C2', expectedValue: 'wrong', value: 'lost' },
    ], before.revision), /Conflict/);
    assert.equal(digest(bytes), digest(encodeCsvText(text, encoding, bom)));
  }
  assert.throws(() => decodeCsvBytes(Uint8Array.of(0xc3, 0x28)), /encoding/);
  assert.throws(() => decodeCsvBytes(Uint8Array.of(0xff, 0xfe, 0, 0, 65, 0, 0, 0)), /UTF-32/);
  assert.throws(() => decodeCsvBytes(Uint8Array.of(0, 0, 0xfe, 0xff, 0, 0, 0, 65)), /UTF-32/);
  assert.throws(() => encodeCsvText('\uD800'), /surrogate/);
});

test('CSV MCP open/query/edit/save/reopen, reusable targets and durable receipts', async () => {
  const fixture = createServiceFixture({ prepare: docs => fs.writeFileSync(path.join(docs, 'prices.csv'), text) });
  let active = fixture.service;
  const f = { ...fixture, call: async (name, args) => {
    const result = await active.call(name, args); assert(!result.isError, JSON.stringify(result.payload)); return result.payload;
  } };
  try {
    const file = path.join(f.docs, 'prices.csv'), originalHash = digest(fs.readFileSync(file)), view = await open(f, file);
    const inventory = await f.call('xlsx_query', { view });
    assert.equal(inventory.data.sheets[0].name, 'CSV'); assert.equal(inventory.data.sheets[0].rows, 3);
    const ordinary = await f.call('office_read', { view, sheet: 'CSV', range: 'A1:C3' });
    assert.equal(ordinary.data.items.find(item => item.cell === 'C3').value, '=literal');
    const queried = await f.call('xlsx_query', { view, sheet: 'CSV', range: 'A1:C3', keys: { column: 'A', values: ['001'] } });
    const target = queried.data.items.find(item => item.cell === 'B2').target;
    const args = { view, operationId: 'csv-price', commands: [{ target, expectedValue: '125', value: '150' }] };
    const changed = await f.call('xlsx_apply', args), replay = await f.call('xlsx_apply', args);
    assert.equal(changed.view.revision, replay.view.revision);
    assert.equal(changed.data.affectedTargets[0].before.value, '125');
    assert.equal(changed.data.affectedTargets[0].after.value, '150');
    const exhausted = await f.call('office_read', { view, scope: 'overview', offset: 1 });
    assert.deepEqual(exhausted.data.items, []);
    let refusalId = 0;
    for (const commands of [
      [{ sheet: 'CSV', cell: 'B2', formula: '1+1' }],
      [{ type: 'formatRange', sheet: 'CSV', range: 'B2', format: { bold: true } }],
      [{ sheet: 'CSV', cell: 'B2', value: '200' }, { sheet: 'CSV', cell: 'C2', expectedValue: 'wrong', value: 'lost' }],
    ]) {
      const refused = await active.call('xlsx_apply', { view: changed.view, operationId: `refuse-${++refusalId}`, commands });
      assert(refused.isError);
      assert.equal(refused.payload.error.code, refusalId === 3 ? 'TARGET_CONFLICT' : 'UNSUPPORTED_FEATURE');
      const unchanged = await f.call('xlsx_query', { view: changed.view, sheet: 'CSV', range: 'B2' });
      assert.equal(unchanged.data.items[0].value, '150');
    }
    const wrongExtension = await active.call('office_save', { view: changed.view, operationId: 'wrong-extension', destination: { rootId: 'output', relativePath: 'bad.xlsx' } });
    assert(wrongExtension.isError); assert(!fs.existsSync(path.join(f.docs, 'bad.xlsx')));
    const saved = await f.call('office_save', { view: changed.view, operationId: 'csv-save', destination: { rootId: 'output', relativePath: 'saved.csv' } });
    assert.equal(fs.readFileSync(saved.data.path, 'utf8'), text.replace('125', '150'));
    const asset = await f.call('office_asset_info', { assetId: saved.data.assetId });
    assert.equal(asset.data.mediaType, 'text/csv');
    active.close(); active = new OfficeService(configFrom(f.config));
    const receipt = await f.call('office_operation_status', { operationId: 'csv-save' });
    assert.equal(receipt.data.data.path, saved.data.path);
    const reopened = await open(f, saved.data.path), read = await f.call('xlsx_query', { view: reopened, sheet: 'CSV', range: 'A2:C3' });
    assert.equal(read.data.items[0].value, '001'); assert.equal(read.data.items[1].value, '150');
    assert.equal(read.data.items[4].value, ''); assert.equal(read.data.items[5].value, '=literal');
    assert.equal(digest(fs.readFileSync(file)), originalHash);
  } finally { active.close(); fs.rmSync(f.root, { recursive: true, force: true }); }
});

test('CSV delimiter detection agrees with the editor and preserves semicolon/tab records', async () => {
  const f = createServiceFixture();
  try {
    for (const [label, delimiter] of [['semicolon', ';'], ['tab', '\t']]) {
      const source = `id${delimiter}price${delimiter}note\r\n001${delimiter}125${delimiter}"comma, and ; and\ttab"\r\n`;
      const file = path.join(f.docs, `${label}.csv`); fs.writeFileSync(file, source);
      const view = await open(f, file), query = await f.call('xlsx_query', { view, sheet: 'CSV', range: 'A2:C2' });
      assert.equal(query.data.items[0].value, '001'); assert.equal(query.data.items[1].value, '125');
      const changed = await f.call('xlsx_apply', { view, operationId: label, commands: [{ sheet: 'CSV', cell: 'B2', expectedValue: '125', value: '150' }] });
      const saved = await f.call('office_save', { view: changed.view, operationId: `save-${label}`, destination: { rootId: 'output', relativePath: `saved-${label}.csv` } });
      assert.equal(fs.readFileSync(saved.data.path, 'utf8'), source.replace('125', '150'));
    }
  } finally { f.close(); }
});

test('CSV creation and common write_cells use the same literal preservation path', async () => {
  const f = createServiceFixture();
  try {
    const created = await f.call('office_create', { operationId: 'create-csv', document: { format: 'csv', rows: [['id', 'price'], ['001', 125], ['=literal', null]] } });
    const edited = await f.call('office_edit', { view: created.view, operationId: 'edit-csv', action: { type: 'write_cells', sheet: 'CSV', cells: [{ address: 'B2', value: 150, expectedValue: '125', expectedFormula: null }] } });
    const saved = await f.call('office_save', { view: edited.view, operationId: 'save-created-csv', destination: { rootId: 'output', relativePath: 'created.csv' } });
    assert.equal(fs.readFileSync(saved.data.path, 'utf8'), 'id,price\r\n001,150\r\n=literal,\r\n');
  } finally { f.close(); }
});

test('CSV MCP reads and keyed edits past row 120000 without a whole-range cell list', async () => {
  const f = createServiceFixture({ prepare: docs => fs.writeFileSync(path.join(docs, 'large.csv'),
    'id,price,note\n' + Array.from({ length: 120001 }, (_, row) => `${row},125,untouched`).join('\n')) });
  try {
    const file = path.join(f.docs, 'large.csv'), view = await open(f, file);
    const query = await f.call('xlsx_query', { view, sheet: 'CSV', range: 'A1:C120002', keys: { column: 'A', values: ['120000'] } });
    assert.equal(query.data.total, 3); assert.equal(query.data.items[0].cell, 'A120002');
    const changed = await f.call('xlsx_apply', { view, operationId: 'last-row', commands: [{ type: 'updateRows', sheet: 'CSV', range: 'A2:C120002', keyColumn: 'A', updates: [{ key: '120000', cells: [{ column: 'B', expectedValue: '125', value: '150' }] }] }] });
    const read = await f.call('xlsx_query', { view: changed.view, sheet: 'CSV', range: 'A1:C120002', offset: 360003, limit: 3 });
    assert.equal(read.data.total, 360006); assert.equal(read.data.items[1].value, '150'); assert.equal(read.data.nextOffset, null);
    const saved = await f.call('office_save', { view: changed.view, operationId: 'save-last-row', destination: { rootId: 'output', relativePath: 'large-saved.csv' } });
    const expected = fs.readFileSync(file, 'utf8').replace(/120000,125,untouched$/, '120000,150,untouched');
    assert.equal(fs.readFileSync(saved.data.path, 'utf8'), expected);
    assert(Buffer.byteLength(JSON.stringify(query)) < 10000);
  } finally { f.close(); }
});

test('CSV queries have no Excel address ceiling and reject table-only selectors', async () => {
  const longText = '中文🙂'.repeat(12000);
  const f = createServiceFixture({ prepare: docs => fs.writeFileSync(path.join(docs, 'long.csv'), 'id,note\r\n001,"' + longText + '"\r\n') });
  try {
    const view = await open(f, path.join(f.docs, 'long.csv'));
    // Only two output cells are materialized; the source has two rows.
    const far = await f.call('xlsx_query', { view, sheet: 'CSV', range: 'AAAA10000000:AAAA10000001', limit: 2 });
    assert.equal(far.data.total, 2); assert.equal(far.data.items[0].cell, 'AAAA10000000');
    assert.deepEqual(far.data.items.map(item => item.value), [null, null]);
    const named = await f.service.call('xlsx_query', { view, table: 'Missing' });
    assert(named.isError); assert.equal(named.payload.error.code, 'UNSUPPORTED_FEATURE');
  } finally { f.close(); }
});

test('CSV long literal fields survive guarded MCP save', async () => {
  const longText = '中文🙂'.repeat(12000);
  const f = createServiceFixture({ prepare: docs => fs.writeFileSync(path.join(docs, 'long.csv'), 'id,note\r\n001,"' + longText + '"\r\n') });
  try {
    const view = await open(f, path.join(f.docs, 'long.csv'));
    const edited = await f.call('xlsx_apply', { view, operationId: 'long-literal', commands: [{ sheet: 'CSV', cell: 'B2', expectedValue: longText, expectedRawValue: longText, value: longText + ' preserved' }] });
    const saved = await f.call('office_save', { view: edited.view, operationId: 'save-long-literal', destination: { rootId: 'output', relativePath: 'long-saved.csv' } });
    assert.equal(fs.readFileSync(saved.data.path, 'utf8'), 'id,note\r\n001,"' + longText + ' preserved"\r\n');
  } finally { f.close(); }
});

test('editing a header quotes other delimiter candidates so reopen preserves the table', () => {
  const source = Buffer.from('id,note\r\n001,"unchanged, quote"\r\n');
  const edited = editCsvWorkbook(source, [{ sheet: 'CSV', cell: 'B1', value: 'a;b;c\tend' }]);
  assert.equal(Buffer.from(edited.bytes).toString(), 'id,"a;b;c\tend"\r\n001,"unchanged, quote"\r\n');
  const reopened = readCsvWorkbook(edited.bytes);
  assert.equal(reopened.delimiter, ','); assert.equal(reopened.columns, 2);
  assert.equal(reopened.text.slice(reopened.rowOffsets[1]), '001,"unchanged, quote"\r\n');
});
