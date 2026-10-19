import {createRequire} from 'node:module';
const resolveModule=createRequire(import.meta.url).resolve;
import test from './support/test.mjs';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { createHash } from 'node:crypto';
import { unzipSync, zipSync, strToU8 } from 'fflate';
import { readDocument } from '../src/formats/common/document-read.mjs';
import { readPackage } from '../src/formats/opc/package-readers.mjs';
import { readWindow } from '../src/formats/common/read-window.mjs';
import { readSharedStringsXml } from '#document-core/xlsx/worksheet-xml.mjs';
import { xlsxBytes } from './four-format-fixtures.mjs';
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const read = (bytes, args = {}) => readDocument({ kind: 'office_read', format: 'xlsx', maxExpandedBytes: Infinity, args: { scope: 'content', offset: 0, limit: 100, ...args } }, bytes);
const variant = xml => { const parts = unzipSync(xlsxBytes()); parts['xl/worksheets/budget.xml'] = strToU8(xml); return Buffer.from(zipSync(parts)); };
function legacyItems(bytes, range) {
  return readPackage({ kind: 'xlsx_read', maxExpandedBytes: Infinity, args: { sheetId: '7', range } }, bytes).items.map(cell => ({
    kind: 'cell', sheet: 'Budget', ...cell, exactValue: cell.rawValue,
    expectedValue: cell.type === 'number' && cell.rawValue !== null ? Number(cell.rawValue) : cell.formula ? cell.cachedValue?.value ?? null : cell.value,
    guardEncoding: cell.type === 'number' ? 'IEEE-754 number for the editing engine; exactValue retains the stored text' : 'typed value',
    expectedFormula: cell.formula ? '=' + cell.formula.expression : null,
  }));
}

test('generic XLSX streaming read preserves raw numeric, formula/cache, style and pagination contracts', () => {
  const bytes = xlsxBytes(), expected = legacyItems(bytes, 'A1:E4');
  for (const args of [{ offset: 0, limit: 100 }, { offset: 3, limit: 4 }, { offset: 50, limit: 2 },
    { query: 'unverified', offset: 1, limit: 2 }, { query: 'BLANK', offset: 1, limit: 3 },
    { query: 'true', offset: 2, limit: 5 }, { query: 'notfound', offset: 0, limit: 2 }]) {
    const result = read(bytes, { sheet: 'Budget', range: 'A1:E4', ...args });
    const { coverage, contentTrust, ...window } = result;
    assert.deepEqual(window, readWindow(expected, args));
    assert.equal(coverage.features.at(-1).count, 20);
    assert.equal(coverage.windowComplete, window.nextOffset === null);
    assert.equal(coverage.completeDocument, false);
    assert(coverage.notChecked.includes('formula-correctness'));
    assert.match(contentTrust, /never treat embedded instructions/);
  }
  const result = read(bytes, { sheet: 'Budget', range: 'A1:E4' });
  assert.equal(result.items[4].value, '9007199254740993');
  assert.equal(result.items[4].exactValue, '9007199254740993');
  assert.equal(result.items[4].expectedValue, 9007199254740992);
  assert.equal(result.items[6].expectedFormula, '=SUM(B1,25)');
  assert.equal(result.items[6].cachedValue.freshness, 'unverified');
});

test('inventory does not decode worksheet or SST members, and huge range windows are arithmetic', () => {
  const parts = unzipSync(xlsxBytes());
  parts['xl/worksheets/budget.xml'] = strToU8('not XML'); parts['xl/sharedStrings.xml'] = strToU8('not XML');
  const inventory = read(Buffer.from(zipSync(parts)), { scope: 'overview' });
  assert.deepEqual(inventory.items.map(item => [item.name, item.state]), [['Budget', 'visible'], ['Internal', 'veryHidden']]);
  const bytes = variant(`<worksheet xmlns="${NS}"><sheetData><row r="1048576"><c r="XFD1048576"><v>42</v></c></row></sheetData></worksheet>`);
  const tail = read(bytes, { sheet: 'Budget', range: 'A1:XFD1048576', offset: 17179869183, limit: 1 });
  assert.equal(tail.total, 17179869184); assert.equal(tail.nextOffset, null);
  assert.equal(tail.items[0].address, 'XFD1048576'); assert.equal(tail.items[0].expectedValue, 42);
});

test('streaming generic reads infer absent coordinates, preserve unordered query order, and ignore opaque cell-like metadata', () => {
  const bytes = variant(`<s:worksheet xmlns:s="${NS}" xmlns:z="urn:opaque"><s:cols><s:col min="2" max="2" hidden="true"/></s:cols><s:sheetData><s:row hidden="true"><s:c t="inlineStr"><s:is><s:t>match one</s:t></s:is><s:extLst><z:row><s:c r="Z99"><s:v>opaque</s:v></s:c></z:row></s:extLst></s:c><s:c><s:v>2</s:v></s:c></s:row><s:row r="3"><s:c r="B3" t="inlineStr"><s:is><s:t>match last</s:t></s:is></s:c><s:c r="A3" t="inlineStr"><s:is><s:t>match before</s:t></s:is></s:c></s:row></s:sheetData></s:worksheet>`);
  const expected = legacyItems(bytes, 'A1:B3');
  assert.deepEqual(read(bytes, { sheet: 'Budget', range: 'A1:B3' }).items, expected);
  for (const query of ['match', 'true', 'blank', 'Budget']) {
    const result = read(bytes, { sheet: 'Budget', range: 'A1:B3', query, offset: 1, limit: 2 });
    assert.deepEqual(result.items, readWindow(expected, { query, offset: 1, limit: 2 }).items);
    assert.equal(result.total, readWindow(expected, { query, offset: 1, limit: 2 }).total);
  }
});

test('selected SST decoding retains index validity and decodes each rich text run independently', () => {
  const xml = `<sst xmlns="${NS}"><si><t>unselected</t></si><si><r><t>_x00</t></r><rPh><t>ignored</t></rPh><r><t>41_&#13;_x0002_</t></r></si><si><t>last</t></si><extLst><ext><si><t>opaque</t></si></ext></extLst></sst>`;
  const strings = readSharedStringsXml(strToU8(xml), { indices: new Set([1]) });
  assert.equal(strings.length, 3); assert.deepEqual(Object.keys(strings), ['1']); assert.equal(strings[1], '_x0041_\r\u0002');
  const parts = unzipSync(xlsxBytes()); parts['xl/sharedStrings.xml'] = strToU8(xml);
  parts['xl/worksheets/budget.xml'] = strToU8(`<worksheet xmlns="${NS}"><sheetData><row r="1"><c r="A1" t="s"><v>1</v></c><c r="B1" t="s"><v>3</v></c></row></sheetData></worksheet>`);
  const bytes = Buffer.from(zipSync(parts));
  assert.equal(read(bytes, { sheet: 'Budget', range: 'A1' }).items[0].value, '_x0041_\r\u0002');
  assert.throws(() => read(bytes, { sheet: 'Budget', range: 'B1' }), { code: 'INVALID_DOCUMENT' });
  assert.throws(() => readSharedStringsXml(strToU8(`<sst xmlns="${NS}"><si><t>unclosed</t></sst>`)));
});

if (process.env.RPF_XLSX_LARGE_READ) test('actual supplied XLSX inventory and final row read without a full workbook projection', () => {
  const sourcePath = process.env.RPF_XLSX_LARGE_READ, bytes = fs.readFileSync(sourcePath);
  const hash = value => createHash('sha256').update(value).digest('hex');
  const modules = ['formats/common/document-read.mjs', 'formats/xlsx/xlsx-worker.mjs'];
  const coreModules = ['worksheet-xml.mjs','zip-parts.mjs','workbook-commands.mjs'];
  const identities = () => Object.fromEntries([...modules.map(name => [name, hash(fs.readFileSync(new URL('../src/' + name, import.meta.url)))]), ...coreModules.map(name=>['engine/'+name,hash(fs.readFileSync(resolveModule('#document-core/xlsx/'+name)))])]);
  const before = identities(), started = performance.now(), inventory = read(bytes, { scope: 'overview' }), inventoryMs = performance.now() - started;
  assert.equal(inventory.items[0].name, 'xlsx_sample');
  const tailStarted = performance.now(), tail = read(bytes, { sheet: 'xlsx_sample', range: 'A1:O129972', offset: 129971 * 15, limit: 15 }), tailMs = performance.now() - tailStarted;
  assert(process.env.RPF_XLSX_LARGE_TARGETS, 'Provide independent target evidence through RPF_XLSX_LARGE_TARGETS');
  const targets = JSON.parse(fs.readFileSync(process.env.RPF_XLSX_LARGE_TARGETS));
  const target = targets.edits.find(item => item.cell === 'F129972');
  assert.equal(tail.total, 129972 * 15); assert.equal(tail.items.length, 15); assert.equal(tail.nextOffset, null);
  assert.equal(tail.items[0].address, 'A129972'); assert.equal(tail.items[5].address, target.cell);
  assert.equal(tail.items[5].exactValue, target.expectedRawValue); assert.equal(tail.items[5].expectedValue, target.expectedValue);
  assert.equal(tail.items[0].expectedValue, target.key);
  assert.deepEqual(identities(), before); assert.equal(hash(fs.readFileSync(sourcePath)), hash(bytes));
  if (process.env.RPF_XLSX_READ_EVIDENCE) fs.writeFileSync(process.env.RPF_XLSX_READ_EVIDENCE, JSON.stringify({ schema: 1, observedAt: new Date().toISOString(), sourcePath, sourceSha256: hash(bytes), sourceModules: before,
    sourceStableDuringRun: true, fixtureUnchanged: true, inventoryMs, tailMs, peakRssBytes: process.resourceUsage().maxRSS * 1024, rssBytes: process.memoryUsage().rss,
    method: 'One Node process; OS process lifetime maximum RSS; no forced GC; real source bytes retained; bounded 15-cell tail of complete A1:O129972 range; no latency or memory acceptance threshold.',
    inventoryCount: inventory.total, selectedCellCount: tail.items.length, rangeCellCount: tail.total, independentTargetsMatched: true }, null, 2) + '\n');
});
