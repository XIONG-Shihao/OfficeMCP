import {fixtureDirectory,fixturePath,oraclePath} from '../scripts/engine-source.mjs';
import test, { after } from './support/test.mjs';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { WorkbookSession, readCalculatedWorkbook } from '#document-core/xlsx/workbook-service.mjs';
import { createWorkbook } from '#document-core/xlsx/workbook-package.mjs';
import { workbookCommand } from '../src/formats/xlsx/xlsx-contract.mjs';
import { OfficeService } from '../src/application/service.mjs';
import { configFrom } from '../src/infrastructure/files.mjs';
import { createServiceFixture } from './service-fixture.mjs';

const repository = path.resolve(import.meta.dirname, '..');
const generatedCorpus = path.join(repository, 'artifacts/xlsx-corpus');
const transientCorpus = fs.existsSync(path.join(generatedCorpus, 'generated.json')) ? null : fs.mkdtempSync(path.join(os.tmpdir(), 'xlsx-gate-corpus-'));
const corpus = transientCorpus ?? generatedCorpus;
if (transientCorpus) {
  const generated = spawnSync('python3', [fixturePath('build-corpus.py'), '--output', corpus], { encoding: 'utf8', timeout: 30000 });
  if (generated.status !== 0) fs.rmSync(transientCorpus, { recursive: true, force: true });
  assert.equal(generated.status, 0, generated.stderr || String(generated.error));
}
const keys = JSON.parse(fs.readFileSync(fixturePath('expectations.json'))).fixtures;
const hash = value => createHash('sha256').update(value).digest('hex');
const fixture = id => fs.readFileSync(path.join(corpus, `${id}.xlsx`));
const cell = (view, target) => { const split = target.lastIndexOf('!'); return view.sheets.find(s => s.name === target.slice(0, split)).cells[target.slice(split + 1)]; };
const reports = [];
function sourceIdentity() {
  const files = [];
  const visit = directory => { for (const entry of fs.readdirSync(directory, { withFileTypes: true })) { const file = path.join(directory, entry.name); if (entry.isDirectory()) visit(file); else if (/\.(mjs|json|py)$/.test(file)) files.push(file); } };
  for (const relative of ['src', '.generated/document-core', '.generated/native']) visit(path.join(repository, relative));
  visit(fixtureDirectory);
  files.push(import.meta.filename, oraclePath('xlsx-package-check.py'), oraclePath('xlsx-reference.py'), path.join(repository, 'npm-shrinkwrap.json'));
  const members = Object.fromEntries(files.sort().map(file => [path.relative(repository, file), hash(fs.readFileSync(file))]));
  return { sha256: hash(JSON.stringify(members)), files: members };
}
const source = sourceIdentity();
function gate(name, run) {
  test(name, { timeout: 90000 }, async () => {
    const start = performance.now();
    try { reports.push({ name, passed: true, result: await run(), elapsedMs: performance.now() - start }); }
    catch (error) { reports.push({ name, passed: false, error: String(error), elapsedMs: performance.now() - start }); throw error; }
  });
}
function near(actual, expected, label) {
  assert.equal(typeof actual, 'number', label);
  assert(Math.abs(actual - expected) <= Math.max(1e-9, Math.abs(expected) * 1e-12), `${label}: ${actual} != ${expected}`);
}
function checkView(view, expected, initial = false) {
  for (const [target, key] of Object.entries(expected)) {
    const actual = cell(view, target);
    assert(actual, `Missing ${target}`);
    if ('calculated' in key) { assert.equal(actual.calculation?.status, 'calculated', target); near(actual.calculation.value, key.calculated, target); }
    if ('status' in key) assert.equal(actual.calculation?.status, key.status, target);
    if ('error' in key) { assert.equal(actual.calculation?.status, 'error', target); assert.equal(actual.calculation.error, key.error, target); }
    if ('rawValue' in key) assert.equal(actual.rawValue, key.rawValue, target);
    if ('formula' in key) assert.equal(actual.formula, key.formula, target);
    if ('value' in key) key.value === null ? assert.equal(actual.dataType, 'blank', target) : assert.equal(actual.value, key.value, target);
    if (initial && 'storedCache' in key) assert.equal(actual.cachedValue, key.storedCache, `${target} stored cache is separate from calculation`);
    if (actual.calculation) assert.equal(actual.calculation.revision, view.revision, `${target} calculation revision`);
  }
}
function savedKeys(expected) {
  return Object.fromEntries(Object.entries(expected).map(([target, key]) => [target, {
    ...Object.fromEntries(Object.entries(key).filter(([property]) => ['value', 'rawValue', 'formula', 'calculated', 'error', 'style', 'mayOmitBlankCell'].includes(property))),
    ...(key.status ? { rawValue: null } : {}),
  }]));
}
function independentlyCheck(sourceBytes, savedBytes, request = {}, expectedPass = true) {
  const temporary = fs.mkdtempSync(path.join(os.tmpdir(), 'xlsx-independent-'));
  try {
    const sourceFile = path.join(temporary, 'source.xlsx'), savedFile = path.join(temporary, 'delivered.xlsx');
    fs.writeFileSync(sourceFile, sourceBytes); fs.writeFileSync(savedFile, savedBytes);
    const result = spawnSync('python3', [oraclePath('xlsx-package-check.py')], {
      input: JSON.stringify({ source: sourceFile, saved: savedFile, ...request }), encoding: 'utf8', timeout: 30000, maxBuffer: 8 * 1024 * 1024,
    });
    assert(result.stdout, result.stderr || String(result.error));
    const report = JSON.parse(result.stdout);
    assert.equal(report.passed, expectedPass, JSON.stringify({ failures: report.failures.slice(0, 20), failureCount: report.failures.length }));
    return report;
  } finally { fs.rmSync(temporary, { recursive: true, force: true }); }
}
function expectedValues(values) { return Object.fromEntries(Object.entries(values).map(([target, value]) => [target, typeof value === 'number' ? { calculated: value } : { value }])); }
function targetsFor(edits) { const targets = {}; for (const edit of edits) (targets[edit.sheet] ??= []).push(edit.cell); return targets; }

for (const id of ['pricing-schedule', 'budget', 'calculation-states', 'dates-1900', 'dates-1904']) {
  gate(`G2 independent ${id} values, states and saved caches`, async () => {
    const original = fixture(id), session = new WorkbookSession(original), view = await session.read();
    checkView(view, keys[id].cells, true);
    assert.equal(view.sheets[0].date1904, keys[id].dateSystem === '1904');
    const result = await session.apply({ operationId: `recalculate-${id}`, expectedRevision: view.revision, commands: [{ type: 'recalculate' }] });
    const reopened = await readCalculatedWorkbook(result.bytes);
    checkView(reopened, savedKeys(keys[id].cells));
    assert.equal(hash(fixture(id)), hash(original));
    const delivered = independentlyCheck(original, result.bytes, { expected: savedKeys(keys[id].cells) });
    return { fixtureId: id, fixtureSha256: hash(original), calculation: result.calculation, delivered };
  });
}

gate('G2 atomic failure, stale revisions, operation retries and cancellation', async () => {
  const original = fixture('pricing-schedule'), session = new WorkbookSession(original), initial = await session.read();
  await assert.rejects(session.apply({ operationId: 'failed-batch', expectedRevision: initial.revision, commands: [
    { sheet: 'Pricing', cell: 'B2', expectedValue: 2, value: 3 },
    { sheet: 'Pricing', cell: 'C2', expectedValue: 999, value: 150 },
  ] }), /unexpected value/);
  assert.equal(hash(session.bytes), hash(original));
  const args = { operationId: 'accepted', expectedRevision: initial.revision, commands: keys['pricing-schedule'].scenario.edits };
  const result = await session.apply(args), retry = await session.apply(args);
  assert.strictEqual(retry, result);
  await assert.rejects(session.apply({ ...args, commands: [{ sheet: 'Pricing', cell: 'B2', value: 4 }] }), /operation ID/);
  await assert.rejects(session.apply({ ...args, operationId: 'stale' }), /revision is stale/);
  const acceptedHash = hash(session.bytes), abort = new AbortController();
  const pending = session.apply({ operationId: 'cancelled', expectedRevision: result.view.revision, commands: [{ sheet: 'Pricing', cell: 'B2', value: 8 }] }, { signal: abort.signal });
  abort.abort(); await assert.rejects(pending, /cancelled/);
  assert.equal(hash(session.bytes), acceptedHash);
  checkView(result.view, expectedValues(keys['pricing-schedule'].scenario.expected));
  return { fixtureSha256: hash(original), acceptedRevision: result.view.revision, atomicFailure: true, replayExactlyOnce: true, staleRefused: true, cancellationPreservedAcceptedBytes: true };
});

gate('G3.5 independent saved XML removes unsupported caches and sets exact recalculation flags', async () => {
  const results = [];
  for (const id of ['calculation-states', 'pricing-schedule']) {
    const parts = unzipSync(fixture(id)), incomplete = id === 'calculation-states';
    if (incomplete) parts['xl/worksheets/sheet1.xml'] = strToU8(strFromU8(parts['xl/worksheets/sheet1.xml']).replace('</sheetData>', '<row r="4"><c r="C4"><f t="shared" si="0" ref="C4:C5">A2*2</f><v>999</v></c></row><row r="5"><c r="C5"><f t="shared" si="0"/><v>999</v></c></row></sheetData>'));
    const original = zipSync(parts), session = new WorkbookSession(original), before = await session.read();
    const result = await session.apply({ operationId: `cache-policy-${id}`, expectedRevision: before.revision, commands: [{ type: 'recalculate' }] });
    const request = { expected: savedKeys(keys[id].cells), expectedCalculationProperties: {
      fullCalcOnLoad: incomplete ? '1' : '0', forceFullCalc: incomplete ? '1' : '0', calcMode: 'auto',
    }, ...(incomplete ? { absentFormulaCaches: { Diagnostics: ['C1', 'C2', 'D1', 'D2', 'C4', 'C5'] } } : {}) };
    if (incomplete) for (const ref of ['C4', 'C5']) {
      assert.equal(cell(result.view, `Diagnostics!${ref}`).formulaType, 'shared');
      assert.equal(cell(result.view, `Diagnostics!${ref}`).calculation.status, 'unresolved');
    }
    const delivered = independentlyCheck(original, result.bytes, request), negativeChecks = [];
    for (const property of ['fullCalcOnLoad', 'forceFullCalc']) {
      const damaged = unzipSync(result.bytes);
      damaged['xl/workbook.xml'] = strToU8(strFromU8(damaged['xl/workbook.xml']).replace(`${property}="${incomplete ? '1' : '0'}"`, `${property}="${incomplete ? '0' : '1'}"`));
      const checked = independentlyCheck(original, zipSync(damaged), request, false);
      assert(checked.failures.some(failure => failure.kind === 'calculation-property' && failure.property === property));
      negativeChecks.push({ property, rejected: true });
    }
    if (incomplete) for (const cache of ['<v>999</v>', '<v/>']) {
      const damaged = unzipSync(result.bytes);
      damaged['xl/worksheets/sheet1.xml'] = strToU8(strFromU8(damaged['xl/worksheets/sheet1.xml']).replace(/<c\b[^>]*r="C1"[^>]*>[\s\S]*?<\/c>/, node => node.replace('</c>', `${cache}</c>`)));
      const checked = independentlyCheck(original, zipSync(damaged), request, false);
      assert(checked.failures.some(failure => failure.kind === 'formula-cache-not-absent' && failure.target === 'Diagnostics!C1'));
      negativeChecks.push({ cache, rejected: true });
    }
    results.push({ baseFixture: id, privateVariantSha256: hash(original), delivered, negativeChecks });
  }
  return results;
});

gate('W2/W4 unique business keys require expected old values and refuse duplicates atomically', async () => {
  const original = fixture('response-matrix'), session = new WorkbookSession(original), view = await session.read();
  const command = { type: 'updateRows', sheet: 'Responses', range: 'A2:D3', keyColumn: 'A', updates: [{ key: 'REQ-001', cells: [{ column: 'C', expectedValue: '待填写', value: 'Compliant' }] }] };
  for (const [name, updates, error] of [
    ['missing', [{ key: 'MISSING', cells: [{ column: 'C', expectedValue: 'Pending', value: 'x' }] }], /Missing key/],
    ['unguarded', [{ key: 'REQ-001', cells: [{ column: 'C', value: 'x' }] }], /expected old values/],
    ['duplicate-request', [command.updates[0], command.updates[0]], /Duplicate requested key/],
  ]) {
    await assert.rejects(session.apply({ operationId: name, expectedRevision: view.revision, commands: [{ ...command, updates }] }), error);
    assert.equal(hash(session.bytes), hash(original));
  }
  const duplicateParts = unzipSync(original);
  duplicateParts['xl/worksheets/sheet1.xml'] = strToU8(strFromU8(duplicateParts['xl/worksheets/sheet1.xml']).replace('REQ-002', 'REQ-001'));
  const duplicateBytes = zipSync(duplicateParts), duplicateSession = new WorkbookSession(duplicateBytes), duplicateView = await duplicateSession.read();
  await assert.rejects(duplicateSession.apply({ operationId: 'duplicate-existing', expectedRevision: duplicateView.revision, commands: [command] }), /Duplicate key/);
  assert.equal(hash(duplicateSession.bytes), hash(duplicateBytes));
  const result = await session.apply({ operationId: 'response', expectedRevision: view.revision, commands: [command] });
  const delivered = independentlyCheck(original, result.bytes, { targets: { Responses: ['C2'] }, expected: { 'Responses!C2': { value: 'Compliant' } } });
  return { fixtureSha256: hash(original), duplicateVariantSha256: hash(duplicateBytes), delivered };
});

gate('W3 budget insertion expands total, mixed and cross-sheet references', async () => {
  const original = fixture('budget'), session = new WorkbookSession(original), view = await session.read();
  const result = await session.apply({ operationId: 'insert-budget', expectedRevision: view.revision, commands: [
    { sheet: 'Budget', operation: 'insertRow', index: 2 },
    { type: 'setCells', sheet: 'Budget', cells: [{ cell: 'A3', value: 'Travel' }, { cell: 'B3', value: 4 }, { cell: 'C3', value: 25 }, { cell: 'D3', formula: 'B3*C3' }] },
  ] });
  const expected = {
    'Budget!A3': { value: 'Travel' }, 'Budget!D2': { formula: '$B2*C$2', calculated: 100 },
    'Budget!D3': { formula: 'B3*C3', calculated: 100 }, 'Budget!D4': { formula: 'B4*C4', calculated: 150 },
    'Budget!D5': { formula: 'SUM(D2:D4)', calculated: 350 }, 'Budget!F2': { formula: 'Rates!B1*D5', calculated: 35 },
    'Budget!G2': { formula: 'AVERAGE(D2:D4)', calculated: 350 / 3 }, 'Budget!H2': { formula: 'MIN(D2:D4)', calculated: 100 },
    'Budget!I2': { formula: 'MAX(D2:D4)', calculated: 150 }, 'Budget!J2': { formula: 'COUNT(D2:D4)', calculated: 3 },
  };
  checkView(result.view, expected);
  const delivered = independentlyCheck(original, result.bytes, { structural: { Budget: { index: 2 } }, targets: { Budget: ['A3', 'B3', 'C3', 'D3'] }, expected });
  return { fixtureSha256: hash(original), delivered, undoRedo: 'Not measured by this backend test; requires editor history journey.' };
});

gate('W3 row insertion preserves local table references and updates named/cross-sheet references', async () => {
  const parts = unzipSync(fixture('pricing-schedule'));
  parts['xl/workbook.xml'] = strToU8(strFromU8(parts['xl/workbook.xml']).replace('</definedNames>', '<definedName name="SecondAmount">Pricing!$D$3</definedName></definedNames>'));
  parts['xl/worksheets/sheet2.xml'] = strToU8(strFromU8(parts['xl/worksheets/sheet2.xml']).replace('</sheetData>', '<row r="2"><c r="B2"><f>Pricing!D4</f><v>330</v></c><c r="C2"><f>SecondAmount</f><v>80</v></c></row></sheetData>'));
  const original = zipSync(parts), session = new WorkbookSession(original), view = await session.read();
  const result = await session.apply({ operationId: 'insert-table', expectedRevision: view.revision, commands: [
    { sheet: 'Pricing', operation: 'insertRow', index: 2 },
    { type: 'setCells', sheet: 'Pricing', cells: [{ cell: 'A3', value: 'NEW' }, { cell: 'B3', value: 4 }, { cell: 'C3', value: 25 }, { cell: 'D3', formula: 'B3*C3' }, { cell: 'E3', formula: 'TaxRate' }, { cell: 'F3', formula: 'D3*(1+E3)' }] },
  ] });
  const expected = {
    'Pricing!D2': { formula: 'B2*C2', calculated: 250 }, 'Pricing!E2': { formula: 'TaxRate', calculated: 0.09 }, 'Pricing!F2': { formula: 'D2*(1+E2)', calculated: 272.5 },
    'Pricing!D3': { formula: 'B3*C3', calculated: 100 }, 'Pricing!E3': { formula: 'TaxRate', calculated: 0.09 }, 'Pricing!F3': { formula: 'D3*(1+E3)', calculated: 109 },
    'Pricing!D4': { formula: 'B4*C4', calculated: 80 }, 'Pricing!E4': { formula: 'TaxRate', calculated: 0.09 }, 'Pricing!F4': { formula: 'D4*(1+E4)', calculated: 87.2 },
    'Pricing!D5': { formula: 'SUM(D2:D4)', calculated: 430 }, 'Pricing!F5': { formula: 'SUM(F2:F4)', calculated: 468.7 },
    'Pricing!D6': { formula: 'SUM(PricingItems[Subtotal])', calculated: 430 }, 'Rates!B2': { formula: 'Pricing!D5', calculated: 430 }, 'Rates!C2': { formula: 'SecondAmount', calculated: 80 },
  };
  checkView(result.view, expected);
  assert.equal(result.view.tables[0].ref, 'A1:F4');
  assert.equal(result.view.namedRanges.find(name => name.name === 'SecondAmount').formula, 'Pricing!$D$4');
  const delivered = independentlyCheck(original, result.bytes, { structural: { Pricing: { index: 2 } }, targets: { Pricing: ['A3', 'B3', 'C3', 'D3', 'E3', 'F3'], Rates: ['B2'] }, expectedNames: { SecondAmount: 'Pricing!$D$4' }, expectedTables: { 'xl/tables/table1.xml': 'A1:F4' }, expected });
  return { baseFixture: 'pricing-schedule', baseSha256: hash(fixture('pricing-schedule')), privateVariantSha256: hash(original), delivered };
});

gate('W3 inserted row inherits style and auto-filled formula without explicit formula command', async () => {
  const parts = unzipSync(fixture('budget'));
  parts['xl/worksheets/sheet1.xml'] = strToU8('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Line</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Base</t></is></c><c r="B2" s="2"><v>2</v></c><c r="C2" s="2"><v>10</v></c><c r="D2" s="2"><f>B2*C2</f><v>20</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>Total</t></is></c><c r="D3" s="2"><f>SUM(D2:D2)</f><v>20</v></c></row></sheetData></worksheet>');
  const original = zipSync(parts), session = new WorkbookSession(original), view = await session.read();
  const result = await session.apply({ operationId: 'automatic-fill', expectedRevision: view.revision, commands: [
    { sheet: 'Budget', operation: 'insertRow', index: 2 },
    { type: 'setCells', sheet: 'Budget', cells: [{ cell: 'A3', value: 'Travel' }, { cell: 'B3', value: 3 }, { cell: 'C3', value: 7 }] },
  ] });
  const expected = { 'Budget!B3': { value: 3, style: '2' }, 'Budget!C3': { value: 7, style: '2' },
    'Budget!D3': { formula: 'B3*C3', calculated: 21, style: '2' }, 'Budget!D4': { formula: 'SUM(D2:D3)', calculated: 41, style: '2' } };
  checkView(result.view, expected);
  const delivered = independentlyCheck(original, result.bytes, { structural: { Budget: { index: 2 } }, targets: { Budget: ['A3', 'B3', 'C3', 'D3'] }, expected });
  return { baseFixture: 'budget', privateVariantSha256: hash(original), delivered };
});

for (const scenario of [
  { label: 'single-cell name at the insertion boundary', name: 'UnitPrice', definition: 'Budget!$B$2', index: 2 },
  { label: 'unqualified local name owned by another sheet', name: 'LocalRate', definition: '$B$1', scope: ' localSheetId="1"', index: 1 },
]) {
  gate(`W3 row insertion preserves ${scenario.label} and rolls back a failed structural batch`, async () => {
    const parts = unzipSync(fixture('budget'));
    parts['xl/workbook.xml'] = strToU8(strFromU8(parts['xl/workbook.xml']).replace('</sheets>', `</sheets><definedNames><definedName name="${scenario.name}"${scenario.scope ?? ''}>${scenario.definition}</definedName></definedNames>`));
    parts['xl/worksheets/sheet1.xml'] = strToU8('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Price</t></is></c></row><row r="2"><c r="B2"><v>125</v></c></row></sheetData></worksheet>');
    const original = zipSync(parts), session = new WorkbookSession(original), view = await session.read();
    const insert = { sheet: 'Budget', operation: 'insertRow', index: scenario.index };
    await assert.rejects(session.apply({ operationId: 'failed-name-insert', expectedRevision: view.revision, commands: [insert,
      { sheet: 'Rates', cell: 'B1', expectedValue: 999, value: 0.2 },
    ] }), /unexpected value/);
    assert.equal(hash(session.bytes), hash(original), 'failed batch must preserve original bytes');
    assert.deepEqual((await session.read()).namedRanges, view.namedRanges, 'failed batch must preserve accepted names');
    const result = await session.apply({ operationId: 'accepted-name-insert', expectedRevision: view.revision, commands: [insert] });
    assert.equal(result.view.namedRanges.find(name => name.name === scenario.name).formula, scenario.definition);
    const delivered = independentlyCheck(original, result.bytes, { structural: { Budget: { index: scenario.index } }, expectedNames: { [scenario.name]: scenario.definition } });
    return { baseFixture: 'budget', privateVariantSha256: hash(original), atomicStructuralFailure: true, delivered };
  });
}

gate('W5 rich package literal/format edits preserve original drawing, image, validation and styles', async () => {
  const parts = unzipSync(fixture('preservation-rich'));
  parts['xl/styles.xml'] = strToU8(strFromU8(parts['xl/styles.xml']).replace('<fills count="2">', '<fills count="3">').replace('</fills>', '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill></fills>').replace('<xf numFmtId="164" fontId="0" fillId="0"', '<xf numFmtId="164" fontId="0" fillId="2"'));
  parts['xl/worksheets/sheet1.xml'] = strToU8(strFromU8(parts['xl/worksheets/sheet1.xml']).replace('<c r="B2">', '<c r="B2" s="2">'));
  const original = zipSync(parts), session = new WorkbookSession(original), view = await session.read();
  const result = await session.apply({ operationId: 'rich-edit', expectedRevision: view.revision, commands: [
    { sheet: 'Data', cell: 'D2', expectedValue: 'Pending', value: 'Reviewed' },
    { type: 'formatRange', sheet: 'Data', range: 'B2:B3', format: { bold: true, italic: true, numberFormat: 'number' } },
  ] });
  for (const ref of ['B2', 'B3']) { assert.equal(cell(result.view, `Data!${ref}`).bold, true); assert.equal(cell(result.view, `Data!${ref}`).italic, true); }
  const delivered = independentlyCheck(original, result.bytes, { targets: { Data: ['D2', 'B2', 'B3'] }, styleTargets: { Data: ['B2', 'B3'] }, appendStyles: true, expected: { 'Data!D2': { value: 'Reviewed' }, 'Data!B2': { value: 125 }, 'Data!B3': { value: 80 } } });
  const savedParts = unzipSync(result.bytes), styles = strFromU8(savedParts['xl/styles.xml']);
  const matches = [...styles.matchAll(/<xf\b[^>]*>/g)];
  assert(styles.includes('fgColor rgb="FFFFFF00"'), 'original yellow fill remains present');
  assert(matches.some(match => /fillId="2"/.test(match[0]) && /numFmtId="2"/.test(match[0])), 'new number format keeps existing fill');
  return { baseFixture: 'preservation-rich', privateVariantSha256: hash(original), delivered };
});

gate('W5 Home ribbon styles survive a preserving save and independent XML inspection', async () => {
  const original = fixture('preservation-rich');
  const session = new WorkbookSession(original), view = await session.read();
  const format = { bold: true, underline: true, fontFamily: 'Aptos Narrow', fontSize: 12,
    fontColor: 'C00000', fillColor: 'FFF2CC', horizontalAlign: 'center', verticalAlign: 'bottom',
    wrapText: true, border: 'all', numberFormat: 'currency' };
  const result = await session.apply({ operationId: 'home-styles', expectedRevision: view.revision,
    commands: [{ type: 'formatRange', sheet: 'Data', range: 'B2:B3', format }] });
  const reopened = await readCalculatedWorkbook(result.bytes);
  for (const ref of ['B2', 'B3']) {
    const styled = cell(reopened, `Data!${ref}`);
    for (const [key, value] of Object.entries(format)) if (key !== 'numberFormat') assert.equal(styled[key], value, `${ref} ${key}`);
    assert.equal(styled.value, view.sheets.find(sheet => sheet.name === 'Data').cells[ref].value);
  }
  const sourceParts = unzipSync(original), savedParts = unzipSync(result.bytes);
  for (const name of Object.keys(sourceParts)) {
    if (name === 'xl/styles.xml' || name === 'xl/worksheets/sheet1.xml' || name === 'xl/workbook.xml') continue;
    assert.equal(hash(savedParts[name]), hash(sourceParts[name]), `unrelated part changed: ${name}`);
  }
  const styles = strFromU8(savedParts['xl/styles.xml']);
  assert.equal(cell(reopened, 'Data!B2').styleId, cell(reopened, 'Data!B3').styleId, 'range formatting reuses one style');
  const styleCheck = spawnSync('python3', ['-c', `import sys, xml.etree.ElementTree as E
root=E.fromstring(sys.stdin.read())
ns={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
for tag in ('fonts','fills','borders','cellXfs'):
 node=root.find('m:'+tag,ns)
 assert node is not None and int(node.attrib['count'])==len(node), tag
xfs=root.find('m:cellXfs',ns)
assert len(xfs)>=4
assert any(x.attrib.get('numFmtId','')=='164' for x in xfs)
print('ok')`], { input: styles, encoding: 'utf8' });
  assert.equal(styleCheck.status, 0, styleCheck.stderr);
  const cleared = await session.apply({ operationId: 'clear-one-style', expectedRevision: result.view.revision,
    commands: [{ type: 'formatRange', sheet: 'Data', range: 'B2', format: { styleId: 0 } }] });
  const clearedView = await readCalculatedWorkbook(cleared.bytes);
  assert.equal(cell(clearedView, 'Data!B2').bold, false);
  assert.equal(cell(clearedView, 'Data!B2').fillColor, undefined);
  assert.equal(cell(clearedView, 'Data!B2').value, cell(view, 'Data!B2').value);
  assert.equal(cell(clearedView, 'Data!B3').fillColor, format.fillColor);
  return { sourceSha256: hash(original), savedSha256: hash(result.bytes), styles: 'independently parsed' };
});

gate('W5 qualified worksheet add, duplicate and rename keep package relationships valid', async () => {
  for (const command of [{ sheet: 'Notes', operation: 'addSheet' },
    { sheet: 'Sheet1', operation: 'duplicateSheet', newName: 'Copy' },
    { sheet: 'Notes', operation: 'renameSheet', newName: 'Details' }])
    assert.equal(workbookCommand.safeParse(command).success, true, 'public MCP command schema accepts qualified sheet actions');
  const source = createWorkbook(), session = new WorkbookSession(source), initial = await session.read();
  const seeded = await session.apply({ operationId: 'seed-sheet', expectedRevision: initial.revision,
    commands: [{ sheet: 'Sheet1', cell: 'A1', value: 'Original' }] });
  const added = await session.apply({ operationId: 'add-sheet', expectedRevision: seeded.view.revision,
    commands: [{ sheet: 'Notes', operation: 'addSheet' }] });
  const copied = await session.apply({ operationId: 'duplicate-sheet', expectedRevision: added.view.revision,
    commands: [{ sheet: 'Sheet1', operation: 'duplicateSheet', newName: 'Copy' }] });
  const renamed = await session.apply({ operationId: 'rename-sheet', expectedRevision: copied.view.revision,
    commands: [{ sheet: 'Notes', operation: 'renameSheet', newName: 'Details' }] });
  const reopened = await readCalculatedWorkbook(renamed.bytes);
  assert.deepEqual(reopened.sheets.map(sheet => sheet.name), ['Sheet1', 'Details', 'Copy']);
  assert.equal(cell(reopened, 'Sheet1!A1').value, 'Original');
  assert.equal(cell(reopened, 'Copy!A1').value, 'Original');
  assert.equal(Object.keys(reopened.sheets.find(sheet => sheet.name === 'Details').cells).length, 0);
  const originalParts = unzipSync(seeded.bytes), deliveredParts = unzipSync(renamed.bytes);
  for (const [name, bytes] of Object.entries(originalParts))
    if (!['xl/workbook.xml', 'xl/_rels/workbook.xml.rels', '[Content_Types].xml'].includes(name))
      assert.equal(hash(deliveredParts[name]), hash(bytes), `unrelated original part changed: ${name}`);
  const python = `import sys,zipfile,xml.etree.ElementTree as E
z=zipfile.ZipFile(sys.argv[1]); ns={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
w=E.fromstring(z.read('xl/workbook.xml')); sheets=w.find('m:sheets',ns)
r=E.fromstring(z.read('xl/_rels/workbook.xml.rels'))
t=E.fromstring(z.read('[Content_Types].xml'))
assert len(sheets)==3
assert len({s.attrib['sheetId'] for s in sheets})==3
rels={x.attrib['Id']:x.attrib['Target'] for x in r}
for s in sheets:
 key=s.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
 path='xl/'+rels[key]
 assert path in z.namelist(),path
 assert any(x.attrib.get('PartName')=='/'+path for x in t),path
print('ok')`;
  const temporary = fs.mkdtempSync(path.join(os.tmpdir(), 'xlsx-sheet-relations-'));
  try {
    const file = path.join(temporary, 'sheet-operations.xlsx'); fs.writeFileSync(file, renamed.bytes);
    const result = spawnSync('python3', ['-c', python, file], { encoding: 'utf8' });
    assert.equal(result.status, 0, result.stderr);
  } finally { fs.rmSync(temporary, { recursive: true, force: true }); }
  const formula = await session.apply({ operationId: 'formula-before-rename', expectedRevision: renamed.view.revision,
    commands: [{ sheet: 'Sheet1', cell: 'B1', formula: '1+1' }] });
  await assert.rejects(session.apply({ operationId: 'refuse-rename-references', expectedRevision: formula.view.revision,
    commands: [{ sheet: 'Sheet1', operation: 'renameSheet', newName: 'Unsafe' }] }), /not qualified/i);
  assert.equal(hash(session.bytes), hash(formula.bytes));
  const rich = fixture('preservation-rich'), richSession = new WorkbookSession(rich), richView = await richSession.read();
  await assert.rejects(richSession.apply({ operationId: 'refuse-linked-duplicate', expectedRevision: richView.revision,
    commands: [{ sheet: 'Data', operation: 'duplicateSheet', newName: 'Unsafe Copy' }] }), /would lose|not qualified/i);
  assert.equal(hash(richSession.bytes), hash(rich));
  const formulaSource = fixture('pricing-schedule'), formulaSession = new WorkbookSession(formulaSource), formulaView = await formulaSession.read();
  const formulaAdd = await formulaSession.apply({ operationId: 'add-next-to-formulas', expectedRevision: formulaView.revision,
    commands: [{ sheet: 'Working Notes', operation: 'addSheet' }] });
  assert(formulaAdd.view.sheets.some(sheet => sheet.name === 'Working Notes'));
  assert.equal(cell(formulaAdd.view, 'Pricing!D2').calculation.value, cell(formulaView, 'Pricing!D2').calculation.value);
  return { sourceSha256: hash(source), resultSha256: hash(renamed.bytes), sheetCount: reopened.sheets.length };
});

gate('W5 merge and unmerge preserve the anchor and refuse occupied ranges', async () => {
  assert.equal(workbookCommand.safeParse({ sheet: 'Sheet1', operation: 'mergeCells', range: 'A1:B2' }).success, true);
  const source = createWorkbook(), session = new WorkbookSession(source), initial = await session.read();
  const seeded = await session.apply({ operationId: 'merge-seed', expectedRevision: initial.revision,
    commands: [{ sheet: 'Sheet1', cell: 'A1', value: 'Header' }] });
  const merged = await session.apply({ operationId: 'merge-safe', expectedRevision: seeded.view.revision,
    commands: [{ sheet: 'Sheet1', operation: 'mergeCells', range: 'A1:B2' },
      { type: 'formatRange', sheet: 'Sheet1', range: 'A1', format: { horizontalAlign: 'center' } }] });
  const reopened = await readCalculatedWorkbook(merged.bytes);
  assert.deepEqual(reopened.sheets[0].merges, ['A1:B2']);
  assert.equal(cell(reopened, 'Sheet1!A1').value, 'Header');
  assert.equal(cell(reopened, 'Sheet1!A1').horizontalAlign, 'center');
  const sheetXml = strFromU8(unzipSync(merged.bytes)['xl/worksheets/sheet1.xml']);
  const parsed = spawnSync('python3', ['-c', `import sys,xml.etree.ElementTree as E
r=E.fromstring(sys.stdin.read());n={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
c=r.find('m:mergeCells',n)
assert c is not None and c.attrib['count']=='1' and len(c)==1 and c[0].attrib['ref']=='A1:B2'
print('ok')`], { input: sheetXml, encoding: 'utf8' });
  assert.equal(parsed.status, 0, parsed.stderr);
  const unmerged = await session.apply({ operationId: 'unmerge-safe', expectedRevision: merged.view.revision,
    commands: [{ sheet: 'Sheet1', operation: 'unmergeCells', range: 'A1:B2' }] });
  assert.deepEqual((await readCalculatedWorkbook(unmerged.bytes)).sheets[0].merges, []);
  assert.equal(cell(unmerged.view, 'Sheet1!A1').value, 'Header');
  const occupied = await session.apply({ operationId: 'occupied-seed', expectedRevision: unmerged.view.revision,
    commands: [{ sheet: 'Sheet1', cell: 'B1', value: 'Keep' }] });
  await assert.rejects(session.apply({ operationId: 'refuse-occupied-merge', expectedRevision: occupied.view.revision,
    commands: [{ sheet: 'Sheet1', operation: 'mergeCells', range: 'A1:B2' }] }), /nonempty cell/i);
  assert.equal(hash(session.bytes), hash(occupied.bytes));
  return { sourceSha256: hash(source), mergedSha256: hash(merged.bytes), anchorPreserved: true };
});

gate('W5 conditional highlights persist as differential styles and live rules', async () => {
  const source = createWorkbook(), session = new WorkbookSession(source), initial = await session.read();
  const seeded = await session.apply({ operationId: 'conditional-seed', expectedRevision: initial.revision,
    commands: [{ type: 'setCells', sheet: 'Sheet1', cells: [
      { cell: 'A1', value: 2 }, { cell: 'A2', value: 5 }, { cell: 'A3', value: 5 }
    ] }] });
  const result = await session.apply({ operationId: 'conditional-rules', expectedRevision: seeded.view.revision,
    commands: [
      { sheet: 'Sheet1', operation: 'addConditionalRule', range: 'A1:A3', rule: 'greaterThan', threshold: 4, fillColor: 'FFF2CC' },
      { sheet: 'Sheet1', operation: 'addConditionalRule', range: 'A1:A3', rule: 'duplicateValues', fillColor: 'FFC7CE' }
    ] });
  const reopened = await readCalculatedWorkbook(result.bytes);
  assert.deepEqual(reopened.sheets[0].conditionalRules.map(({ range, rule, threshold, fillColor, priority }) => ({ range, rule, threshold, fillColor, priority })), [
    { range: 'A1:A3', rule: 'greaterThan', threshold: 4, fillColor: '#FFF2CC', priority: 1 },
    { range: 'A1:A3', rule: 'duplicateValues', threshold: undefined, fillColor: '#FFC7CE', priority: 2 }
  ]);
  assert.deepEqual(['A1', 'A2', 'A3'].map(ref => reopened.sheets[0].cells[ref].value), [2, 5, 5]);
  const parts = unzipSync(result.bytes);
  const inspected = spawnSync('python3', ['-c', `import sys,json,xml.etree.ElementTree as E
d=json.load(sys.stdin);n={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
s=E.fromstring(d['styles']);w=E.fromstring(d['sheet'])
dxfs=s.find('m:dxfs',n);assert dxfs is not None and int(dxfs.attrib['count'])==2 and len(dxfs)==2
rules=w.findall('m:conditionalFormatting/m:cfRule',n)
assert len(rules)==2
assert [r.attrib['type'] for r in rules]==['cellIs','duplicateValues']
assert [r.attrib['dxfId'] for r in rules]==['0','1']
assert [r.attrib['priority'] for r in rules]==['1','2']
assert rules[0].find('m:formula',n).text=='4'
print('ok')`], { input: JSON.stringify({ styles: strFromU8(parts['xl/styles.xml']), sheet: strFromU8(parts['xl/worksheets/sheet1.xml']) }), encoding: 'utf8' });
  assert.equal(inspected.status, 0, inspected.stderr);
  await assert.rejects(session.apply({ operationId: 'invalid-conditional', expectedRevision: result.view.revision,
    commands: [{ sheet: 'Sheet1', operation: 'addConditionalRule', range: 'A1:A3', rule: 'greaterThan', fillColor: 'FFF2CC' }] }), /Invalid or oversized/i);
  assert.equal(hash(session.bytes), hash(result.bytes));
  return { sourceSha256: hash(source), savedSha256: hash(result.bytes), rules: 2 };
});

gate('W5 formatted table creation writes a valid relationship, headers and stripe style', async () => {
  assert.equal(workbookCommand.safeParse({ sheet: 'Sheet1', operation: 'createTable', range: 'A1:B3', tableName: 'Items' }).success, true);
  const source = createWorkbook(), session = new WorkbookSession(source), initial = await session.read();
  const seeded = await session.apply({ operationId: 'table-seed', expectedRevision: initial.revision,
    commands: [{ type: 'setCells', sheet: 'Sheet1', cells: [
      { cell: 'A1', value: 'Item' }, { cell: 'B1', value: 'Qty' },
      { cell: 'A2', value: 'Paper' }, { cell: 'B2', value: 3 },
      { cell: 'A3', value: 'Ink' }, { cell: 'B3', value: 5 }
    ] }] });
  const result = await session.apply({ operationId: 'create-table', expectedRevision: seeded.view.revision,
    commands: [{ sheet: 'Sheet1', operation: 'createTable', range: 'A1:B3', tableName: 'Items' }] });
  const reopened = await readCalculatedWorkbook(result.bytes);
  assert.deepEqual(reopened.tables.map(table => ({ name: table.name, ref: table.ref, columns: table.columns })),
    [{ name: 'Items', ref: 'A1:B3', columns: [{ name: 'Item' }, { name: 'Qty' }] }]);
  assert.deepEqual(reopened.sheets[0].tableRanges, [{ name: 'Items', ref: 'A1:B3' }]);
  assert.equal(cell(reopened, 'Sheet1!B3').value, 5);
  const parts = unzipSync(result.bytes), sourceParts = unzipSync(seeded.bytes);
  for (const [name, bytes] of Object.entries(sourceParts))
    if (!['xl/worksheets/sheet1.xml', 'xl/workbook.xml', '[Content_Types].xml'].includes(name))
      assert.equal(hash(parts[name]), hash(bytes), `unrelated original part changed: ${name}`);
  const inspected = spawnSync('python3', ['-c', `import sys,json,xml.etree.ElementTree as E
d=json.load(sys.stdin);n={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
w=E.fromstring(d['sheet']); t=E.fromstring(d['table']); r=E.fromstring(d['rels']); c=E.fromstring(d['types'])
parts=w.find('m:tableParts',n);assert parts is not None and parts.attrib['count']=='1' and len(parts)==1
assert t.attrib['name']=='Items' and t.attrib['ref']=='A1:B3'
assert [x.attrib['name'] for x in t.find('m:tableColumns',n)]==['Item','Qty']
assert t.find('m:tableStyleInfo',n).attrib['name']=='TableStyleMedium2'
assert any(x.attrib['Type'].endswith('/table') for x in r)
assert any(x.attrib.get('PartName')=='/xl/tables/table1.xml' for x in c)
print('ok')`], { input: JSON.stringify({ sheet: strFromU8(parts['xl/worksheets/sheet1.xml']),
    table: strFromU8(parts['xl/tables/table1.xml']), rels: strFromU8(parts['xl/worksheets/_rels/sheet1.xml.rels']),
    types: strFromU8(parts['[Content_Types].xml']) }), encoding: 'utf8' });
  assert.equal(inspected.status, 0, inspected.stderr);
  await assert.rejects(session.apply({ operationId: 'duplicate-table', expectedRevision: result.view.revision,
    commands: [{ sheet: 'Sheet1', operation: 'createTable', range: 'A1:B3', tableName: 'Items2' }] }), /overlaps/i);
  assert.equal(hash(session.bytes), hash(result.bytes));
  return { sourceSha256: hash(source), resultSha256: hash(result.bytes), table: 'Items' };
});

gate('G3 independent package oracle rejects lost formulas, opaque bytes, styles and wrong table/name shifts', async () => {
  const original = fixture('pricing-schedule'), negativeChecks = [];
  for (const [name, part, change, request] of [
    ['formula', 'xl/worksheets/sheet1.xml', text => text.replace('B2*C2', 'B2+C2'), {}],
    ['unrelated-value', 'xl/worksheets/sheet1.xml', text => text.replace('<v>80</v>', '<v>81</v>'), { targets: { Pricing: ['B2'] } }],
    ['style', 'xl/styles.xml', text => text.replace('name val="Calibri"', 'name val="Other"'), { appendStyles: true }],
    ['table-range', 'xl/tables/table1.xml', text => text.replaceAll('A1:F3', 'A1:F4'), { expectedTables: { 'xl/tables/table1.xml': 'A1:F5' } }],
    ['defined-name', 'xl/workbook.xml', text => text.replace('Rates!$B$1', 'Rates!$B$2'), { expectedNames: { TaxRate: 'Rates!$B$3' } }],
  ]) {
    const parts = unzipSync(original); parts[part] = strToU8(change(strFromU8(parts[part])));
    const report = independentlyCheck(original, zipSync(parts), request, false);
    negativeChecks.push({ name, rejected: !report.passed, failures: report.failures });
  }
  const rich = fixture('preservation-rich'), parts = unzipSync(rich); parts['xl/media/image1.png'][0] ^= 1;
  const opaque = independentlyCheck(rich, zipSync(parts), {}, false); negativeChecks.push({ name: 'opaque-media', rejected: !opaque.passed, failures: opaque.failures });
  const missingShift = independentlyCheck(original, original, { expectedTables: { 'xl/tables/table1.xml': 'A1:F4' } }, false);
  negativeChecks.push({ name: 'missing-table-shift', rejected: !missingShift.passed, failures: missingShift.failures });
  return { negativeChecks };
});

gate('G3 workbook protection and unsupported structural edits refuse without changing accepted bytes', async () => {
  for (const id of ['protected', 'preservation-rich']) {
    const original = fixture(id), session = new WorkbookSession(original), view = await session.read();
    const commands = id === 'protected' ? keys[id].refusedScenario.edits : [{ sheet: 'Data', operation: 'insertRow', index: 2 }];
    await assert.rejects(session.apply({ operationId: 'must-refuse', expectedRevision: view.revision, commands }), /Read-only|preserve|qualified|unsupported/i);
    assert.equal(hash(session.bytes), hash(original));
  }
  return { protectedEditsRefused: true, unsafeRichStructuralEditsRefused: true };
});

function rowOperationFixture(mixed = false, duplicate = false) {
  const parts = unzipSync(fixture('budget'));
  const header = '<row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Qty</t></is></c><c r="C1" t="inlineStr"><is><t>Price</t></is></c><c r="D1" t="inlineStr"><is><t>Total</t></is></c></row>';
  const rows = [['B', 2, 10], ['A', 3, 7], [duplicate ? 'A' : 'C', 1, 4]].map(([name, qty, price], index) => {
    const row = index + 2, formula = mixed ? '$B2*C$2' : `B${row}*C${row}`;
    return `<row r="${row}"><c r="A${row}" t="inlineStr"><is><t>${name}</t></is></c><c r="B${row}"><v>${qty}</v></c><c r="C${row}"><v>${price}</v></c>${mixed && row !== 2 ? '' : `<c r="D${row}" s="2"><f>${formula}</f><v>${qty * price}</v></c>`}</row>`;
  }).join('');
  parts['xl/worksheets/sheet1.xml'] = strToU8(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${header}${rows}</sheetData></worksheet>`);
  return zipSync(parts);
}

gate('G5 multi-column sort uses stable secondary and primary keys in one accepted batch', async () => {
  const original = rowOperationFixture(false, true), session = new WorkbookSession(original), view = await session.read();
  const result = await session.apply({ operationId: 'two-key-sort', expectedRevision: view.revision, commands: [
    { sheet: 'Budget', operation: 'sortRows', range: 'A2:D4', column: 'B', direction: 'desc' },
    { sheet: 'Budget', operation: 'sortRows', range: 'A2:D4', column: 'A', direction: 'asc' },
  ] });
  const expected = { 'Budget!A2': { value: 'A' }, 'Budget!B2': { value: 3 }, 'Budget!D2': { formula: 'B2*C2', calculated: 21 },
    'Budget!A3': { value: 'A' }, 'Budget!B3': { value: 1 }, 'Budget!D3': { formula: 'B3*C3', calculated: 4 },
    'Budget!A4': { value: 'B' }, 'Budget!B4': { value: 2 }, 'Budget!D4': { formula: 'B4*C4', calculated: 20 } };
  checkView(result.view, expected);
  const delivered = independentlyCheck(original, result.bytes, { rowMappings: { Budget: { 2: 4, 3: 2, 4: 3 } }, expected });
  return { sourceSha256: hash(original), delivered };
});

gate('G5 sorting row formulas, filtering without deletion, column width and frozen panes persist', async () => {
  const original = rowOperationFixture(), session = new WorkbookSession(original), view = await session.read();
  const result = await session.apply({ operationId: 'row-display', expectedRevision: view.revision, commands: [
    { sheet: 'Budget', operation: 'sortRows', range: 'A2:D4', column: 'A', direction: 'asc' },
    { sheet: 'Budget', operation: 'filterRows', range: 'A1:D4', column: 'A', values: ['A'] },
    { sheet: 'Budget', operation: 'resizeColumn', index: 1, width: 32 },
    { sheet: 'Budget', operation: 'resizeRow', index: 1, height: 42 },
    { sheet: 'Budget', operation: 'freeze', rows: 1, columns: 1 },
  ] });
  const expected = { 'Budget!A2': { value: 'A' }, 'Budget!D2': { formula: 'B2*C2', calculated: 21 },
    'Budget!A3': { value: 'B' }, 'Budget!D3': { formula: 'B3*C3', calculated: 20 },
    'Budget!A4': { value: 'C' }, 'Budget!D4': { formula: 'B4*C4', calculated: 4 } };
  checkView(result.view, expected);
  const metadata = { hiddenRows: [3, 4], filter: { ref: 'A1:D4', columnId: '0', values: ['A'] },
    columns: [{ min: '2', max: '2', width: '32', customWidth: '1' }],
    rowHeights: [{ row: '2', height: '42', customHeight: '1' }],
    pane: { xSplit: '1', ySplit: '1', topLeftCell: 'B2', state: 'frozen', activePane: 'bottomRight' } };
  const delivered = independentlyCheck(original, result.bytes, { rowMappings: { Budget: { 2: 3, 3: 2 } }, expectedSheetMetadata: { Budget: metadata }, expected });
  const formulaFiltered = await session.apply({ operationId: 'formula-filter', expectedRevision: result.view.revision, commands: [{ sheet: 'Budget', operation: 'filterRows', range: 'A1:D4', column: 'D', values: [21] }] });
  const filtered = independentlyCheck(result.bytes, formulaFiltered.bytes, { expectedSheetMetadata: { Budget: { ...metadata, filter: { ref: 'A1:D4', columnId: '3', values: ['21'] } } }, expected });
  const restored = await session.apply({ operationId: 'clear-view-controls', expectedRevision: formulaFiltered.view.revision, commands: [{ sheet: 'Budget', operation: 'filterRows', range: 'A1:D4', column: 'D', values: [] }, { sheet: 'Budget', operation: 'freeze', rows: 0, columns: 0 }] });
  const cleared = independentlyCheck(formulaFiltered.bytes, restored.bytes, { expectedSheetMetadata: { Budget: { ...metadata, hiddenRows: [], filter: { ref: 'A1:D4', columnId: null, values: [] }, pane: null } }, expected });
  return { privateVariantSha256: hash(original), delivered, filtered, cleared };
});

gate('G5 formula fill shifts mixed references and retains source number style', async () => {
  const original = rowOperationFixture(true), session = new WorkbookSession(original), view = await session.read();
  const result = await session.apply({ operationId: 'fill-mixed', expectedRevision: view.revision, commands: [{ type: 'fillRange', sheet: 'Budget', source: 'D2', range: 'D3:D4' }] });
  const expected = { 'Budget!D3': { formula: '$B3*C$2', calculated: 30, style: '2' }, 'Budget!D4': { formula: '$B4*C$2', calculated: 10, style: '2' } };
  checkView(result.view, expected);
  const delivered = independentlyCheck(original, result.bytes, { targets: { Budget: ['D3', 'D4'] }, expected });
  return { privateVariantSha256: hash(original), delivered };
});

gate('G5 Paste Formulas shifts references without copying literal cells or source styles', async () => {
  assert.equal(workbookCommand.safeParse({ type: 'pasteFormulas', sheet: 'Sheet1', source: 'A1:B1', range: 'A3:B3' }).success, true);
  const source = createWorkbook(), session = new WorkbookSession(source), initial = await session.read();
  const seeded = await session.apply({ operationId: 'formula-paste-seed', expectedRevision: initial.revision,
    commands: [{ type: 'setCells', sheet: 'Sheet1', cells: [
      { cell: 'A1', value: 2 }, { cell: 'B1', formula: 'A1*2' }, { cell: 'A3', value: 7 }
    ] }, { type: 'formatRange', sheet: 'Sheet1', range: 'B3', format: { fillColor: 'FFC7CE' } }] });
  const result = await session.apply({ operationId: 'formula-paste', expectedRevision: seeded.view.revision,
    commands: [{ type: 'pasteFormulas', sheet: 'Sheet1', source: 'A1:B1', range: 'A3:B3' }] });
  const reopened = await readCalculatedWorkbook(result.bytes);
  assert.equal(cell(reopened, 'Sheet1!A3').value, 7);
  assert.equal(cell(reopened, 'Sheet1!B3').formula, 'A3*2');
  assert.equal(cell(reopened, 'Sheet1!B3').calculation.value, 14);
  assert.equal(cell(reopened, 'Sheet1!B3').fillColor, 'FFC7CE');
  await assert.rejects(session.apply({ operationId: 'mismatched-formula-paste', expectedRevision: result.view.revision,
    commands: [{ type: 'pasteFormulas', sheet: 'Sheet1', source: 'A1:B1', range: 'A4:A5' }] }), /matching source and target/i);
  assert.equal(hash(session.bytes), hash(result.bytes));
  return { sourceSha256: hash(source), resultSha256: hash(result.bytes) };
});

gate('G2 filling exact numeric storage preserves its digits or refuses atomically', async () => {
  const original = fixture('calculation-states'), session = new WorkbookSession(original), view = await session.read();
  let result;
  try { result = await session.apply({ operationId: 'fill-precision', expectedRevision: view.revision, commands: [{ type: 'fillRange', sheet: 'Diagnostics', source: 'E1', range: 'H1' }] }); }
  catch (error) {
    assert.match(String(error), /precision|unsafe|stored numeric/i);
    assert.equal(hash(session.bytes), hash(original));
    return { fixtureSha256: hash(original), outcome: 'explicit refusal before accepted state change', reason: String(error) };
  }
  assert.equal(cell(result.view, 'Diagnostics!H1').rawValue, '9007199254740993');
  const delivered = independentlyCheck(original, result.bytes, { targets: { Diagnostics: ['H1'] }, expected: { 'Diagnostics!H1': { rawValue: '9007199254740993' } } });
  return { fixtureSha256: hash(original), outcome: 'exact stored digits retained', delivered };
});

gate('W4 keyed updates do not match a rounded numeric key to a different stored identity', async () => {
  const original = fixture('calculation-states'), session = new WorkbookSession(original), view = await session.read();
  await assert.rejects(session.apply({ operationId: 'rounded-key', expectedRevision: view.revision, commands: [{
    type: 'updateRows', sheet: 'Diagnostics', range: 'E1:H1', keyColumn: 'E',
    updates: [{ key: 9007199254740992, cells: [{ column: 'H', expectedValue: null, value: 'must-not-match' }] }],
  }] }), /precision|unsafe|stored numeric/i);
  assert.equal(hash(session.bytes), hash(original));
  return { fixtureSha256: hash(original), storedKey: '9007199254740993', requestedKey: 9007199254740992, refusedWithoutChangingBytes: true };
});

if (process.env.OFFICEMCP_REAL_WORKBOOK) gate('G1.3 real cost-cell update preserves unrelated members and unsupported formula definitions', async () => {
  const file = path.resolve(process.env.OFFICEMCP_REAL_WORKBOOK), original = fs.readFileSync(file), session = new WorkbookSession(original), view = await session.read();
  const result = await session.apply({ operationId: 'real-cost', expectedRevision: view.revision, commands: [{ sheet: '项目成本', cell: 'D4', expectedRawValue: '254063.62', value: 254064.62 }] });
  const delivered = independentlyCheck(original, result.bytes, { targets: { 项目成本: ['D4'] }, expected: { '项目成本!D4': { value: 254064.62, rawValue: '254064.62', style: '127', formula: null } } });
  assert.equal(hash(fs.readFileSync(file)), hash(original));
  assert.equal(result.calculation.unresolved > 0, true);
  await assert.rejects(session.apply({ operationId: 'unsafe-structure', expectedRevision: result.view.revision, commands: [{ sheet: '项目成本', operation: 'insertRow', index: 3 }] }), /preserve|qualified|unsupported/i);
  return { sourceSha256: hash(original), calculation: result.calculation, delivered, realPricingWorkflow: 'Not demonstrated: this is an existing cost-cell edit, not an existing quantity/price workflow.' };
});

else test('G1.3 real cost-cell update preserves unrelated members and unsupported formula definitions', {skip:'Set OFFICEMCP_REAL_WORKBOOK to an authorized private workbook path.'}, () => {});

gate('W1/W4/W5 MCP query, keyed price edit, durable retry, bounded evidence and save/reopen', async () => {
  const s = createServiceFixture({ prefix: 'xlsx-oss-workflow-', limits: { workerTimeoutMs: 60000 } });
  const calls = [];
  const instrument = service => {
    const call = service.call.bind(service);
    service.call = async (name, args) => {
      const start = performance.now(), result = await call(name, args);
      calls.push({ name, elapsedMs: performance.now() - start, returnedJsonBytes: Buffer.byteLength(JSON.stringify(result)), isError: Boolean(result.isError) });
      return result;
    };
  };
  instrument(s.service);
  let restarted;
  try {
    const original = fixture('pricing-schedule'), file = path.join(s.docs, 'pricing.xlsx'); fs.writeFileSync(file, original);
    const opened = await s.call('office_open', { path: file });
    const view = { documentId: opened.document.id, revision: opened.document.sourceRevision };
    const query = await s.call('xlsx_query', { view, sheet: 'Pricing', range: 'A1:F5' });
    assert.equal(query.data.items.find(item => item.cell === 'D2').calculation.value, 250);
    const quantityTarget = query.data.items.find(item => item.cell === 'B2').target;
    assert.equal(quantityTarget.documentId, view.documentId);
    const otherFile = path.join(s.docs, 'identical-other-document.xlsx'); fs.writeFileSync(otherFile, original);
    const other = await s.call('office_open', { path: otherFile });
    const otherView = { documentId: other.document.id, revision: other.document.sourceRevision };
    assert.notEqual(otherView.documentId, view.documentId); assert.equal(otherView.revision, view.revision);
    const crossDocument = await s.service.call('xlsx_apply', { view: otherView, operationId: 'wrong-document-target', commands: [{ target: quantityTarget, expectedValue: 2, value: 3 }] });
    assert.equal(crossDocument.isError, true);
    const quantity = await s.call('xlsx_apply', { view, operationId: 'quantity', commands: [{ target: quantityTarget, expectedValue: 2, value: 3 }] });
    const staleTarget = await s.service.call('xlsx_apply', { view: quantity.view, operationId: 'stale-target', commands: [{ target: quantityTarget, expectedValue: 3, value: 4 }] });
    assert.equal(staleTarget.isError, true);
    const freshTarget = quantity.data.affectedTargets.find(item => item.cell === 'B2').target;
    assert.equal(freshTarget.documentId, quantity.view.documentId);
    assert.equal(freshTarget.candidateId, quantity.view.candidateId);
    assert.equal(freshTarget.revision, quantity.view.revision);
    const failed = await s.service.call('xlsx_apply', { view: quantity.view, operationId: 'bad-batch', commands: [{ sheet: 'Pricing', cell: 'B2', value: 9 }, { sheet: 'Pricing', cell: 'C2', expectedValue: 999, value: 150 }] });
    assert.equal(failed.isError, true);
    const unchanged = await s.call('xlsx_query', { view: quantity.view, sheet: 'Pricing', range: 'B2:B2' });
    assert.equal(unchanged.data.items[0].value, 3);
    const args = { view: quantity.view, operationId: 'keyed-price', commands: [{ type: 'updateRows', sheet: 'Pricing', range: 'A2:F3', keyColumn: 'A', updates: [{ key: '00123', cells: [{ column: 'C', expectedValue: 125, value: 150 }] }] }] };
    const edited = await s.call('xlsx_apply', args);
    assert(Buffer.byteLength(JSON.stringify(edited)) <= 262144);
    const stale = await s.service.call('xlsx_apply', { ...args, operationId: 'stale' });
    assert.equal(stale.isError, true);
    const staleRead = await s.service.call('xlsx_query', { view: quantity.view, sheet: 'Pricing', range: 'B2:B2' });
    assert.equal(staleRead.isError, true);
    s.service.close(); restarted = new OfficeService(configFrom(s.config)); instrument(restarted);
    const retry = await restarted.call('xlsx_apply', args);
    assert.equal(retry.isError, undefined);
    assert.deepEqual(retry.payload.view, edited.view);
    const read = await restarted.call('xlsx_query', { view: edited.view, sheet: 'Pricing', range: 'A1:F5' });
    assert(!read.isError, JSON.stringify(read.payload));
    for (const [target, expected] of Object.entries(keys['pricing-schedule'].scenario.expected)) near(read.payload.data.items.find(item => item.cell === target.split('!')[1]).calculation.value, expected, target);
    const saved = await restarted.call('office_save', { view: edited.view, operationId: 'deliver', destination: { rootId: 'output', relativePath: 'delivered.xlsx' } });
    assert(!saved.isError, JSON.stringify(saved.payload));
    const output = fs.readFileSync(saved.payload.data.path);
    const delivered = independentlyCheck(original, output, { targets: targetsFor(keys['pricing-schedule'].scenario.edits), expected: expectedValues(keys['pricing-schedule'].scenario.expected) });
    assert.equal(hash(fs.readFileSync(file)), hash(original));
    return { fixtureSha256: hash(original), calls, totalCalls: calls.length, returnedJsonBytes: calls.reduce((sum, item) => sum + item.returnedJsonBytes, 0), mutationReturnedBytes: Buffer.byteLength(JSON.stringify(edited)), delivered, guiParity: 'Not measured here; separate browser/native journey required.' };
  } finally { restarted?.close(); s.close(); }
});

gate('W5 MCP create, guarded edit and delivery work without an optional engine', async () => {
  const s = createServiceFixture({ prefix: 'xlsx-oss-create-', limits: { workerTimeoutMs: 60000 } });
  try {
    const created = await s.call('office_create', { operationId: 'create-budget', document: { format: 'xlsx', sheets: [{ name: 'Budget', rows: [['Item', 'Amount'], ['Services', 125], ['Total', { formula: 'SUM(B2:B2)' }]], header: true, table: false }] } });
    const sourceFile = s.service.view({ documentId: created.view.documentId, candidateId: created.view.candidateId, expectedRevision: created.view.revision }).file;
    const original = fs.readFileSync(sourceFile);
    const edited = await s.call('xlsx_apply', { view: created.view, operationId: 'edit-created', commands: [{ sheet: 'Budget', cell: 'B2', expectedValue: 125, value: 150 }] });
    const saved = await s.call('office_save', { view: edited.view, operationId: 'save-created', destination: { rootId: 'output', relativePath: 'created-budget.xlsx' } });
    const delivered = independentlyCheck(original, fs.readFileSync(saved.data.path), { targets: { Budget: ['B2'] }, expected: { 'Budget!B2': { value: 150 }, 'Budget!B3': { formula: 'SUM(B2:B2)', calculated: 150 } } });
    return { createdSha256: hash(original), delivered, optionalEngineConfigured: false };
  } finally { s.close(); }
});

gate('G2 MCP named-table queries stay inside the selected table and paginate reusable targets', async () => {
  const s = createServiceFixture({ prefix: 'xlsx-table-query-', limits: { workerTimeoutMs: 60000 } });
  try {
    const original = fixture('pricing-schedule'), file = path.join(s.docs, 'table.xlsx'); fs.writeFileSync(file, original);
    const opened = await s.call('office_open', { path: file }), view = { documentId: opened.document.id, revision: opened.document.sourceRevision };
    const whole = await s.call('xlsx_query', { view, table: 'PricingItems', limit: 5 });
    assert.equal(whole.data.total, 18); assert.equal(whole.data.items.length, 5); assert.equal(whole.data.nextOffset, 5);
    const next = await s.call('xlsx_query', { view, table: 'PricingItems', offset: whole.data.nextOffset, limit: 5 });
    assert.equal(next.data.items[0].cell, 'F1');
    for (const item of [...whole.data.items, ...next.data.items]) assert.equal(item.target.documentId, view.documentId);
    const inside = await s.call('xlsx_query', { view, table: 'PricingItems', range: 'B2:D2' });
    assert.equal(inside.data.items.length, 3); assert.equal(inside.data.items.at(-1).calculation.value, 250);
    for (const args of [{ table: 'PricingItems', range: 'A1:F4' }, { table: 'PricingItems', sheet: 'Rates' }, { table: 'MissingTable' }]) {
      const refused = await s.service.call('xlsx_query', { view, ...args }); assert.equal(refused.isError, true);
    }
    return { fixtureSha256: hash(original), totalTableCells: 18, pageSize: 5, insideRangeCells: 3, outsideRangeRefused: true, wrongSheetRefused: true, missingTableRefused: true };
  } finally { s.close(); }
});

gate('G4.3 MCP bounds large change evidence and recovers durable receipts, fresh targets and saved artifacts', async () => {
  const s = createServiceFixture({ prefix: 'xlsx-evidence-coverage-', limits: { workerTimeoutMs: 60000 } });
  let restarted;
  const call = async (service, name, args) => { const result = await service.call(name, args); assert(!result.isError, JSON.stringify(result.payload)); return result.payload; };
  try {
    const original = fixture('scale-10000'), file = path.join(s.docs, 'large-change.xlsx'); fs.writeFileSync(file, original);
    const opened = await s.call('office_open', { path: file }), view = { documentId: opened.document.id, revision: opened.document.sourceRevision };
    const read = await s.call('xlsx_query', { view, sheet: 'Scale', range: 'A1:A101', limit: 100 });
    assert.equal(read.data.total, 101); assert.equal(read.data.items.length, 100); assert.equal(read.data.nextOffset, 100);
    assert(!read.data.items.some(item => item.cell === 'A101'), 'the unread tail is not manufactured as read evidence');
    assert.equal(read.data.coverage.completeDocument, false);
    for (const omission of ['unrequested-ranges', 'external-data-freshness', 'visual-layout']) assert(read.data.coverage.omissions.includes(omission));
    const cells = Array.from({ length: 101 }, (_, index) => ({ cell: `A${index + 1}`, expectedValue: index + 1, value: index + 1001 }));
    const edited = await s.call('xlsx_apply', { view, operationId: 'large-edit', commands: [{ type: 'setCells', sheet: 'Scale', cells }] });
    assert.equal(edited.data.postconditions.affectedCellCount, 202);
    assert.equal(edited.data.postconditions.affectedCellsTruncated, true);
    assert.equal(edited.data.affectedTargets.length, 100);
    assert(Buffer.byteLength(JSON.stringify(edited)) <= 262144);
    const target = edited.data.affectedTargets.find(item => item.cell === 'A1').target;
    assert.deepEqual(target, { sheet: 'Scale', cell: 'A1', ...edited.view });
    const diagnosticsFile = path.join(s.docs, 'unsupported.xlsx'); fs.writeFileSync(diagnosticsFile, fixture('calculation-states'));
    const diagnostics = await s.call('office_open', { path: diagnosticsFile });
    const unresolved = await s.call('xlsx_query', { view: { documentId: diagnostics.document.id, revision: diagnostics.document.sourceRevision }, sheet: 'Diagnostics', range: 'C1:C2' });
    for (const item of unresolved.data.items) assert.equal(item.calculation.status, 'unresolved');
    s.service.close(); restarted = new OfficeService(configFrom(s.config));
    const recovered = await call(restarted, 'office_operation_status', { operationId: 'large-edit' });
    assert.equal(recovered.operationId, 'large-edit'); assert.equal(recovered.data.operationId, 'large-edit');
    assert.equal(recovered.data.data.applied, true);
    assert.equal(recovered.data.data.candidateId, edited.view.candidateId);
    assert.equal(recovered.data.data.candidateRevision, edited.view.revision);
    assert.equal(recovered.data.data.postconditions.affectedCellsTruncated, true);
    const tail = await call(restarted, 'xlsx_query', { view: edited.view, sheet: 'Scale', range: 'A1:A101', offset: 100, limit: 100 });
    assert.equal(tail.data.items.length, 1); assert.equal(tail.data.items[0].cell, 'A101'); assert.equal(tail.data.items[0].value, 1101);
    const next = await call(restarted, 'xlsx_apply', { view: edited.view, operationId: 'fresh-target-edit', commands: [{ target, expectedValue: 1001, value: 2001 }] });
    const saved = await call(restarted, 'office_save', { view: next.view, operationId: 'large-save', destination: { rootId: 'output', relativePath: 'delivered-large.xlsx' } });
    assert.equal(saved.data.published, true); assert.equal(saved.data.path, fs.realpathSync(path.join(s.docs, 'delivered-large.xlsx')));
    const output = fs.readFileSync(saved.data.path); assert.equal(saved.data.outputRevision, `sha256:${hash(output)}`);
    restarted.close(); restarted = new OfficeService(configFrom(s.config));
    const savedReceipt = await call(restarted, 'office_operation_status', { operationId: 'large-save' });
    assert.equal(savedReceipt.data.data.path, saved.data.path); assert.equal(savedReceipt.data.data.outputRevision, saved.data.outputRevision);
    const expected = Object.fromEntries(cells.flatMap(({ cell }, index) => [[`Scale!${cell}`, { value: index === 0 ? 2001 : index + 1001 }], [`Scale!J${index + 1}`, { calculated: index === 0 ? 2003 : 2 * (index + 1) + 1001 }]]));
    const delivered = independentlyCheck(original, output, { targets: { Scale: cells.map(item => item.cell) }, expected });
    assert.equal(hash(fs.readFileSync(file)), hash(original));
    return { fixtureSha256: hash(original), affectedCount: 202, returnedAffectedTargets: 100, mutationReturnedBytes: Buffer.byteLength(JSON.stringify(edited)), coverage: read.data.coverage, unreadTailOffset: read.data.nextOffset, unresolvedFormulaStates: unresolved.data.items.map(item => item.calculation.status), durableEditReceiptRecovered: true, postEditTargetReused: true, durableSavedArtifactRecovered: true, delivered };
  } finally { restarted?.close(); s.close(); }
});

after(() => {
  const fixtureManifestSha256 = hash(fs.readFileSync(path.join(corpus, 'generated.json')));
  if (transientCorpus) fs.rmSync(transientCorpus, { recursive: true, force: true });
  if (!process.env.RPF_XLSX_EVIDENCE) return;
  const output = path.resolve(repository, process.env.RPF_XLSX_EVIDENCE);
  const report = { schema: 'rpf.xlsx.shared-workflow-gates.v1', measuredAt: new Date().toISOString(), source, sourceStableDuringRun: source.sha256 === sourceIdentity().sha256,
    build: { kind: 'source tests', bundleSha256: null }, machine: { node: process.version, platform: os.platform(), release: os.release(), arch: os.arch(), cpu: os.cpus()[0]?.model },
    fixtureManifestSha256, procedure: 'node --test --test-concurrency=1 OfficeMCP/test/xlsx-shared-workflows.test.mjs', cases: reports,
    limitations: ['No native editor paint, dirty-buffer lifecycle or undo/redo measured by this harness.', 'Synthetic pricing does not establish the first real pricing workbook milestone.', 'Reference answers are hand-authored; delivered ZIP/XML is checked independently with Python stdlib.'] };
  fs.mkdirSync(path.dirname(output), { recursive: true }); fs.writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
});
