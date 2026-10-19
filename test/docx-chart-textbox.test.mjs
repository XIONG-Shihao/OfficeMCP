import test from './support/test.mjs';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { unzipSync, strFromU8 } from 'fflate';
import ExcelJS from 'exceljs';
import { DOMParser } from '@xmldom/xmldom';
import { complexFixture } from './docx-complex-fixture.mjs';
import { createServiceFixture } from './service-fixture.mjs';
const tree = bytes => new DOMParser().parseFromString(strFromU8(bytes), 'application/xml');
async function open(f, options = {}) { const file = path.join(f.docs, 'complex.docx'), original = await complexFixture(options); fs.writeFileSync(file, original); return { file, original, view: (await f.call('docx_open', { path: file })).view }; }
async function object(f, view, kind) { return (await f.call('docx_read', { view, representation: 'domain', domain: 'objects' })).data.items.find(i => i.kind === kind).target; }
const data = { seriesName: 'Revenue', expectedValues: [10, 20], values: [15, 25], expectedCategories: ['Jan', 'Feb'], categories: ['March', 'April'] };
test('chart title, chart values/workbook and textbox edits survive public save, recovery and reopen', async () => {
  const f = createServiceFixture({ toolset: 'docx-agent' });
  try {
    const source = await open(f); let view = source.view;
    view = (await f.call('docx_apply', { view, operationId: 'title', action: 'set_chart_title', selectorTarget: await object(f, view, 'chart'), arguments: { expectedText: 'Original Sales', text: 'Revised Sales' }, readback: {} })).view;
    const preview = await f.call('docx_apply', { view, mode: 'preview', action: 'set_chart_data', selectorTarget: await object(f, view, 'chart'), arguments: data }); assert(preview.data.preview);
    view = (await f.call('docx_apply', { view, operationId: 'data', action: 'set_chart_data', selectorTarget: await object(f, view, 'chart'), arguments: data, readback: {} })).view;
    view = (await f.call('docx_apply', { view, operationId: 'textbox', action: 'replace_textbox_text', selectorTarget: await object(f, view, 'textbox'), arguments: { expectedText: 'Payment 30 days.Second paragraph.', find: '30', replace: '45' }, readback: {} })).view;
    const save = await f.call('docx_save', { view, operationId: 'save', destination: { rootId: 'output', relativePath: 'result.docx' } });
    const before = unzipSync(source.original), after = unzipSync(fs.readFileSync(save.data.path));
    for (const name of Object.keys(before)) if (!['word/document.xml', 'word/charts/chart1.xml', 'word/embeddings/data.xlsx'].includes(name)) assert.deepEqual(after[name], before[name], name);
    const chart = tree(after['word/charts/chart1.xml']); assert.equal(chart.getElementsByTagNameNS('*', 't')[0].textContent, 'Revised Sales');
    const vals = Array.from(chart.getElementsByTagNameNS('*', 'numCache')[0].getElementsByTagNameNS('*', 'v')).map(n => Number(n.textContent)); assert.deepEqual(vals, [15, 25]);
    const book = new ExcelJS.Workbook(); await book.xlsx.load(Buffer.from(after['word/embeddings/data.xlsx'])); const sheet = book.worksheets[0]; assert.equal(sheet.getCell('B2').value, 15); assert.equal(sheet.getCell('B3').value, 25); assert.equal(sheet.getCell('A2').value, 'March'); assert.equal(sheet.getCell('C2').value, 'Keep A'); assert.equal(sheet.getCell('B2').numFmt, '$0.00');
    const innerBefore = unzipSync(before['word/embeddings/data.xlsx']), innerAfter = unzipSync(after['word/embeddings/data.xlsx']); for (const name of Object.keys(innerBefore)) if (name !== 'xl/worksheets/sheet1.xml') assert.deepEqual(innerAfter[name], innerBefore[name], name);
    const document = tree(after['word/document.xml']); const box = document.getElementsByTagNameNS('*', 'txbxContent')[0]; assert.equal(box.textContent, 'Payment 45 days.Second paragraph.'); assert.equal(box.getElementsByTagNameNS('*', 'b').length, 1);
    const reopened = (await f.call('docx_open', { path: save.data.path })).view; const objects = (await f.call('docx_read', { view: reopened, representation: 'domain', domain: 'objects' })).data.items; assert(objects.some(i => i.kind === 'chart' && i.text === 'Revised Sales')); assert(objects.some(i => i.kind === 'textbox' && i.text.includes('45 days'))); assert.deepEqual(fs.readFileSync(source.file), source.original);
    assert((await f.call('docx_operation_get', { operationId: 'data' })).view);
  } finally { f.close(); }
});
for (const options of [{ formula: true }, { external: true }]) test('unsafe chart sources refuse without a candidate: ' + JSON.stringify(options), async () => {
  const f = createServiceFixture({ toolset: 'docx-agent' }); try { const { view } = await open(f, options); const result = await f.service.call('docx_apply', { view, operationId: 'bad', action: 'set_chart_data', selectorTarget: await object(f, view, 'chart'), arguments: data }); assert(result.isError); assert.equal(result.payload.error.code, 'UNSUPPORTED_FEATURE'); assert.equal(f.service.rows('candidate').length, 0); } finally { f.close(); }
});
test('literal chart data edits and stale/cross-paragraph refusals are atomic', async () => {
  const f = createServiceFixture({ toolset: 'docx-agent' }); try {
    const { view } = await open(f, { literal: true }), chart = await object(f, view, 'chart');
    const bad = await f.service.call('docx_apply', { view, operationId: 'bad', steps: [{ action: 'set_chart_title', selectorTarget: chart, arguments: { expectedText: 'Original Sales', text: 'Good' } }, { action: 'set_chart_data', selectorTarget: chart, arguments: { ...data, expectedValues: [0, 20] } }] }); assert(bad.isError); assert.equal(f.service.rows('candidate').length, 0);
    const crossing = await f.service.call('docx_apply', { view, operationId: 'cross', action: 'replace_textbox_text', selectorTarget: await object(f, view, 'textbox'), arguments: { expectedText: 'Payment 30 days.Second paragraph.', find: 'days.Second', replace: 'unsafe' } }); assert(crossing.isError); assert.equal(f.service.rows('candidate').length, 0);
    assert((await f.call('docx_apply', { view, operationId: 'good', action: 'set_chart_data', selectorTarget: chart, arguments: data })).view);
  } finally { f.close(); }
});
test('undeclared OPC members refuse an object mutation before publication',async()=>{
 const f=createServiceFixture({toolset:'docx-agent'});try{const original=await complexFixture(),parts=unzipSync(original);parts['[Content_Types].xml']=Buffer.from(strFromU8(parts['[Content_Types].xml']).replace(/<Default Extension="bin" ContentType="application\/octet-stream"\/>/,''));const {zipSync}=await import('fflate'),file=path.join(f.docs,'undeclared.docx');fs.writeFileSync(file,Buffer.from(zipSync(parts)));const view=(await f.call('docx_open',{path:file})).view;const result=await f.service.call('docx_apply',{view,operationId:'invalid',action:'set_chart_title',selectorTarget:await object(f,view,'chart'),arguments:{expectedText:'Original Sales',text:'Should refuse'}});assert(result.isError);assert.equal(result.payload.error.code,'INVALID_DOCUMENT');assert.equal(f.service.rows('candidate').length,0);
 }finally{f.close();}
});
test('inapplicable object filters and section scopes refuse rather than selecting unrelated content',async()=>{
 const f=createServiceFixture({toolset:'docx-agent'});try{const {view}=await open(f);const query=await f.service.call('docx_query',{view,select:{type:'entity',domain:'objects',tag:'does-not-exist'},require:'exactlyOne'});assert(query.isError);assert.equal(query.payload.error.code,'UNSUPPORTED_FEATURE');const read=await f.service.call('docx_read',{view,representation:'domain',domain:'objects',section:{headingPath:['Unrelated']}});assert(read.isError);assert.equal(read.payload.error.code,'INVALID_ARGUMENT');}finally{f.close();}
});
