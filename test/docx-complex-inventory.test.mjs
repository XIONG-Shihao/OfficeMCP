import test from './support/test.mjs';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { Document, Packer, Paragraph, Math as WordMath, MathRun } from 'docx';
import { DOMParser } from '@xmldom/xmldom';
import { complexObjectInventory } from '../src/formats/docx/docx-complex-inventory.mjs';
import { createServiceFixture } from './service-fixture.mjs';
test('complex inventory reads cached chart values, textbox text, equations and embedded identities with declared guarded OPC targets', () => {
  const parser = new DOMParser();
  const chart = parser.parseFromString('<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><a:t>Sales</a:t></c:title><c:plotArea><c:barChart><c:ser><c:tx><c:v>Revenue</c:v></c:tx><c:cat><c:strCache><c:pt><c:v>Q1</c:v></c:pt></c:strCache></c:cat><c:val><c:numCache><c:pt><c:v>125</c:v></c:pt></c:numCache></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>', 'application/xml');
  const body = parser.parseFromString('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:body><w:p><w:txbxContent><w:p><w:r><w:t>Existing textbox</w:t></w:r></w:p></w:txbxContent><m:oMath><m:r><m:t>x+y</m:t></m:r></m:oMath></w:p></w:body></w:document>', 'application/xml');
  const items = complexObjectInventory([['word/charts/chart1.xml', chart], ['word/document.xml', body]], ['word/embeddings/oleObject1.bin']);
  assert.deepEqual(items.find(i => i.kind === 'chart').series[0].values, ['125']);
  assert.equal(items.find(i => i.kind === 'textbox').text, 'Existing textbox');
  assert.equal(items.find(i => i.kind === 'equation').text, 'x+y');
  assert(items.some(i => i.kind === 'embeddedObject' && !i.readable));
  assert(items.filter(i => ['chart','textbox','equation'].includes(i.kind)).every(i => i.mutationTargetAvailable && i.nativeTarget.kind === 'opc_object')); assert(items.filter(i => i.kind === 'embeddedObject').every(i => !i.nativeTarget));
});
test('public MCP reads and queries an existing equation with an opaque guarded object target', async () => {
  const f = createServiceFixture({ toolset: 'docx-agent' });
  try {
    const source = path.join(f.docs, 'equation.docx');
    fs.writeFileSync(source, await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph({ children: [new WordMath({ children: [new MathRun('x+y')] })] })] }] })));
    const before = fs.readFileSync(source), view = (await f.call('docx_open', { path: source })).view;
    const read = await f.call('docx_read', { view, representation: 'domain', domain: 'objects', limit: 20 });
    const equation = read.data.items.find(i => i.kind === 'equation'); assert(equation); assert.equal(equation.text, 'x+y'); assert(equation.target);
    const query = await f.call('docx_query', { view, select: { type: 'entity', domain: 'objects', id: equation.itemId }, require: 'exactlyOne' });
    assert.equal(query.data.items.length, 1); assert(query.data.items[0].target);
    assert.deepEqual(fs.readFileSync(source), before);
  } finally { f.close(); }
});
