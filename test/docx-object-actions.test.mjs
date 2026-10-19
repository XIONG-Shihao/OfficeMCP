import test from './support/test.mjs';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { Document, Packer, Paragraph, Math as WordMath, MathRun } from 'docx';
import { unzipSync, strFromU8 } from 'fflate';
import { createServiceFixture } from './service-fixture.mjs';
test('equation object edit previews, saves and reopens without changing any other member', async () => {
  const f = createServiceFixture({ toolset: 'docx-agent' });
  try {
    const source = path.join(f.docs, 'equation.docx');
    fs.writeFileSync(source, await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph('Untouched outside content.'), new Paragraph({ children: [new WordMath({ children: [new MathRun('x+y')] })] })] }] })));
    const original = fs.readFileSync(source), view = (await f.call('docx_open', { path: source })).view;
    const objects = await f.call('docx_read', { view, representation: 'domain', domain: 'objects' });
    const target = objects.data.items.find(item => item.kind === 'equation').target; assert(target);
    const request = { view, action: 'replace_equation_text', selectorTarget: target, arguments: { expectedText: 'x+y', find: '+', replace: '-' }, readback: { maxItems: 5, maxChars: 200 } };
    const preview = await f.call('docx_apply', { ...request, mode: 'preview' }); assert(preview.data.preview); assert.equal(f.service.rows('candidate').length, 0);
    const applied = await f.call('docx_apply', { ...request, operationId: 'equation-edit' });
    assert.equal(applied.data.engine, 'OfficeMCP guarded OOXML object adapter v1');
    assert.equal(applied.data.readback.items[0].after.text, 'x-y'); assert(applied.data.readback.items[0].target);
    const saved = await f.call('docx_save', { view: applied.view, operationId: 'save', destination: { rootId: 'output', relativePath: 'edited.docx' } });
    const before = unzipSync(original), after = unzipSync(fs.readFileSync(saved.data.path));
    for (const name of Object.keys(before)) if (name !== 'word/document.xml') assert.deepEqual(after[name], before[name], name);
    assert(strFromU8(after['word/document.xml']).includes('x-y')); assert(strFromU8(after['word/document.xml']).includes('Untouched outside content.'));
    const reopened = (await f.call('docx_open', { path: saved.data.path })).view;
    const read = await f.call('docx_read', { view: reopened, representation: 'domain', domain: 'objects' }); assert(read.data.items.some(item => item.text === 'x-y'));
    assert.deepEqual(fs.readFileSync(source), original);
  } finally { f.close(); }
});
test('later stale object edit rolls back the entire batch', async () => {
  const f = createServiceFixture({ toolset: 'docx-agent' });
  try {
    const source = path.join(f.docs, 'equation.docx'); fs.writeFileSync(source, await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph({ children: [new WordMath({ children: [new MathRun('x+y')] })] })] }] })));
    const view = (await f.call('docx_open', { path: source })).view, objects = await f.call('docx_read', { view, representation: 'domain', domain: 'objects' });
    const target = objects.data.items.find(item => item.kind === 'equation').target;
    const result = await f.service.call('docx_apply', { view, operationId: 'failed-batch', steps: [
      { action: 'replace_equation_text', selectorTarget: target, arguments: { expectedText: 'x+y', find: '+', replace: '-' } },
      { action: 'replace_equation_text', selectorTarget: target, arguments: { expectedText: 'wrong', find: 'x', replace: 'z' } },
    ] });
    assert(result.isError); assert.equal(result.payload.error.code, 'TARGET_CONFLICT'); assert.equal(f.service.rows('candidate').length, 0);
    const read = await f.call('docx_read', { view, representation: 'domain', domain: 'objects' }); assert(read.data.items.some(item => item.text === 'x+y'));
  } finally { f.close(); }
});
