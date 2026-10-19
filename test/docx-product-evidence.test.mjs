import test from './support/test.mjs';
import assert from 'node:assert/strict';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { productFixture } from '../scripts/test/docx-product-fixture.mjs';
import { verifyJob } from '../scripts/test/docx-product-evidence.mjs';
import { queryAgentDocument, resolveInlineLocator } from '../src/formats/docx/docx-agent-query.mjs';
import { ambiguityEvidence } from '../scripts/test/docx-agent-evidence.mjs';

test('independent clause evidence rejects correct text with formatting, cell, footer or media damage', async () => {
  const original = await productFixture(), parts = unzipSync(original);
  parts['word/document.xml'] = strToU8(strFromU8(parts['word/document.xml']).replace('<w:t xml:space="preserve">30</w:t>', '<w:t xml:space="preserve">45</w:t>'));
  const valid = zipSync(parts);
  assert(verifyJob('formatted-clause', original, valid).passed);
  for (const corrupt of [
    p => { p['word/document.xml'] = strToU8(strFromU8(p['word/document.xml']).replace('<w:b/>', '')); },
    p => { p['word/document.xml'] = strToU8(strFromU8(p['word/document.xml']).replace('<w:t xml:space="preserve">300</w:t>', '<w:t xml:space="preserve">301</w:t>')); },
    p => { const n = Object.keys(p).find(n => /^word\/footer/.test(n)); p[n] = strToU8(strFromU8(p[n]).replace('500', '750')); },
    p => { const n = Object.keys(p).find(n => n.startsWith('word/media/') && !n.endsWith('/')); p[n] = new Uint8Array([1, 2, 3]); },
    p => { p['word/document.xml'] = strToU8(strFromU8(p['word/document.xml']).replace('cx="152400"', 'cx="152401"')); },
  ]) {
    const damaged = unzipSync(valid); corrupt(damaged);
    assert.throws(() => verifyJob('formatted-clause', original, zipSync(damaged)), /Only the intended/);
  }
});

test('zero and multiple text/node matches return target conflicts before exposing a mutation locator', async () => {
  for (const total of [0, 2]) {
    const requests = [], doc = { query: { match: async args => {
      requests.push(args); assert.equal(args.require, 'any');
      return { total, items: [] };
    } } };
    await assert.rejects(queryAgentDocument(doc, { select: { type: 'text', pattern: 'Duplicate' }, require: 'exactlyOne', offset: 0, limit: 10 }), { code: 'TARGET_CONFLICT' });
    await assert.rejects(resolveInlineLocator(doc, { action: 'rewrite_block', select: { type: 'text', pattern: 'Duplicate' } }), { code: 'TARGET_CONFLICT' });
    assert.equal(requests.length, 2);
  }
});

test('an agent may safely refuse after reading two distinct matches without attempting an edit', () => {
  const steps = [{ tool: 'docx_read', payload: { data: { items: [
    { kind: 'paragraph', itemId: 'a', story: 'body', text: 'Duplicate risk' },
    { kind: 'paragraph', itemId: 'b', story: 'body', text: 'Duplicate risk' },
  ] } } }];
  const evidence = ambiguityEvidence(steps, 'Refused: two identical paragraphs make the request ambiguous.');
  assert(evidence.passed); assert(!evidence.serverRefusalObserved); assert.equal(evidence.distinctReadMatches, 2);
  assert(!ambiguityEvidence([...steps, { tool: 'docx_apply', isError: false }], 'ambiguous').passed);
});
test('repeated reads of one match or an unsupported model claim do not establish ambiguity', () => {
  const read = { tool: 'docx_read', payload: { data: { items: [{ kind: 'paragraph', itemId: 'a', story: 'body', text: 'Duplicate risk' }] } } };
  assert(!ambiguityEvidence([read, read], 'two matches').passed);
  assert(!ambiguityEvidence([], 'ambiguous').passed);
});
