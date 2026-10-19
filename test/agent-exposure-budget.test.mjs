import test from './support/test.mjs';
import assert from 'node:assert/strict';
import { exposureBudget } from '../scripts/test/agent-exposure-budget.mjs';
test('UTF-8 result boundaries withhold oversized or binary content instead of truncating it', () => {
  const budget = exposureBudget({ perReply: 1024, total: 2048 });
  assert.equal(budget.expose('safe'), 'safe');
  const original = '敏感'.repeat(200);
  const refused = budget.expose(original);
  assert.equal(JSON.parse(refused).clientError, 'MODEL_EXPOSURE_BUDGET');
  assert(!refused.includes('敏感')); assert.equal(budget.state.refusedReplies, 1);
  assert(!budget.expose('pixels', { content: [{ type: 'image', data: 'pixels' }] }).includes('pixels'));
  assert(budget.state.sentBytes <= 2048);
});
test('cumulative budget includes error messages and fails closed when no response can fit', () => {
  const budget = exposureBudget({ perReply: 1024, total: 1100 });
  budget.expose('x'.repeat(1000));
  assert.throws(() => budget.expose('y'.repeat(200)), /exhausted/);
  assert.equal(budget.state.sentBytes, 1000);
});
