import test from './support/test.mjs';
import assert from 'node:assert/strict';
import { requireAcceptedCorpus } from '../scripts/test/docx-corpus-report.mjs';
const consumer = { passed: true, productIdentity: 'product', sdk: { version: 'sdk' } };
const passing = () => ({ passed: true, finished: 'now', productIdentity: 'product', sdk: { version: 'sdk' },
  runs: [{ document: 'input', job: 'edit', passed: true, evidence: { passed: true } }] });
test('stored pass flags cannot override a failed or missing independent document verdict', () => {
  assert.equal(requireAcceptedCorpus(passing(), consumer), 1);
  const failed = passing(); failed.runs[0].evidence.passed = false;
  assert.throws(() => requireAcceptedCorpus(failed, consumer), /independent evidence/);
  delete failed.runs[0].evidence; assert.throws(() => requireAcceptedCorpus(failed, consumer), /independent evidence/);
});
test('corpus acceptance rejects incomplete runs, runtime mismatch and breached model budgets', () => {
  const incomplete = passing(); delete incomplete.finished; assert.throws(() => requireAcceptedCorpus(incomplete, consumer), /incomplete/);
  const mismatch = passing(); mismatch.productIdentity = 'other'; assert.throws(() => requireAcceptedCorpus(mismatch, consumer), /differ/);
  const oversized = passing(); Object.assign(oversized.runs[0], { model: 'agent', budget: { sentBytes: 260000 }, steps: [{ modelOutputTextBytes: 40000 }] });
  assert.throws(() => requireAcceptedCorpus(oversized, consumer), /Per-reply/);
  oversized.runs[0].budget.sentBytes = 300000; assert.throws(() => requireAcceptedCorpus(oversized, consumer), /Cumulative/);
});
