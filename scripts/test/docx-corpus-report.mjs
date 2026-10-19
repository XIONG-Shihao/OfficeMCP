/** Fail-closed acceptance for stored independent corpus evidence. */
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
export function requireAcceptedCorpus(report, consumer) {
  assert(consumer.passed, 'Consumer installation did not pass');
  assert(report.passed && report.finished && report.runs.length > 0, 'Corpus run is incomplete or failed');
  assert.equal(report.productIdentity, consumer.productIdentity, 'Corpus and installed product differ');
  assert.deepEqual(report.sdk, consumer.sdk, 'Corpus and installed SDK differ');
  for (const run of report.runs) {
    assert(run.passed && run.evidence?.passed === true, `${run.document}/${run.job}: independent evidence did not pass`);
    if (run.model) {
      assert(run.budget?.sentBytes <= 256 * 1024, 'Cumulative model exposure exceeded');
      assert(run.steps.every(step => step.modelOutputTextBytes <= 32 * 1024), 'Per-reply model exposure exceeded');
    }
  }
  return report.runs.length;
}
if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const [consumerFile, ...reports] = process.argv.slice(2);
  assert(consumerFile && reports.length, 'Usage: node docx-corpus-report.mjs consumer.json corpus-report.json ...');
  const consumer = JSON.parse(fs.readFileSync(consumerFile));
  const accepted = reports.map(file => ({ file, acceptedCases: requireAcceptedCorpus(JSON.parse(fs.readFileSync(file)), consumer) }));
  console.log(JSON.stringify({ passed: true, productIdentity: consumer.productIdentity, accepted }, null, 2));
}
