/** Generic public MCP + model loop; no app integration or provider credentials. */
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { agentLoop } from './docx-agent-workflows.mjs';
import { corpusClient, fileHash, verifyCorpusJob } from './docx-corpus-workflows.mjs';
import { exposureBudget } from './agent-exposure-budget.mjs';
import { productIdentity, sdkIdentity } from './docx-product-evidence.mjs';
import { toolMetrics } from '../benchmark-tool-metrics.mjs';

const project = fileURLToPath(new URL('../../', import.meta.url));
const representativeJobs = {
  'bid-business': ['text-edit', 'save-recovery'], 'bid-pricing': ['table-edit'],
  'bid-stress': ['text-edit', 'table-edit', 'save-recovery'],
  'public-nda': ['text-edit'], 'public-controls': ['template-control'], 'public-comments': ['comment-read'],
};
async function main() {
  assert.equal(process.env.OFFICEMCP_AGENT_LIVE, '1', 'Live model tests require OFFICEMCP_AGENT_LIVE=1');
  const apiKey = process.env.OFFICEMCP_AGENT_API_KEY, baseUrl = process.env.OFFICEMCP_AGENT_BASE_URL;
  assert(apiKey && baseUrl, 'Supply provider settings through environment variables');
  const models = (process.env.OFFICEMCP_AGENT_MODELS ?? '').split(',').filter(Boolean); assert(models.length);
  const sourceRoot = path.resolve(process.env.OFFICEMCP_DOCX_SOURCE ?? project);
  const corpus = JSON.parse(fs.readFileSync(process.env.OFFICEMCP_DOCX_CORPUS));
  const output = path.resolve(process.env.OFFICEMCP_DOCX_OUTPUT); fs.mkdirSync(output, { recursive: true });
  const reportFile = path.join(output, 'agents.json'); assert(!fs.existsSync(reportFile), 'Use a fresh evidence directory');
  const report = { passed: false, started: new Date().toISOString(), productIdentity: productIdentity(sourceRoot), sdk: sdkIdentity(sourceRoot),
    harness: 'Generic public stdio MCP + Chat Completions; official modern client; selected applicable jobs, one run per model/job.',
    corpus: process.env.OFFICEMCP_DOCX_CORPUS, models, baseUrl, requestOptions: JSON.parse(process.env.OFFICEMCP_AGENT_REQUEST_OPTIONS ?? '{}'),
    scope: 'One real bid and its two excerpts; supplemental public synthetic NDA, control and review fixtures. No population or cross-harness success-rate claim.', runs: [] };
  const selectedDocuments = process.env.OFFICEMCP_DOCX_DOCUMENTS?.split(','), selectedJobs = process.env.OFFICEMCP_AGENT_JOBS?.split(',');
  const persist = () => fs.writeFileSync(reportFile, JSON.stringify(report, null, 2) + '\n');
  for (const model of models) for (const doc of corpus.documents.filter(d => !selectedDocuments || selectedDocuments.includes(d.id))) {
    assert(/^[a-zA-Z0-9_.-]+$/.test(model));
    if (doc.classification.startsWith('real')) {
      assert.equal(process.env.OFFICEMCP_CORPUS_REAL_EXCERPTS_ALLOWED, '1', 'Private bid excerpts require explicit caller authorization');
      assert(['deepseek-flash', 'deepseek-v4-pro'].includes(model), 'Bid excerpt authorization covers only the two selected DeepSeek models');
    }
    const jobs = doc.jobs.filter(j => (selectedJobs ?? representativeJobs[doc.id] ?? []).includes(j.id));
    for (const job of jobs) {
      const root = path.join(output, model, doc.id, job.id), docs = path.join(root, 'output'); fs.mkdirSync(docs, { recursive: true });
      const config = path.join(root, 'config.json');
      fs.writeFileSync(config, JSON.stringify({ toolset: 'docx-agent', readRoots: [path.dirname(doc.file), docs], writeRoots: { output: docs },
        stateDirectory: path.join(root, 'state'), limits: { workerTimeoutMs: 300000 } }));
      const createClient = () => corpusClient(sourceRoot, config, root, 'official-modern');
      let client = createClient(), saveFaultInjected = false;
      const budget = exposureBudget(), record = { model, document: doc.id, classification: doc.classification, job: job.id, passed: false, budget: budget.state };
      report.runs.push(record); persist();
      try {
        assert.equal(await fileHash(doc.file), doc.sourceHash);
        const tools = await client.connect();
        const result = await agentLoop({ client: { callTool: (name, args) => client.callTool(name, args) }, tools, apiKey, baseUrl, model,
          timeoutMs: 900000, maxTurns: 48, requestOptions: report.requestOptions,
          instruction: `Work only on ${doc.file}. ${job.instruction} Save the result as reviewed.docx under write rootId output. Do not overwrite the input. Use bounded targeted reads; at most 32 KiB per tool reply and 256 KiB new tool results can be exposed. Use the advertised MCP tools only. Report the precise change or the existing comments you read.`,
          onProgress: progress => { Object.assign(record, progress); persist(); },
          afterTool: async ({ step, args, output: text, result: raw }) => {
            if (job.id === 'save-recovery' && step.tool === 'docx_save' && !step.isError && !saveFaultInjected) {
              saveFaultInjected = true; await client.close(); client = createClient(); await client.connect();
              text = JSON.stringify({ clientError: 'Save response unavailable; server restarted. Recover the existing operation by operationId before any further save.', operationId: args.operationId });
            }
            return budget.expose(text, raw);
          },
        });
        Object.assign(record, result, { saveFaultInjected, toolCalls: result.steps.length, toolErrors: result.steps.filter(s => s.isError).length,
          repeatedRequests: toolMetrics(result.steps.map(s => ({ part: { tool: s.tool, state: { input: s.arguments } } }))).repeatedRequests,
          inputJsonBytes: result.steps.reduce((n, s) => n + s.inputJsonBytes, 0), modelOutputTextBytes: budget.state.sentBytes,
          mcpElapsedMs: result.steps.reduce((n, s) => n + s.elapsedMs, 0), providerElapsedMs: result.apiSteps.reduce((n, s) => n + s.elapsedMs, 0) });
        // Release the heavy SDK session before the independent ZIP/XML audit.
        await client.close();
        const saved = path.join(docs, 'reviewed.docx'); assert(fs.existsSync(saved), 'No saved output');
        record.outputHash = await fileHash(saved);
        const receipt = result.steps.find(s => s.tool === 'docx_save' && !s.isError && s.payload?.data?.outputRevision === 'sha256:' + record.outputHash);
        assert(receipt, 'No matching public Save receipt');
        if (job.id === 'save-recovery') {
          assert(saveFaultInjected, 'No Save response loss injected');
          assert(result.steps.some(s => s.tool === 'docx_operation_get' && !s.isError && s.arguments.operationId === receipt.arguments.operationId
            && s.payload?.data?.outputRevision === receipt.payload.data.outputRevision), 'No durable matching save receipt recovery');
        }
        assert.deepEqual(fs.readdirSync(docs).filter(n => n.endsWith('.docx')), ['reviewed.docx'], 'Unexpected duplicate output');
        record.evidence = await verifyCorpusJob(doc.file, saved, job, root);
        assert(record.evidence.passed, 'Independent saved-document mismatch: ' + JSON.stringify(record.evidence.mismatches));
        if (job.id === 'comment-read') {
          const readItems = result.steps.filter(s => !s.isError && s.tool === 'docx_read' && s.arguments.domain === 'comments').flatMap(s => s.payload?.data?.items ?? []);
          for (const comment of doc.inventory.comments) {
            const item = readItems.find(c => c.text === comment.text); assert(item, 'An existing comment was not read');
            assert.equal(item.status, comment.resolved ? 'resolved' : 'open');
            assert(result.finalText.includes(comment.text), 'An existing comment was not reported');
          }
        }
        assert.equal(await fileHash(doc.file), doc.sourceHash, 'Source document changed'); record.passed = true;
      } catch (error) { record.error = String(error.stack ?? error).replaceAll(apiKey, '[REDACTED]'); }
      finally { await client.close().catch(() => {}); }
      persist(); console.log(`${record.passed ? 'PASS' : 'FAIL'} ${model} ${doc.id} ${job.id}${record.error ? ': ' + record.error.split('\n')[0] : ''}`);
    }
  }
  report.passed = report.runs.length > 0 && report.runs.every(r => r.passed); report.finished = new Date().toISOString(); persist();
  console.log(JSON.stringify({ passed: report.passed, cases: report.runs.length, report: reportFile })); process.exitCode = report.passed ? 0 : 1;
}
if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) await main();
