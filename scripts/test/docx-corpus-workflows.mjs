/** Real/corpus jobs using only discovered public MCP tools and an independent oracle. */
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { createHash } from 'node:crypto';
import { Client } from '@modelcontextprotocol/client';
import { StdioClientTransport } from '@modelcontextprotocol/client/stdio';
import { stdioClient, toolPayload } from './docx-mcp-client.mjs';
import { productIdentity, sdkIdentity } from './docx-product-evidence.mjs';
const execute = promisify(execFile);
const project = fileURLToPath(new URL('../../', import.meta.url));
export const corpusOracle = fileURLToPath(new URL('docx_corpus_evidence.py', import.meta.url));
export async function fileHash(file) {
  const hash = createHash('sha256'); for await (const chunk of fs.createReadStream(file)) hash.update(chunk); return hash.digest('hex');
}
export async function verifyCorpusJob(source, saved, job, output) {
  const spec = path.join(output, 'expected.json'); fs.writeFileSync(spec, JSON.stringify(job.spec));
  const { stdout } = await execute(process.env.OFFICEMCP_PYTHON ?? 'python3', [corpusOracle, 'verify', source, saved, spec], { timeout: 120000, maxBuffer: 2 * 1024 ** 2 });
  const evidence = JSON.parse(stdout); fs.writeFileSync(path.join(output, 'independent.json'), JSON.stringify(evidence, null, 2) + '\n');
  return evidence;
}
export function corpusClient(sourceRoot, config, root, profile, timeoutMs = 300000) {
  if (profile !== 'official-modern') return stdioClient({ cli: path.join(sourceRoot, 'src/cli.mjs'), config, cwd: root, version: profile, timeoutMs });
  const client = new Client({ name: 'standalone-corpus-check', version: '1' }, { capabilities: {}, versionNegotiation: { mode: { pin: '2026-07-28' } } });
  const transport = new StdioClientTransport({ command: process.execPath, args: [path.join(sourceRoot, 'src/cli.mjs'), 'serve', '--config', config], cwd: root, stderr: 'pipe' });
  transport.stderr?.resume();
  return { connect: async () => { await client.connect(transport); return (await client.listTools()).tools; },
    callTool: (name, args) => client.callTool({ name, arguments: args }, undefined, { timeout: timeoutMs }), close: () => client.close() };
}

export async function runCorpus({ sourceRoot, manifest, output, profiles = ['2026-07-28', '2025-11-25', 'official-modern'], documents, jobs }) {
  fs.mkdirSync(output, { recursive: true }); assert(!fs.existsSync(path.join(output, 'corpus-workflows.json')), 'Choose a fresh evidence directory');
  const corpus = JSON.parse(fs.readFileSync(manifest));
  const report = { passed: false, started: new Date().toISOString(), sourceRoot, productIdentity: productIdentity(sourceRoot), sdk: sdkIdentity(sourceRoot),
    scope: 'Real supplied bid, two derived excerpts, supplemental public Word-style samples; fixed public MCP requests and independent saved ZIP/XML invariants. No app/renderer or universal fidelity claim.', manifest, runs: [] };
  for (const doc of corpus.documents.filter(d => !documents || documents.includes(d.id))) {
    // Stress shape is tested once with the official client. Ordinary documents
    // qualify the protocol matrix without repeating heavy SDK opens needlessly.
    for (const profile of doc.id === 'bid-stress' ? profiles.filter(p => p === 'official-modern') : profiles) {
      for (const job of doc.jobs.filter(j => !jobs || jobs.includes(j.id))) {
        const root = path.join(output, doc.id, profile, job.id), docs = path.join(root, 'output'); fs.mkdirSync(docs, { recursive: true });
        const config = path.join(root, 'config.json');
        fs.writeFileSync(config, JSON.stringify({ toolset: 'docx-agent', readRoots: [path.dirname(doc.file), docs], writeRoots: { output: docs }, stateDirectory: path.join(root, 'state'),
          limits: { workerTimeoutMs: 300000 } }));
        const createClient = () => corpusClient(sourceRoot, config, root, profile);
        let client = createClient(), view;
        const record = { document: doc.id, classification: doc.classification, profile, job: job.id, passed: false, steps: [] }; report.runs.push(record);
        const raw = async (name, args) => {
          const start = performance.now(), r = await client.callTool(name, args), p = toolPayload(r);
          record.steps.push({ tool: name, elapsedMs: performance.now() - start, inputJsonBytes: Buffer.byteLength(JSON.stringify(args)), outputJsonBytes: Buffer.byteLength(JSON.stringify(p)), isError: r.isError === true, error: p.error });
          return { r, p };
        };
        const call = async (name, args) => { const r = await raw(name, args); assert(!r.r.isError, JSON.stringify(r.p)); return r.p; };
        const query = async (select, withinTarget) => {
          const p = await call('docx_query', { view, select, ...(withinTarget ? { withinTarget } : {}), require: 'exactlyOne', offset: 0, limit: 2 });
          assert.equal(p.data.items.length, 1); return p.data.items[0];
        };
        try {
          assert.equal(await fileHash(doc.file), doc.sourceHash); await client.connect();
          view = job.id === 'save-recovery'
            ? (await call('docx_create', { operationId: 'copy', action: 'from_template', path: doc.file })).view
            : (await call('docx_open', { path: doc.file })).view;
          if (job.id === 'text-edit') {
            const q = await query(job.select); record.selectedText = q.text;
            view = (await call('docx_apply', { view, operationId: 'text-edit', action: 'replace_text', selectorTarget: q.target,
              arguments: { edits: [{ find: job.spec.find, replace: job.spec.replace }], caseSensitive: true, changeMode: 'direct' } })).view;
            await query({ type: 'text', pattern: job.select.pattern.replace(job.spec.find, job.spec.replace) });
          } else if (job.id === 'comment-read') {
            const comments = await call('docx_read', { view, representation: 'domain', domain: 'comments', offset: 0, limit: 20 });
            record.existingCommentsRead = comments.data;
            for (const expected of doc.inventory.comments ?? []) assert(comments.data.items.some(c => c.text === expected.text));
            await call('docx_close', { view, disposition: 'discard' });
            view = (await call('docx_create', { operationId: 'comment-copy', action: 'from_template', path: doc.file })).view;
          } else if (job.id === 'table-edit') {
            if (job.route === 'package-cell') {
              const id = `word/document.xml#table-${job.spec.tableIndex + 1}/row-${job.spec.row + 1}/cell-${job.spec.column + 1}`;
              const cell = await query({ type: 'entity', domain: 'table_cells', id });
              view = (await call('docx_apply', { view, operationId: 'table-cell-edit', action: 'replace_table_cell_text', selectorTarget: cell.target,
                arguments: { expectedText: job.spec.find, find: job.spec.find, replace: job.spec.replace } })).view;
              record.method = 'Explicit merged-owner package cell target; SDK range preflight then guarded adapter if needed.';
            } else {
            const table = await query(job.select); record.tableRows = table.rows ?? table.details?.rows;
            const keyed = await raw('docx_apply', { view, operationId: 'table-edit', selectorTarget: table.target,
              workflow: { type: 'update_table_by_key', keyColumn: job.keyColumn, updates: [{ key: job.key, column: job.columnHeader, expectedText: job.spec.find, text: job.spec.replace }] } });
            if (!keyed.r.isError) {
              view = keyed.p.view; record.method = 'Public keyed cell workflow, preserving unrelated merges.';
            } else {
              assert(['UNSUPPORTED_FEATURE', 'VALIDATION_FAILED'].includes(keyed.p.error.code)); record.keyedWorkflowRefusal = keyed.p.error.code;
              const cell = await query({ type: 'text', pattern: job.spec.find }, table.target);
              view = (await call('docx_apply', { view, operationId: 'cell-text-edit', action: 'replace_text', selectorTarget: cell.target,
                arguments: { edits: [{ find: job.spec.find, replace: job.spec.replace }], caseSensitive: true, changeMode: 'direct' } })).view;
              record.method = 'Unsupported keyed shape; unique table-scoped SDK literal range fallback.';
            }
            }
          } else if (job.id === 'template-control') {
            const q = await query({ type: 'entity', domain: 'content_controls', tag: job.spec.tag });
            const controls = doc.inventory.controls.find(c => c.tag === job.spec.tag);
            view = (await call('docx_apply', { view, operationId: 'control-edit', action: 'set_content_control_text', selectorTarget: q.target,
              arguments: { expectedText: controls.text, text: job.spec.replace } })).view;
          }
          const savedArgs = { view, operationId: 'save', destination: { rootId: 'output', relativePath: 'reviewed.docx' } };
          let saved = await call('docx_save', savedArgs);
          if (job.id === 'save-recovery') {
            saved = undefined; await client.close(); client = createClient(); await client.connect();
            saved = await call('docx_operation_get', { operationId: 'save' });
            const retry = await call('docx_save', savedArgs); assert.equal(saved.data.outputRevision, retry.data.outputRevision);
            record.recovery = 'Completed response discarded, CLI restarted, exact receipt recovered, same Save retried.';
          }
          const savedFile = saved.data.path; record.output = savedFile;
          assert.equal('sha256:' + await fileHash(savedFile), saved.data.outputRevision);
          await call('docx_close', { view, disposition: 'discard' }); view = undefined;
          record.evidence = await verifyCorpusJob(doc.file, savedFile, job, root);
          assert(record.evidence.passed, 'Independent preservation mismatch: ' + JSON.stringify(record.evidence.mismatches));
          view = (await call('docx_open', { path: savedFile })).view;
          if (job.id === 'text-edit') await query({ type: 'text', pattern: job.select.pattern.replace(job.spec.find, job.spec.replace) });
          if (job.id === 'template-control') { const c = await query({ type: 'entity', domain: 'content_controls', tag: job.spec.tag }); assert.equal(c.text, job.spec.replace); }
          assert.equal(await fileHash(doc.file), doc.sourceHash); record.passed = true;
        } catch (error) { record.error = String(error.stack ?? error); }
        finally { if (view) await call('docx_close', { view, disposition: 'discard' }).catch(() => {}); await client.close(); }
        fs.writeFileSync(path.join(output, 'corpus-workflows.json'), JSON.stringify(report, null, 2) + '\n');
        console.log(`${record.passed ? 'PASS' : 'FAIL'} ${doc.id} ${profile} ${job.id}${record.error ? ': ' + record.error.split('\n')[0] : ''}`);
      }
    }
  }
  report.passed = report.runs.length > 0 && report.runs.every(r => r.passed); report.finished = new Date().toISOString();
  fs.writeFileSync(path.join(output, 'corpus-workflows.json'), JSON.stringify(report, null, 2) + '\n'); return report;
}
if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const sourceRoot = path.resolve(process.env.OFFICEMCP_DOCX_SOURCE ?? project), manifest = path.resolve(process.env.OFFICEMCP_DOCX_CORPUS);
  const output = path.resolve(process.env.OFFICEMCP_DOCX_OUTPUT ?? path.join(project, 'artifacts/docx-corpus', String(Date.now())));
  const report = await runCorpus({ sourceRoot, manifest, output,
    profiles: process.env.OFFICEMCP_DOCX_PROFILES?.split(',') ?? undefined,
    documents: process.env.OFFICEMCP_DOCX_DOCUMENTS?.split(','), jobs: process.env.OFFICEMCP_DOCX_JOBS?.split(',') });
  console.log(JSON.stringify({ passed: report.passed, cases: report.runs.length, report: path.join(output, 'corpus-workflows.json') })); process.exitCode = report.passed ? 0 : 1;
}
