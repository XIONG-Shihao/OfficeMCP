import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { Client } from '@modelcontextprotocol/client';
import { StdioClientTransport } from '@modelcontextprotocol/client/stdio';
import { stdioClient, toolPayload } from './docx-mcp-client.mjs';
import { productFixture, PRODUCT_JOBS } from './docx-product-fixture.mjs';
import { hash, productIdentity, verifyJob, inspectPackage } from './docx-product-evidence.mjs';

export async function runWorkflows({ sourceRoot, output, versions = ['2026-07-28', '2025-11-25', 'official-modern'] }) {
  fs.mkdirSync(output, { recursive: true });
  const fixture = await productFixture(), report = { passed: false, sourceRoot, productIdentity: productIdentity(sourceRoot), platform: process.platform + '-' + process.arch,
    scope: 'Synthetic DOCX jobs through public stdio only; independent saved-package checks. No model, rendering or universal document-shape claim.', profiles: [] };
  const jobs = [...PRODUCT_JOBS, { id: 'formatted-clause-unscoped' }, { id: 'formatted-clause-domain-target' }];
  for (const version of versions) {
    const root = path.join(output, version); fs.mkdirSync(root, { recursive: true });
    const docs = path.join(root, 'docs'); fs.mkdirSync(docs, { recursive: true });
    const source = path.join(docs, 'input.docx'); fs.writeFileSync(source, fixture);
    const config = path.join(root, 'config.json');
    fs.writeFileSync(config, JSON.stringify({ toolset: 'docx-agent', readRoots: [docs], writeRoots: { output: docs }, stateDirectory: path.join(root, 'state') }));
    const env = { ...process.env, NODE_PATH: '', RPF_OFFICEMCP_HOST_REQUIRED: '1', OFFICEMCP_PUBLICATION_ADAPTER: path.join(root, 'must-not-import.mjs') };
    const createClient = () => {
      if (version === 'official-modern') {
        const client = new Client({ name: 'independent-product-check', version: '1' }, { capabilities: {}, versionNegotiation: { mode: { pin: '2026-07-28' } } });
        const transport = new StdioClientTransport({ command: process.execPath, args: [path.join(sourceRoot, 'src/cli.mjs'), 'serve', '--config', config], cwd: root, env, stderr: 'pipe' });
        transport.stderr?.resume();
        return { connect: async () => { await client.connect(transport); return (await client.listTools()).tools; },
          callTool: (name, args) => client.callTool({ name, arguments: args }), close: () => client.close() };
      }
      return stdioClient({ cli: path.join(sourceRoot, 'src/cli.mjs'), config, cwd: root, version,
      // Tripwires prove stale private app configuration is not a dependency.
        env });
    };
    fs.writeFileSync(path.join(root, 'must-not-import.mjs'), 'throw Error("Private host adapter imported");');
    let client = createClient(), currentSteps;
    const profile = { version, jobs: [] }; report.profiles.push(profile);
    async function raw(name, args) {
      const start = performance.now(), response = await client.callTool(name, args), payload = toolPayload(response);
      currentSteps?.push({ tool: name, elapsedMs: performance.now() - start, inputJsonBytes: Buffer.byteLength(JSON.stringify(args)),
        outputJsonBytes: Buffer.byteLength(JSON.stringify(payload)), isError: response.isError === true, errorCode: payload.error?.code });
      return { response, payload };
    }
    async function call(name, args) { const r = await raw(name, args); assert(!r.response.isError, JSON.stringify(r.payload)); return r.payload; }
    async function target(view, select) { const r = await call('docx_query', { view, select, require: 'exactlyOne', offset: 0, limit: 10 }); assert.equal(r.data.items.length, 1); return r.data.items[0].target; }
    try {
      const tools = await client.connect(); assert.equal(tools.length, 14); assert(tools.every(t => t.name.startsWith('docx_')));
      const help = await call('docx_inspect', {}); assert.equal(help.contractVersion, 'docx/6');
      assert(!/rpf_|integrated app|codex/i.test(help.data.usage.toolPrefix));
      for (const job of jobs) {
        const record = { id: job.id, passed: false, steps: [] }; profile.jobs.push(record); currentSteps = record.steps;
        let view = job.id === 'save-recovery'
          ? (await call('docx_create', { operationId: 'unchanged-copy', action: 'from_template', path: source })).view
          : (await call('docx_open', { path: source })).view;
        try {
          if (job.id === 'ambiguous-refusal') {
            const refusal = await raw('docx_apply', { view, operationId: job.id, action: 'rewrite_block',
              select: { type: 'text', pattern: 'Duplicate risk' }, arguments: { text: 'Reviewed risk', changeMode: 'direct' } });
            assert(refusal.response.isError); assert.equal(refusal.payload.error.code, 'TARGET_CONFLICT');
            const still = await call('docx_query', { view, select: { type: 'text', pattern: 'Duplicate risk' }, require: 'all' });
            assert.equal(still.data.total, 2); assert(!fs.existsSync(path.join(docs, job.id + '.docx')));
            record.refusalCode = refusal.payload.error.code;
          } else {
            if (job.id === 'formatted-clause-domain-target') {
              const paragraphs = await call('docx_read', { view, representation: 'domain', domain: 'paragraphs', offset: 0, limit: 10 });
              const selected = paragraphs.data.items.find(p => p.text === 'Payment within 30 days.').target;
              view = (await call('docx_apply', { view, operationId: job.id, action: 'replace_text', selectorTarget: selected,
                arguments: { edits: [{ find: '30 days', replace: '45 days' }], caseSensitive: true, changeMode: 'direct' } })).view;
            } else if (job.id === 'formatted-clause-unscoped') {
              view = (await call('docx_apply', { view, operationId: job.id, action: 'replace_text',
                arguments: { edits: [{ find: 'Payment within 30 days.', replace: 'Payment within 45 days.' }], caseSensitive: false, changeMode: 'direct' } })).view;
            } else if (job.id === 'formatted-clause') {
              const selected = await target(view, { type: 'text', pattern: 'Payment within 30 days.' });
              view = (await call('docx_apply', { view, operationId: job.id, action: 'replace_text', selectorTarget: selected,
                arguments: { edits: [{ find: '30', replace: '45' }], caseSensitive: true, changeMode: 'direct' } })).view;
              const stale = await raw('docx_apply', { view, operationId: 'stale-target', action: 'replace_text', selectorTarget: selected,
                arguments: { edits: [{ find: '45', replace: '60' }], changeMode: 'direct' } });
              assert(stale.response.isError); record.staleTargetRefused = stale.payload.error.code;
              await target(view, { type: 'text', pattern: 'Payment within 45 days.' });
            } else if (job.id === 'table-by-key') {
              const selected = await target(view, { type: 'table', header: 'Price USD' });
              const wrong = await raw('docx_apply', { view, operationId: 'wrong-cell-value', selectorTarget: selected,
                workflow: { type: 'update_table_by_key', keyColumn: 'Service', updates: [{ key: 'Support', column: 'Price USD', expectedText: '999', text: '150' }] } });
              assert(wrong.response.isError); record.wrongValueRefused = wrong.payload.error.code;
              view = (await call('docx_apply', { view, operationId: job.id, selectorTarget: selected,
                workflow: { type: 'update_table_by_key', keyColumn: 'Service', updates: [{ key: 'Support', column: 'Price USD', expectedText: '125', text: '150' }] } })).view;
            } else if (job.id === 'template-control') {
              const selected = await target(view, { type: 'entity', domain: 'content_controls', tag: 'client' });
              view = (await call('docx_apply', { view, operationId: job.id, action: 'set_content_control_text', selectorTarget: selected,
                arguments: { expectedText: 'Original client', text: 'Meridian Labs' } })).view;
            } else if (job.id === 'comment-read') {
              const comments = await call('docx_read', { view, representation: 'domain', domain: 'comments', limit: 20 });
              assert(comments.data.items.some(c => c.text === 'Approved by legal' && c.status === 'open'));
              await call('docx_close', { view, disposition: 'discard' });
              view = (await call('docx_create', { operationId: 'comment-copy', action: 'from_template', path: source })).view;
            }
            if (view.candidateId) { const validation = await call('docx_validate', { view }); assert.equal(validation.data.passed, true); }
            const saveArgs = { view, operationId: job.id + '-save', destination: { rootId: 'output', relativePath: job.id + '.docx' } };
            if (job.id === 'save-recovery') {
              // Complete publication, intentionally discard response, disconnect and restart.
              await call('docx_save', saveArgs); await client.close(); client = createClient(); await client.connect();
              const recovered = await call('docx_operation_get', { operationId: saveArgs.operationId });
              assert.equal(hash(fs.readFileSync(recovered.data.path)), recovered.data.outputRevision.replace(/^sha256:/, ''));
              record.recovery = 'Completed Save response deliberately discarded; ordinary CLI restarted; durable receipt and exact output recovered.';
            }
            const saved = await call('docx_save', saveArgs), retry = await call('docx_save', saveArgs);
            assert.equal(saved.data.outputRevision, retry.data.outputRevision); assert.equal(saved.data.path, retry.data.path);
            record.outputHash = hash(fs.readFileSync(saved.data.path)); assert.equal('sha256:' + record.outputHash, saved.data.outputRevision);
            record.evidence = verifyJob(job.id.startsWith('formatted-clause') ? 'formatted-clause' : job.id, fixture, fs.readFileSync(saved.data.path));
            const reopened = (await call('docx_open', { path: saved.data.path })).view;
            if (job.id.startsWith('formatted-clause')) await target(reopened, { type: 'text', pattern: 'Payment within 45 days.' });
            if (job.id === 'table-by-key') {
              const table = await call('docx_read', { view: reopened, representation: 'domain', domain: 'tables', limit: 10 });
              assert.equal(table.data.items[0].rows[1][1].text, '150');
            }
            if (job.id === 'template-control') {
              const controls = await call('docx_read', { view: reopened, representation: 'domain', domain: 'content_controls', limit: 10 });
              assert(controls.data.items.some(c => c.text === 'Meridian Labs'));
            }
            await call('docx_close', { view: reopened, disposition: 'discard' });
          }
          assert.equal(hash(fs.readFileSync(source)), hash(fixture)); record.passed = true;
        } catch (error) { record.error = String(error.stack ?? error); }
        finally { await call('docx_close', { view, disposition: 'discard' }).catch(error => { record.closeError = String(error); record.passed = false; }); }
        console.log(`${record.passed ? 'PASS' : 'FAIL'} ${version} ${job.id}`);
      }
      currentSteps = undefined;
      const blank = await call('docx_create', { operationId: 'blank-create', document: { format: 'docx', blocks: [{ kind: 'paragraph', text: '' }] } });
      const saved = await call('docx_save', { view: blank.view, operationId: 'blank-save', destination: { rootId: 'output', relativePath: 'blank.docx' } });
      const blankPackage = inspectPackage(fs.readFileSync(saved.data.path));
      assert(blankPackage.paragraphs.every(p => p.text === '')); assert.equal(blankPackage.tables.length, 0);
      await call('docx_close', { view: blank.view, disposition: 'discard' }); profile.blankCreate = true;
    } catch (error) { profile.error = String(error.stack ?? error); }
    finally { await client.close(); }
  }
  report.passed = report.profiles.every(p => !p.error && p.blankCreate && p.jobs.length === jobs.length && p.jobs.every(j => j.passed));
  fs.writeFileSync(path.join(output, 'workflows.json'), JSON.stringify(report, null, 2) + '\n');
  return report;
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const sourceRoot = path.resolve(process.env.OFFICEMCP_DOCX_SOURCE ?? fileURLToPath(new URL('../../', import.meta.url)));
  const output = path.resolve(process.env.OFFICEMCP_DOCX_OUTPUT ?? path.join(sourceRoot, 'artifacts/docx-workflows', String(Date.now())));
  const report = await runWorkflows({ sourceRoot, output });
  console.log(JSON.stringify({ passed: report.passed, report: path.join(output, 'workflows.json') })); process.exitCode = report.passed ? 0 : 1;
}
