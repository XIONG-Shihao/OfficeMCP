import {spawnSync} from 'node:child_process';
import {fixturePath} from '../engine-source.mjs';
// Functional large-workbook probe through the production worker.
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { createHash } from 'node:crypto';
import { createServiceFixture } from '../../test/service-fixture.mjs';
const repository = path.resolve(import.meta.dirname, '../..');
const hash = bytes => createHash('sha256').update(bytes).digest('hex');
function identity() {
  const files = [];
  const walk = directory => { for (const entry of fs.readdirSync(directory, { withFileTypes: true })) { const file = path.join(directory, entry.name); if (entry.isDirectory()) walk(file); else if (/\.(mjs|json|py)$/.test(file)) files.push(file); } };
  walk(path.join(repository, 'src')); walk(path.join(repository, '.generated/document-core')); walk(path.join(repository, '.generated/native')); files.push(import.meta.filename);
  const members = Object.fromEntries(files.sort().map(file => [path.relative(repository, file), hash(fs.readFileSync(file))]));
  return { sha256: hash(JSON.stringify(members)), files: members };
}
const corpus=path.resolve(process.env.RPF_XLSX_CORPUS??path.join(repository,'artifacts/bench/corpus'));
if(!fs.existsSync(path.join(corpus,'scale-100000.xlsx'))){const built=spawnSync('python3',[fixturePath('build-corpus.py'),'--output',corpus],{encoding:'utf8'});if(built.status!==0)throw Error(built.stderr||'Fixture generation failed')}
const source = identity(), file = path.join(corpus, 'scale-100000.xlsx'), original = fs.readFileSync(file);
const s = createServiceFixture({ prefix: 'xlsx-worker-scale-', limits: { workerTimeoutMs: 60000 } });
const report = { schema: 'rpf.xlsx.production-worker-probe.v1', measuredAt: new Date().toISOString(), source,
  fixtureId: 'scale-100000', fixtureSha256: hash(original), fixtureBytes: original.length,
  build: { kind: 'source', bundleSha256: null }, machine: { node: process.version, platform: os.platform(), release: os.release(), cpu: os.cpus()[0]?.model },
  procedure: 'office_open then xlsx_query A1:J10, using the production worker runner and its normal XLSX Node heap policy',
  timingQualification: 'Diagnostic single run; concurrent source tests may run. No latency-budget acceptance claim.',
  workerPeakRssBytes: null, workerMemoryStatus: 'Peak worker RSS is not exposed by the production runner; functional completion is recorded separately from memory observations.', calls: [], passed: false };
async function call(name, args) {
  const started = performance.now(), result = await s.service.call(name, args);
  report.calls.push({ name, elapsedMs: performance.now() - started, returnedJsonBytes: Buffer.byteLength(JSON.stringify(result)), isError: Boolean(result.isError) });
  if (result.isError) throw Error(JSON.stringify(result.payload.error));
  return result.payload;
}
try {
  const copy = path.join(s.docs, 'scale.xlsx'); fs.writeFileSync(copy, original);
  const opened = await call('office_open', { path: copy }), view = { documentId: opened.document.id, revision: opened.document.sourceRevision };
  const read = await call('xlsx_query', { view, sheet: 'Scale', range: 'A1:J10', limit: 100 });
  report.returnedCells = read.data.items.length; report.calculation = read.data.calculation;
  report.mismatches = read.data.items.filter(item => {
    const row = Number(item.cell.slice(1)), column = item.cell.charCodeAt(0) - 65;
    return column === 9 ? item.calculation?.status !== 'calculated' || item.calculation.value !== 2 * row + 1 : item.value !== row + column;
  }).map(item => item.cell);
  report.passed = report.returnedCells === 100 && report.mismatches.length === 0 && report.calls.at(-1).returnedJsonBytes <= 262144;
} catch (error) { report.error = String(error); }
finally { s.close(); }
report.originalUnchanged = hash(fs.readFileSync(file)) === hash(original);
report.sourceStableDuringRun = source.sha256 === identity().sha256;
report.serviceParentPeakRssBytes = process.resourceUsage().maxRSS * 1024;
const output = path.resolve(process.env.RPF_XLSX_MCP_OUTPUT??path.join(repository, 'artifacts/bench/xlsx-worker-probe.json')); fs.mkdirSync(path.dirname(output),{recursive:true}); fs.writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
process.stdout.write(JSON.stringify({ output, passed: report.passed, sourceStableDuringRun: report.sourceStableDuringRun, calls: report.calls, error: report.error }) + '\n');
process.exitCode = report.passed && report.originalUnchanged ? 0 : 1;
