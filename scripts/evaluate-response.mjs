// Local realistic acceptance: no model/provider calls and no source modifications.
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { execFileSync } from 'node:child_process';
import { pathToFileURL } from 'node:url';
const runtimeSource = path.resolve(process.env.OFFICEMCP_RESPONSE_SOURCE ?? new URL('../', import.meta.url).pathname);
const { OfficeService } = await import(pathToFileURL(path.join(runtimeSource, 'src/application/service.mjs')));
const { configFrom, digest } = await import(pathToFileURL(path.join(runtimeSource, 'src/infrastructure/files.mjs')));
const label = process.env.OFFICEMCP_RESPONSE_LABEL;
if (label && !/^[a-zA-Z0-9_-]+$/.test(label)) throw new Error('Use a filename-safe response label.');
const source = path.resolve(process.argv[2] ?? new URL('../test_doc/response.docx', import.meta.url).pathname);
const output = path.join(path.dirname(source), 'officemcp-results');
fs.mkdirSync(output, { recursive: true });
const reportPath = path.join(output, `response-workflows${label ? '-' + label : ''}.json`);
const root = fs.mkdtempSync(path.join(os.tmpdir(), 'officemcp-response-')), stamp = Date.now();
const configPath = path.join(root, 'config.json');
fs.writeFileSync(configPath, JSON.stringify({ readRoots: [path.dirname(source)], writeRoots: { output }, stateDirectory: path.join(root, 'state'), toolset: 'full' }));
const config = configFrom(configPath), service = new OfficeService(config);
const report = { scope: 'MCP semantic/package acceptance; document rendering and page-placement checks are outside this runner', started: new Date().toISOString(), sourceBytes: fs.statSync(source).size, sourceHash: digest(fs.readFileSync(source)), limits: Object.fromEntries(['maxFileBytes','maxExpandedBytes','maxStateBytes','workerTimeoutMs'].map(k => [k,config[k]])), steps: [], workflows: {}, memoryMeasurement: 'Sum of ps RSS for the coordinator and descendants sampled approximately each second; shared pages can be counted more than once. Local acceptance measurements, not a controlled performance benchmark.', peakProcessTreeRssBytes: null };
report.runtimeSource = runtimeSource;
report.runtimeVersion = JSON.parse(fs.readFileSync(path.join(runtimeSource, 'package.json'), 'utf8')).version;
let activeStep = null;
function sampleMemory() {
  try {
    const entries = execFileSync('ps', ['-axo', 'pid=,ppid=,rss='], { encoding: 'utf8' }).trim().split('\n').map(s => s.trim().split(/\s+/).map(Number));
    const included = new Set([process.pid]);
    for (let oldSize = -1; oldSize !== included.size;) { oldSize = included.size; for (const [pid,parent] of entries) if (included.has(parent)) included.add(pid); }
    const total = entries.reduce((n,[pid,,rss]) => n + (included.has(pid) ? rss*1024 : 0),0);
    report.peakProcessTreeRssBytes = Math.max(report.peakProcessTreeRssBytes ?? 0, total);
    if (activeStep) activeStep.peakProcessTreeRssBytes = Math.max(activeStep.peakProcessTreeRssBytes ?? 0, total);
  } catch { report.memoryMeasurement = 'unavailable: process inspection denied'; }
}
const timer = setInterval(sampleMemory, 1000);
async function call(tool, args) {
  const started = performance.now();
  const record = activeStep = { tool, peakProcessTreeRssBytes: null };
  sampleMemory();
  const r = await service.call(tool,args);
  Object.assign(record, { elapsedMs: Math.round(performance.now()-started), jsonBytes: Buffer.byteLength(JSON.stringify(r.payload)), nativeImageBytes: (r.content ?? []).reduce((n,i) => n+Buffer.from(i.data,'base64').length,0), status:r.payload.status, ...(r.isError ? {error:r.payload.error} : {}) });
  report.steps.push(record); sampleMemory(); activeStep = null; console.log(JSON.stringify(record));
  if(r.isError) throw Object.assign(new Error(r.payload.error.message),{code:r.payload.error.code});
  return r.payload;
}
try {
  let view = (await call('docx_open',{path:source})).view;
  const inventory = await call('docx_inspect',{view});
  report.workflows.orientation = {passed:true, counts:inventory.data.counts};
  const read = await call('docx_read',{view,representation:'linear',limit:30});
  let p, q;
  for (const candidate of read.data.items.filter(i => i.text?.length>=12 && i.text.length<=1000 && i.target && ['paragraph','heading'].includes(i.kind)).sort((a,b)=>b.text.length-a.text.length)) {
    const found=await call('docx_query',{view,select:{type:'text',pattern:candidate.text},require:'all'});
    if(found.data.total===1){p=candidate;q=found;break;}
  }
  if(!p) throw new Error('No unique initial paragraph.');
  report.workflows.targetedRetrieval={passed:q.data.items.length===1};
  const tables=await call('docx_read',{view,representation:'domain',domain:'tables',limit:2});
  const footers=await call('docx_read',{view,representation:'domain',domain:'headers_footers',limit:10});
  const images=await call('docx_read',{view,representation:'domain',domain:'images',limit:2});
  report.workflows.crossDomainRead={passed:true,tables:tables.data.items.length,footers:footers.data.items.length,images:images.data.items.length};
  const marker=' [OfficeMCP verified edit]', operationId=`response-edit-${stamp}`;
  view=(await call('docx_apply',{view,operationId,action:'replace_text',selectorTarget:q.data.items[0].target,arguments:{edits:[{find:p.text,replace:p.text+marker}],caseSensitive:true}})).view;
  const recovered=await call('docx_operation_get',{operationId});
  report.workflows.receiptRecovery={passed:recovered.view.revision===view.revision,scope:'completed receipt recovery, not process-kill injection'};
  await call('docx_query',{view,select:{type:'text',pattern:marker},require:'exactlyOne'});
  report.workflows.guardedTextEdit={passed:true};
  const validation=await call('docx_validate',{view});
  report.workflows.preservation={passed:validation.data.passed,checks:validation.data.checks,packageSummary:validation.data.packageSummary,notVerified:validation.data.notVerified};
  const saved=await call('docx_save',{view,operationId:`response-save-${stamp}`,destination:{rootId:'output',relativePath:`response-verified-${stamp}.docx`}});
  report.output=saved.data.path;
  await call('docx_close',{view,disposition:'discard'});
  const reopened=await call('docx_open',{path:saved.data.path});
  await call('docx_query',{view:reopened.view,select:{type:'text',pattern:marker},require:'exactlyOne'});
  report.workflows.saveReopen={passed:true,outputBytes:fs.statSync(saved.data.path).size};
  report.passed=Object.values(report.workflows).every(w=>w.passed);
}catch(error){report.passed=false; report.error={code:error.code,message:error.message};}
finally{
  clearInterval(timer); sampleMemory(); service.close();
  report.originalUnchanged=digest(fs.readFileSync(source))===report.sourceHash;
  report.finished=new Date().toISOString();
  fs.writeFileSync(reportPath,JSON.stringify(report,null,2)+'\n');
  if(report.passed) fs.rmSync(root,{recursive:true,force:true}); else report.diagnosticDirectory=root;
  fs.writeFileSync(reportPath,JSON.stringify(report,null,2)+'\n');
}
console.log(JSON.stringify({passed:report.passed,report:reportPath,output:report.output,error:report.error}));
process.exitCode=report.passed?0:1;
