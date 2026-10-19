/** Previous-turn native worker versus fresh and reused immutable native indexes. */
import fs from 'node:fs/promises';
import path from 'node:path';
import assert from 'node:assert/strict';
import {spawn} from 'node:child_process';
import {pathToFileURL,fileURLToPath} from 'node:url';
import {createHash} from 'node:crypto';
import {sampleSpreadsheetMemory} from '../engine-source.mjs';
if(!process.env.RPF_XLSX_SOURCE_XLSX)throw Error('Set RPF_XLSX_SOURCE_XLSX to the supplied xlsx_sample fixture path');
const source=path.resolve(process.env.RPF_XLSX_SOURCE_XLSX);
if(!process.env.RPF_XLSX_MCP_BASELINE_WORKER)throw Error('Set RPF_XLSX_MCP_BASELINE_WORKER to the preserved worker module');
if(process.argv[2]==='--child'){
 const mode=process.argv[3],module=mode==='previous'?process.env.RPF_XLSX_MCP_BASELINE_WORKER:fileURLToPath(new URL('../../src/formats/xlsx/xlsx-worker.mjs',import.meta.url));
 const {runXlsxWorkbook}=await import(pathToFileURL(path.resolve(module))),bytes=await fs.readFile(source),begin=performance.now();
 const result=await runXlsxWorkbook({kind:'xlsx_oss_read',input:source,maxExpandedBytes:1024**3,args:{sheet:'xlsx_sample',range:'F129970:F129972',offset:0,limit:3}},bytes);
 assert.equal(result.items.at(-1).value,21);
 console.log(JSON.stringify({mode,elapsedMs:performance.now()-begin,cacheHit:result.indexCacheHit??false,itemsSha256:createHash('sha256').update(JSON.stringify(result.items)).digest('hex')}));
}else{
 const runs=[];for(const mode of ['previous','current-first','current-reused']){
  const child=spawn(process.execPath,[import.meta.filename,'--child',mode],{stdio:['ignore','pipe','pipe']});let output='',errors='';child.stdout.on('data',b=>output+=b);child.stderr.on('data',b=>errors+=b);
  const done=new Promise((resolve,reject)=>{child.on('error',reject);child.on('exit',code=>code===0?resolve():reject(Error(errors||`exit ${code}`)))});let sampler;
  try{sampler=await sampleSpreadsheetMemory(child.pid);await done;runs.push({...JSON.parse(output),memory:await sampler.stop()});sampler=null;}finally{await sampler?.stop();if(child.exitCode===null)child.kill();}
 }
 for(const run of runs)run.memoryQualification=run.elapsedMs<500?'Too short for the 500ms sampler; do not interpret its peak as working memory':'Sampled observation only; not a kernel peak';
 assert(runs.every(r=>r.itemsSha256===runs[0].itemsSha256));assert.equal(runs[2].cacheHit,true);
 const report={passed:true,source,runs,limits:'Same supplied 1.74M-cell workbook and three-cell read. Previous-turn sealed MCP/native binary versus current release helper, separate fresh workers; excludes Electron, RSS sampled every 500ms, single observations, warm OS cache, no forced GC. Cache hit recorded rather than assumed.'};
 const output=path.resolve(process.env.RPF_XLSX_MCP_OUTPUT??fileURLToPath(new URL('../../artifacts/bench/mcp-cache.json',import.meta.url)));await fs.mkdir(path.dirname(output),{recursive:true});await fs.writeFile(output,JSON.stringify(report,null,2)+'\n');console.log(JSON.stringify({runs:runs.map(({memory,...r})=>({...r,peakRssBytes:memory.sampledPeakRssBytes}))},null,2));
}
