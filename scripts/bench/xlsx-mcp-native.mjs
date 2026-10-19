import {fileURLToPath} from 'node:url';
/** Isolated MCP read-worker comparison. A separate parent samples RSS so a
 * synchronous JavaScript parse cannot starve the sampling timer. */
import fs from 'node:fs/promises';
import path from 'node:path';
import assert from 'node:assert/strict';
import {spawn} from 'node:child_process';
import {createHash} from 'node:crypto';
import {sampleSpreadsheetMemory} from '../engine-source.mjs';
if(!process.env.RPF_XLSX_SOURCE_XLSX)throw Error('Set RPF_XLSX_SOURCE_XLSX to the supplied xlsx_sample fixture path');
const source=path.resolve(process.env.RPF_XLSX_SOURCE_XLSX);
if(process.argv[2]==='--child'){
 const mode=process.argv[3];const {runXlsxWorkbook}=await import('../../src/formats/xlsx/xlsx-worker.mjs');
 const bytes=await fs.readFile(source),begin=performance.now();
 const result=await runXlsxWorkbook({kind:'xlsx_oss_read',...(mode==='current'?{input:source}:{}),maxExpandedBytes:1024**3,args:{sheet:'xlsx_sample',range:'F129970:F129972',offset:0,limit:3}},bytes);
 assert.equal(result.items.at(-1).value,21);assert.equal(result.items.length,3);
 console.log(JSON.stringify({mode,elapsedMs:performance.now()-begin,storage:result.storage??'legacy-full-view',itemsSha256:createHash('sha256').update(JSON.stringify(result.items)).digest('hex')}));
}else{
 const runs=[];
 for(const mode of ['previous','current']){
  const child=spawn(process.execPath,[import.meta.filename,'--child',mode],{stdio:['ignore','pipe','pipe']});let output='',errors='';child.stdout.on('data',b=>output+=b);child.stderr.on('data',b=>errors+=b);
  const done=new Promise((resolve,reject)=>{child.on('error',reject);child.on('exit',code=>code===0?resolve():reject(Error(errors||`worker exit ${code}`)))});
  let sampler;try{sampler=await sampleSpreadsheetMemory(child.pid);await done;runs.push({...JSON.parse(output),memory:await sampler.stop()});sampler=undefined;}finally{await sampler?.stop();if(child.exitCode===null)child.kill();}
 }
 assert.equal(runs[0].itemsSha256,runs[1].itemsSha256);
 const report={passed:true,source,workload:'MCP calculated read of three distant cells in supplied 1,740,219-cell workbook',runs,limits:'Previous full-JavaScript branch versus native branch, same worker function and source; current native helper is optimized. Separate fresh Node workers, sampled every 500 ms; no Electron, no forced GC, warm OS cache. One run per branch, not a release memory cap.'};
 const output=path.resolve(process.env.RPF_XLSX_MCP_OUTPUT??fileURLToPath(new URL('../../artifacts/bench/mcp-read.json',import.meta.url)));await fs.mkdir(path.dirname(output),{recursive:true});await fs.writeFile(output,JSON.stringify(report,null,2)+'\n');
 console.log(JSON.stringify({passed:true,runs:runs.map(({memory,...r})=>({...r,sampledPeakRssBytes:memory.sampledPeakRssBytes})),report:output},null,2));
}
