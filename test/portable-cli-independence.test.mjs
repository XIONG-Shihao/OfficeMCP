import test from './support/test.mjs';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {spawn} from 'node:child_process';
import {createInterface} from 'node:readline';
import {fileURLToPath} from 'node:url';

test('ordinary CLI ignores obsolete host environment and never imports an app module',async()=>{
 const root=fs.mkdtempSync(path.join(os.tmpdir(),'office-portable-cli-'));
 const docs=path.join(root,'docs');fs.mkdirSync(docs);
 const config=path.join(root,'config.json'),adapter=path.join(root,'forbidden-adapter.mjs');
 fs.writeFileSync(config,JSON.stringify({toolset:'docx-agent',readRoots:[docs],writeRoots:{output:docs},stateDirectory:path.join(root,'state')}));
 fs.writeFileSync(adapter,'throw Error("An app adapter must never execute");');
 const child=spawn(process.execPath,[fileURLToPath(new URL('../src/cli.mjs',import.meta.url)),'serve','--config',config],{env:{...process.env,RPF_OFFICEMCP_HOST_REQUIRED:'1',OFFICEMCP_PUBLICATION_ADAPTER:adapter},stdio:['pipe','pipe','pipe']});
 const exited=new Promise(resolve=>child.once('exit',resolve));
 let stderr='';child.stderr.on('data',chunk=>{stderr+=String(chunk)});
 const lines=createInterface({input:child.stdout});
 try{
  const response=new Promise((resolve,reject)=>{const timer=setTimeout(()=>reject(Error('CLI discovery timed out')),10000);lines.once('line',line=>{clearTimeout(timer);resolve(JSON.parse(line))})});
  child.stdin.write(JSON.stringify({jsonrpc:'2.0',id:1,method:'tools/list',params:{_meta:{'io.modelcontextprotocol/protocolVersion':'2026-07-28','io.modelcontextprotocol/clientInfo':{name:'independent',version:'1'},'io.modelcontextprotocol/clientCapabilities':{}}}})+'\n');
  const result=await response;assert(result.result.tools.length>0);assert(!result.error);assert(!stderr.includes('must never execute'));
  const helpResponse=new Promise((resolve,reject)=>{const timer=setTimeout(()=>reject(Error('CLI help timed out')),10000);lines.once('line',line=>{clearTimeout(timer);resolve(JSON.parse(line))})});
  child.stdin.write(JSON.stringify({jsonrpc:'2.0',id:2,method:'tools/call',params:{name:'docx_inspect',arguments:{},_meta:{'io.modelcontextprotocol/protocolVersion':'2026-07-28','io.modelcontextprotocol/clientInfo':{name:'independent',version:'1'},'io.modelcontextprotocol/clientCapabilities':{}}}})+'\n');
  const help=await helpResponse;assert(!help.error);assert(!help.result.isError);
  const guidance=help.result.structuredContent.data.usage.toolPrefix;
  assert.match(guidance,/advertised by your MCP connection/);
  assert.doesNotMatch(guidance,/rpf_|integrated app|codex|electron/i,'public guidance must not assume the consuming app or harness');
 }finally{child.stdin.end();await exited;lines.close();fs.rmSync(root,{recursive:true,force:true})}
});
