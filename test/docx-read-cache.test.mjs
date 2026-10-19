import test from './support/test.mjs';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import {zipSync,unzipSync,strFromU8,strToU8} from 'fflate';
import {DOMParser} from '@xmldom/xmldom';
import { createServiceFixture } from './service-fixture.mjs';
import {readProcessTree} from '../src/infrastructure/docx-process-budget.mjs';
const pause=ms=>new Promise(r=>setTimeout(r,ms));
async function close(s){s.service.close();const deadline=Date.now()+10000;while(s.service.docxPool.workers.size){assert(Date.now()<deadline);await pause(20);}fs.rmSync(s.root,{recursive:true,force:true});}

test('real SDK queries cache once per view and mutation invalidates reads/snapshots and stale targets',async()=>{
 const s=createServiceFixture({prefix:'docx-read-cache-',toolset:'docx-agent'}),raw=[];
 const run=s.service.docxPool.run.bind(s.service.docxPool);s.service.docxPool.run=async(...args)=>{const result=await run(...args);raw.push(result.engineSession);return result;};
 try {
  const file=path.join(s.docs,'source.docx');fs.copyFileSync(new URL('./fixtures/docx/images.docx',import.meta.url),file);const original=fs.readFileSync(file);
  const xml=new DOMParser().parseFromString(strFromU8(unzipSync(original)['word/document.xml']),'application/xml'),ns='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const originalText=Array.from(xml.getElementsByTagNameNS(ns,'p')[0].getElementsByTagNameNS(ns,'t')).map(node=>node.textContent).join('');assert(originalText);
  const view=(await s.call('docx_open',{path:file})).view,args={view,select:{type:'node',nodeType:'paragraph'},require:'all',offset:0,limit:1};
  const first=await s.call('docx_query',args),second=await s.call('docx_query',args);assert.deepEqual(first.data.items,second.data.items);assert.equal(raw[1].cache.reads.hits,1);
  const complete=await s.call('docx_read',{view,representation:'complete',offset:0,limit:1});const snapshot=complete.data.readSnapshotId;assert(snapshot);assert(raw.at(-1).cache.snapshots.entries===1);
  const edited=await s.call('docx_apply',{view,operationId:'cache-edit',action:'rewrite_block',selectorTarget:first.data.items[0].target,arguments:{text:originalText+' CACHE_MUTATION_',changeMode:'direct'}});
  assert.equal(raw.at(-1).cache.reads.entries,0);assert.equal(raw.at(-1).cache.snapshots.entries,0);
  const changed=await s.call('docx_query',{...args,view:edited.view});assert.notEqual(changed.data.items[0].target.id,first.data.items[0].target.id);assert(raw.at(-1).cache.generation>raw[1].cache.generation);
  const currentText=await s.call('docx_query',{view:edited.view,select:{type:'text',pattern:'CACHE_MUTATION_'},require:'exactlyOne',offset:0,limit:1});assert.equal(currentText.data.matchedTotal,1);
  const stale=await s.service.call('docx_apply',{view:edited.view,operationId:'cache-stale-target',action:'rewrite_block',selectorTarget:first.data.items[0].target,arguments:{text:'Must not use a stale target'}});assert(stale.isError);assert.equal(stale.payload.error.code,'TARGET_CONFLICT');
  const staleSnapshot=await s.service.call('docx_read',{view:edited.view,representation:'complete',readSnapshotId:snapshot,offset:0,limit:1});assert(staleSnapshot.isError);assert.equal(staleSnapshot.payload.error.code,'TARGET_CONFLICT');
  const externalParts=unzipSync(original);externalParts['word/document.xml']=strToU8(strFromU8(externalParts['word/document.xml']).replace(originalText,'EXTERNAL_CACHE_REVISION'));
  const external=zipSync(externalParts);fs.writeFileSync(file,external);
  const externalView=(await s.call('docx_open',{path:file})).view;assert.notEqual(externalView.revision,view.revision);
  const externalRead=await s.call('docx_query',{view:externalView,select:{type:'text',pattern:'EXTERNAL_CACHE_REVISION'},require:'exactlyOne',offset:0,limit:1});assert.equal(externalRead.data.matchedTotal,1);
  // Existing views are immutable snapshots. Opening a changed source is a
  // separate identity; source publication still checks its captured revision.
  const retained=await s.call('docx_query',{view:edited.view,select:{type:'text',pattern:'CACHE_MUTATION_'},require:'exactlyOne',offset:0,limit:1});assert.equal(retained.data.matchedTotal,1);
  const saved=await s.service.call('docx_save',{view:edited.view,operationId:'cache-external-save',destination:{rootId:'output',relativePath:'saved.docx'}});
  // save-as may legitimately publish to a new destination while leaving the
  // changed external source alone; it must never write its old bytes over it.
  assert(saved.isError||saved.payload.data.published);assert.deepEqual(fs.readFileSync(file),Buffer.from(external));
  const document=s.service.get('document',externalView.documentId);fs.writeFileSync(document.snapshot,original);
  const tampered=await s.service.call('docx_query',{view:externalView,select:{type:'text',pattern:'EXTERNAL_CACHE_REVISION'},require:'exactlyOne',offset:0,limit:1});assert(tampered.isError);assert.equal(tampered.payload.error.code,'INTEGRITY_FAILED');
  assert(raw.every(row=>row.cache.reads.bytes<=1024*1024&&row.cache.snapshots.bytes<=1024*1024));assert(original.length>0);
 } finally {await close(s);}
});

test('complete-read caches are bounded across distinct snapshot handles',async()=>{
 const s=createServiceFixture({prefix:'docx-snapshot-cache-',toolset:'docx-agent'}),raw=[];
 const run=s.service.docxPool.run.bind(s.service.docxPool);s.service.docxPool.run=async(...args)=>{const result=await run(...args);raw.push(result.engineSession.cache);return result;};
 try {
  const file=path.join(s.docs,'source.docx');fs.copyFileSync(new URL('./fixtures/docx/text.docx',import.meta.url),file);const view=(await s.call('docx_open',{path:file})).view;
  for(let i=0;i<35;i++)await s.call('docx_read',{view,representation:'complete',offset:0,limit:1});
  assert(raw.every(row=>row.reads.entries<=32&&row.snapshots.entries<=32&&row.reads.bytes<=1024*1024&&row.snapshots.bytes<=1024*1024));
 } finally {await close(s);}
});

test('SDK child crash cannot make a warm worker serve an unavailable binding',async()=>{
 const s=createServiceFixture({prefix:'docx-sdk-cache-crash-',toolset:'docx-agent'});
 try{
  const file=path.join(s.docs,'source.docx');fs.copyFileSync(new URL('./fixtures/docx/text.docx',import.meta.url),file);const original=fs.readFileSync(file),view=(await s.call('docx_open',{path:file})).view;
  const args={view,select:{type:'node',nodeType:'paragraph'},require:'all',offset:0,limit:1};const first=await s.call('docx_query',args);await s.call('docx_query',args);
  const pid=s.service.docxPool.slot.child.pid,tree=await readProcessTree([pid]);const sdk=tree.processes.filter(row=>row.pid!==pid).sort((a,b)=>b.rssBytes-a.rssBytes)[0];assert(sdk.rssBytes>150*1024*1024);process.kill(sdk.pid,'SIGKILL');await pause(50);
  const recovered=await s.call('docx_query',args);assert.equal(recovered.data.items.length,first.data.items.length);assert.equal(s.service.docxPool.slot.child.pid,pid);assert.deepEqual(fs.readFileSync(file),original);
  const fresh=await s.call('docx_query',{...args,limit:2});assert(fresh.data.items.length>=1);
 }finally{await close(s);}
});
