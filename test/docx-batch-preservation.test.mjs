import test from './support/test.mjs';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import {createServiceFixture} from './service-fixture.mjs';
import {docxBytes} from './fixtures.mjs';
import {unzipSync,strFromU8} from 'fflate';
import {DOMParser} from '@xmldom/xmldom';
const pause=ms=>new Promise(r=>setTimeout(r,ms));
async function cleanup(s){s.service.close();const until=Date.now()+10000;while(s.service.docxPool.workers.size){assert(Date.now()<until);await pause(20);}fs.rmSync(s.root,{recursive:true,force:true});}
async function fixture(){const s=createServiceFixture({prefix:'docx-batch-preservation-',toolset:'docx-agent'}),file=path.join(s.docs,'source.docx'),original=docxBytes(['Alpha clause']);fs.writeFileSync(file,original);return{...s,file,original,view:(await s.call('docx_open',{path:file})).view};}

test('a second failed batch step discards earlier SDK mutation and accepts no partial candidate',async()=>{
 const s=await fixture();
 try{
  const result=await s.service.call('docx_apply',{view:s.view,operationId:'fail-second',steps:[
   {action:'insert_heading',arguments:{text:'BATCH_PARTIAL',level:2,placement:{at:'document_end'},changeMode:'direct'}},
   {action:'replace_text',arguments:{edits:[{find:'Missing target',replace:'Must not publish'}],changeMode:'direct'}}
  ]});
  assert(result.isError);assert.deepEqual(fs.readFileSync(s.file),Buffer.from(s.original));assert.equal(s.service.rows('candidate').length,0);
  assert.match(result.payload.error.message,/step 2.*after 1 completed/);
  const read=await s.call('docx_read',{view:s.view,representation:'text',offset:0,limit:10});assert(read.data.items.some(item=>item.text.includes('Alpha clause')));assert(!read.data.items.some(item=>item.text.includes('BATCH_PARTIAL')));
  const status=await s.service.call('docx_operation_get',{operationId:'fail-second'});assert(status.isError&&status.payload.status==='failed');assert.match(status.payload.error.message,/step 2.*after 1 completed/);
 }finally{await cleanup(s);}
});

test('an SDK replacement receipt that skipped an edit is refused despite top-level success',async()=>{
 const s=await fixture();
 try{
  const result=await s.service.call('docx_apply',{view:s.view,operationId:'skip-edit',action:'replace_text',arguments:{edits:[{find:'Alpha clause',replace:'Partial clause'},{find:'Missing target',replace:'Must not publish'}],changeMode:'direct'}});
  assert(result.isError,'Incomplete SDK replacements must not become a candidate');assert.deepEqual(fs.readFileSync(s.file),Buffer.from(s.original));assert.equal(s.service.rows('candidate').length,0);
  assert.equal(result.payload.error.code,'PARTIAL_MUTATION');
 }finally{await cleanup(s);}
});

test('one explicit batch retains both subreceipts and saves/reopens only its complete verified state',async()=>{
 const s=await fixture();
 try{
  const edited=await s.call('docx_apply',{view:s.view,operationId:'complete-batch',steps:[
   {action:'replace_text',arguments:{edits:[{find:'Alpha clause',replace:'Updated clause'}],changeMode:'direct'}},
   {action:'insert_heading',arguments:{text:'BATCH_COMPLETE',level:2,placement:{at:'document_end'},changeMode:'direct'}}
  ]});assert.equal(edited.data.steps.length,2);assert(edited.data.steps.every(step=>step.verificationPassed));
  const saved=await s.call('docx_save',{view:edited.view,operationId:'batch-save',destination:{rootId:'output',relativePath:'saved.docx'}});
  const xml=strFromU8(unzipSync(fs.readFileSync(saved.data.path))['word/document.xml']);
  const document=new DOMParser().parseFromString(xml,'application/xml');
  const text=Array.from(document.getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main','t')).map(n=>n.textContent).join('');
  assert(text.includes('Updated clause')&&text.includes('BATCH_COMPLETE')&&!text.includes('Alpha clause'),'Word text can span several preserved runs');
  const reopened=(await s.call('docx_open',{path:saved.data.path})).view;const read=await s.call('docx_query',{view:reopened,select:{type:'text',pattern:'BATCH_COMPLETE'},require:'exactlyOne',offset:0,limit:1});assert.equal(read.data.matchedTotal,1);assert.deepEqual(fs.readFileSync(s.file),Buffer.from(s.original));
 }finally{await cleanup(s);}
});
