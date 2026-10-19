import test from './support/test.mjs';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import { createServiceFixture } from './service-fixture.mjs';
import { readProcessTree } from '../src/infrastructure/docx-process-budget.mjs';

const fixture=fileURLToPath(new URL('./fixtures/docx/images.docx',import.meta.url));
const pause=ms=>new Promise(r=>setTimeout(r,ms));
async function until(fn){const deadline=Date.now()+15000;while(!await fn()){assert(Date.now()<deadline,'Resource cleanup deadline');await pause(20);}}
async function read(s,view){return s.call('docx_query',{view,select:{type:'node',nodeType:'paragraph'},require:'all',offset:0,limit:1});}

test('resident DOCX workers are bounded across documents, reuse valid models and fully reap SDK groups',async()=>{
  const s=createServiceFixture({prefix:'docx-resident-',toolset:'docx-agent',limits:{maxDocxActors:2,maxDocxResidentActors:2,docxIdleMs:60000}}),views=[],observations=[];
  let timer;
  try {
    for(let i=0;i<5;i++){const file=path.join(s.docs,`${i}.docx`);fs.copyFileSync(fixture,file);views.push((await s.call('docx_open',{path:file})).view);}
    timer=setInterval(()=>observations.push(s.service.docxPool.stats()),5);
    const raw=[],run=s.service.docxPool.run.bind(s.service.docxPool);s.service.docxPool.run=async(...args)=>{const result=await run(...args);raw.push(result.engineSession);return result;};
    const concurrent=await Promise.all(Array.from({length:6},()=>read(s,views[0]))),first=concurrent[0];
    assert(raw.every(session=>session.pid===raw[0].pid&&session.loads===1),'Concurrent same-view actions share one actual SDK model');
    const worker=s.service.docxPool.slot.child.pid;const second=await read(s,views[0]);assert.equal(s.service.docxPool.slot.child.pid,worker);
    assert.equal(first.data.items[0].itemId,second.data.items[0].itemId);
    for(let i=1;i<5;i++)await read(s,views[i]);
    assert(s.service.docxPool.workers.size<=2);assert(observations.every(x=>x.resident<=2));
    const last=await read(s,views[0]);assert.equal(first.data.items[0].text,last.data.items[0].text);
    const roots=[...s.service.docxPool.workers].map(slot=>slot.child.pid);
    const tree=await readProcessTree(roots,{groups:process.platform==='win32'?[]:roots});assert(tree.processes.length>roots.length,'The budget includes actual SDK descendants');
    s.service.docxPool.close();await until(()=>s.service.docxPool.workers.size===0);
    assert.equal((await readProcessTree(roots,{groups:process.platform==='win32'?[]:roots})).processes.length,0);
    assert.equal(s.service.docxPool.pendingRequests,0);
  } finally {clearInterval(timer);s.service.close();await until(()=>s.service.docxPool.workers.size===0);fs.rmSync(s.root,{recursive:true,force:true});}
});

test('two active documents can queue safely behind a single resident SDK worker',async()=>{
  const s=createServiceFixture({prefix:'docx-one-resident-',toolset:'docx-agent',limits:{maxDocxActors:2,maxDocxResidentActors:1}}),views=[];
  try {
    for(let i=0;i<3;i++){const file=path.join(s.docs,`${i}.docx`);fs.copyFileSync(fixture,file);views.push((await s.call('docx_open',{path:file})).view);}
    const results=await Promise.all(views.map(view=>read(s,view)));assert(results.every(r=>r.data.items.length));assert.equal(s.service.docxPool.workers.size,1);assert.equal(s.service.docxPool.pendingRequests,0);
  } finally {s.service.close();await until(()=>s.service.docxPool.workers.size===0);fs.rmSync(s.root,{recursive:true,force:true});}
});

test('idle timeout closes worker and SDK; a saved view reopens without stale handles',async()=>{
  const s=createServiceFixture({prefix:'docx-idle-budget-',toolset:'docx-agent',limits:{docxIdleMs:150}});
  try {
    const file=path.join(s.docs,'source.docx');fs.copyFileSync(fixture,file);const view=(await s.call('docx_open',{path:file})).view;
    const first=await read(s,view),roots=[...s.service.docxPool.workers].map(slot=>slot.child.pid);
    await until(()=>s.service.docxPool.workers.size===0);
    assert.equal((await readProcessTree(roots,{groups:process.platform==='win32'?[]:roots})).processes.length,0);
    const second=await read(s,view);assert.equal(first.data.items[0].text,second.data.items[0].text);assert(!roots.includes(s.service.docxPool.slot.child.pid));
  } finally {s.service.close();await until(()=>s.service.docxPool.workers.size===0);fs.rmSync(s.root,{recursive:true,force:true});}
});

test('real SDK process memory above an explicit low budget refuses a result and reaps all children',async()=>{
  const s=createServiceFixture({prefix:'docx-memory-budget-',toolset:'docx-agent',limits:{maxDocxProcessBytes:96*1024*1024}});
  try {
    const file=path.join(s.docs,'source.docx');fs.copyFileSync(fixture,file);const before=fs.readFileSync(file),view=(await s.call('docx_open',{path:file})).view;
    const result=await s.service.call('docx_query',{view,select:{type:'node',nodeType:'paragraph'},require:'all',offset:0,limit:1});
    assert(result.isError);assert.equal(result.payload.error.code,'LIMIT_EXCEEDED');
    await until(()=>s.service.docxPool.workers.size===0);assert.deepEqual(fs.readFileSync(file),before);assert.equal(s.service.docxPool.activeRuns,0);
  } finally {s.service.close();await until(()=>s.service.docxPool.workers.size===0);fs.rmSync(s.root,{recursive:true,force:true});}
});

test('queued cancellation and global request admission hold before a real actor can run',async()=>{
  const s=createServiceFixture({prefix:'docx-queue-budget-',limits:{maxDocxQueuedRequests:2,maxDocxActors:1}}),pool=s.service.docxPool;
  const hold={};hold.promise=new Promise(r=>{hold.resolve=r;});
  // Holding the existing global semaphore leaves every SDK body unadmitted.
  await pool.acquire();
  try {
    const abort=new AbortController();
    const a=pool.run('docx_native_read','unused-a',{documentId:'a'},{},abort.signal);a.catch(()=>{});
    const b=pool.run('docx_native_read','unused-b',{documentId:'b'},{},abort.signal);b.catch(()=>{});
    await until(()=>pool.waiters.length===2);
    await assert.rejects(pool.run('docx_native_read','unused-c',{documentId:'c'},{}),{code:'BUSY'});
    abort.abort();await Promise.all([assert.rejects(a,{code:'CANCELLED'}),assert.rejects(b,{code:'CANCELLED'})]);assert.equal(pool.pendingRequests,0);assert.equal(pool.workers.size,0);
  } finally {pool.release();s.close();}
});
