import test from './support/test.mjs';
import assert from 'node:assert/strict';
import { readOwnedNativeTable, selectProcessTree, ProcessBudget } from '../src/infrastructure/docx-process-budget.mjs';
test('native counters retry a live worker missing from the first process snapshot and include its child RSS',async()=>{
  let reads=0,pauses=0;
  const rows=[{pid:10,parent:1,group:10,rssBytes:100},{pid:11,parent:10,group:10,rssBytes:200}];
  const observed=await readOwnedNativeTable([10],{readTable:async()=>++reads===1?[]:rows,probe:()=>{},pause:async()=>{pauses++}});
  assert.equal(reads,2);assert.equal(pauses,1);assert.equal(selectProcessTree(observed,[10],[10]).rssBytes,300);
});
test('an exited root still includes orphaned owned-group descendants',async()=>{
  const rows=[{pid:11,parent:1,group:10,rssBytes:200}];
  const observed=await readOwnedNativeTable([10],{readTable:async()=>rows,probe:()=>{throw Object.assign(Error('gone'),{code:'ESRCH'})},pause:async()=>assert.fail('No retry needed')});
  assert.equal(selectProcessTree(observed,[10],[10]).rssBytes,200);
});
test('persistently missing live RSS, probe permissions and reader errors remain fail-closed',async()=>{
  let reads=0;
  await assert.rejects(readOwnedNativeTable([10],{readTable:async()=>{reads++;return[]},probe:()=>{},pause:async()=>{}}),/unavailable/);
  assert.equal(reads,3);
  await assert.rejects(readOwnedNativeTable([10],{readTable:async()=>[],probe:()=>{throw Object.assign(Error('denied'),{code:'EPERM'})}}),{code:'EPERM'});
  await assert.rejects(readOwnedNativeTable([10],{readTable:async()=>{throw Error('counter failure')}}),/counter failure/);
});
test('a recovered observation still enforces the memory limit',async()=>{
  let reads=0;
  const budget=new ProcessBudget({roots:()=>[10],maximumBytes:100,sample:async()=>selectProcessTree(await readOwnedNativeTable([10],{
    readTable:async()=>++reads===1?[]:[{pid:10,parent:1,group:10,rssBytes:101}],probe:()=>{},pause:async()=>{},
  }),[10])});
  await assert.rejects(budget.check(),{code:'LIMIT_EXCEEDED'});assert.equal(budget.maximumObservedBytes,101);
});
