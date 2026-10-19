import test from './support/test.mjs';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {zipSync,strToU8} from 'fflate';
import {Document,Packer,Paragraph} from 'docx';
import {createAgentToolkit} from '@superdoc/sdk';
import {preflightSerializedBaseline,serializedHeadingEvidence,releaseSerializedBaseline} from '../src/formats/docx/docx-serialized-heading-proof.mjs';
import {freshRevisionProof} from '../src/infrastructure/streaming-file-revision.mjs';
import {executeDocxMutationPlan} from '../src/formats/docx/docx-mutation-plan.mjs';
import {clientFor} from '../src/formats/docx/docx-sdk.mjs';

const step={action:'insert_heading',arguments:{text:'Preflight heading',level:2,placement:{at:'document_end'},changeMode:'direct'}};
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const old='<w:p><w:r><w:t>Original</w:t></w:r></w:p>';
const heading='<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Preflight heading</w:t></w:r></w:p>';
function bytes(body){return zipSync({'[Content_Types].xml':strToU8('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/octet-stream"/></Types>'),
 'word/document.xml':strToU8(`<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`),
 'word/styles.xml':strToU8(`<w:styles xmlns:w="${W}"><w:style w:styleId="Heading2"><w:name w:val="heading 2"/></w:style></w:styles>`),
 'word/media/test.bin':strToU8('Preserved binary')});}
function fixture(){const root=fs.mkdtempSync(path.join(os.tmpdir(),'docx-preflight-'));const task={input:path.join(root,'source.docx'),output:path.join(root,'candidate.docx'),scratch:root,maxFileBytes:1024*1024,maxExpandedBytes:1024*1024,args:{steps:[step]}};
 fs.writeFileSync(task.input,bytes(old));fs.writeFileSync(task.output,bytes(old+heading));const baseline={file:path.join(root,'baseline.docx')};fs.writeFileSync(baseline.file,bytes(old));return{root,task,baseline,revision:freshRevisionProof(task.input,task.maxFileBytes).revision,close(){releaseSerializedBaseline(baseline);fs.rmSync(root,{recursive:true,force:true})}};}

test('preflight eligibility preserves the exact candidate proof and is bound to inputs and limits',()=>{
 const f=fixture();try{assert(preflightSerializedBaseline(f.task,f.baseline,f.revision));const {proof,independent}=serializedHeadingEvidence(f.task,f.baseline,step);assert(proof.checks.includes('source-document-domains-preserved-through-baseline-serialization'));assert.equal(independent.inputRevision,f.revision);
  assert(preflightSerializedBaseline(f.task,f.baseline,f.revision));assert.throws(()=>serializedHeadingEvidence({...f.task,maxFileBytes:f.task.maxFileBytes-1},f.baseline,step),{code:'INTEGRITY_FAILED'});
 }finally{f.close();}
});
test('normalization incompatibility selects full verification before an edit; integrity and archive failures remain errors',()=>{
 const f=fixture();try{fs.writeFileSync(f.baseline.file,bytes(old.replace('Original','Normalized')));assert.equal(preflightSerializedBaseline(f.task,f.baseline,f.revision),false);
  assert.throws(()=>preflightSerializedBaseline(f.task,f.baseline,'sha256:'+'0'.repeat(64)),{code:'INTEGRITY_FAILED'});
  fs.writeFileSync(f.baseline.file,'not a package');assert.throws(()=>preflightSerializedBaseline(f.task,f.baseline,f.revision));
 }finally{f.close();}
});
test('changed source/baseline after preflight and candidate-only changes still fail the proof',()=>{
 for(const kind of ['source','baseline','candidate']){const f=fixture();try{assert(preflightSerializedBaseline(f.task,f.baseline,f.revision));
  if(kind==='source')fs.writeFileSync(f.task.input,bytes(old.replace('Original','Modified')));
  if(kind==='baseline')fs.writeFileSync(f.baseline.file,bytes(old.replace('Original','Modified')));
  if(kind==='candidate')fs.writeFileSync(f.task.output,bytes(old.replace('Original','Modified')+heading));
  assert.throws(()=>serializedHeadingEvidence(f.task,f.baseline,step),kind==='candidate'?{code:'PRESERVATION_FAILED'}:{code:'INTEGRITY_FAILED'});
 }finally{f.close();}}
});
test('incompatible real SDK baseline edits once, exports once after mutation, and never restarts',async()=>{
 const root=fs.mkdtempSync(path.join(os.tmpdir(),'docx-preflight-sdk-')),input=path.join(root,'source.docx');
 const source=await Packer.toBuffer(new Document({sections:[{children:[new Paragraph('Original paragraph'),new Paragraph('Untouched paragraph')]}]}));fs.writeFileSync(input,source);
 const task={input,output:path.join(root,'candidate.docx'),scratch:root,timeoutMs:30000,maxFileBytes:8*1024*1024,maxExpandedBytes:16*1024*1024,args:{steps:[step]}};
 const client=clientFor(task);let doc,actions=0,exports=0,reopens=0,restarts=0;
 try{doc=await client.open({doc:input,runtime:'v2',userName:'OfficeMCP'});const toolkit=await createAgentToolkit({provider:'generic',preset:'core'});
  const observed=new Proxy(doc,{get(target,key){const value=Reflect.get(target,key);if(key==='save')return async args=>{exports++;return value.call(target,args)};return typeof value==='function'?value.bind(target):value;}});
  const result=await executeDocxMutationPlan({client:{open:async args=>{reopens++;return client.open(args)}},doc:observed,task,userName:'OfficeMCP',
   executeStep:async s=>{actions++;const receipt=await toolkit.dispatch(doc,'superdoc_perform_action',{action:s.action,...s.arguments});assert.equal(receipt.status,'ok');return receipt;},
   discard:()=>doc.close({discard:true}),restart:async()=>{restarts++;assert.fail('Normalization eligibility must not restart an already edited SDK session')},setDocument:next=>{doc=next}});
  assert.equal(result.saved.engineEvidence.verification,'superdoc-sdk-readback-v1');assert.equal(actions,1);assert.equal(exports,2,'one baseline and one accepted candidate export');assert.equal(reopens,1,'only required semantic verification reopen');assert.equal(restarts,0);
  assert.deepEqual(fs.readFileSync(input),source);assert.match(await doc.getText({}),/Preflight heading/);assert(!fs.readdirSync(root).some(name=>name.startsWith('heading-baseline-')));
 }finally{if(doc)await doc.close({discard:true}).catch(()=>{});await client.dispose();fs.rmSync(root,{recursive:true,force:true});}
});
test('prepared resources are discarded when a request postcondition fails before mutation',async()=>{
 const f=fixture();let discarded=0,actions=0;
 try{const doc={save:async({out})=>fs.copyFileSync(f.task.input,out),getText:async()=> 'Original',close:async()=>{discarded++}};
  await assert.rejects(executeDocxMutationPlan({client:{},doc,task:{...f.task,args:{steps:[step],expect:{beforeOccurrences:{Missing:1}}}},
   executeStep:async()=>{actions++},discard:()=>doc.close()}),{code:'TARGET_CONFLICT'});
  assert.equal(actions,0);assert.equal(discarded,1);assert(!fs.readdirSync(f.root).some(name=>name.startsWith('heading-baseline-')));
 }finally{f.close();}
});
test('candidate-only incompatibility retains full semantic restart instead of accepting a weaker proof',async()=>{
 const f=fixture();let actions=0,restarts=0,current;
 const rich=bytes(old+heading.replace('<w:r>','<w:r><w:rPr><w:b/></w:rPr>'));
 const make=mutated=>{const d={mutated,save:async({out})=>fs.writeFileSync(out,d.mutated?rich:bytes(old)),close:async()=>{},
  get:async()=>({body:[{kind:'paragraph',paragraph:{inlines:[{kind:'run',run:{text:d.mutated?'Original Preflight heading':'Original'}}]}}]}),
  headerFooters:{parts:{list:async()=>({items:[],total:0})}}};
  for(const domain of ['images','sections','footnotes','fields','hyperlinks','contentControls','trackChanges','comments','captions','crossRefs'])d[domain]={list:async()=>({items:[],total:0})};return d;};
 const client={open:async()=>make(true)};current=make(false);
 try{const result=await executeDocxMutationPlan({client,doc:current,task:f.task,executeStep:async()=>{actions++;current.mutated=true;return{status:'ok'}},
  discard:async()=>{},restart:async()=>{restarts++;current=make(false);return{client,doc:current}}});
  assert.equal(restarts,1);assert.equal(actions,2);assert.equal(result.saved.engineEvidence.verification,'superdoc-sdk-readback-v1');assert(!fs.readdirSync(f.root).some(name=>name.startsWith('heading-baseline-')));
 }finally{f.close();}
});
