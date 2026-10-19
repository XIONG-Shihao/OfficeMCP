import test from './support/test.mjs';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import {createServiceFixture} from './service-fixture.mjs';
import {validateDocxCandidate} from '../src/formats/docx/docx-validation.mjs';
import {zipSync,unzipSync} from 'fflate';
import {digest} from '../src/infrastructure/files.mjs';
const pause=ms=>new Promise(r=>setTimeout(r,ms));
test('actual SDK readback proof is reused only for the exact bound candidate and independent media checks still run',async()=>{
 const s=createServiceFixture({prefix:'docx-bound-proof-',toolset:'docx-agent'});
 try{
  const file=path.join(s.docs,'source.docx');fs.copyFileSync(new URL('./fixtures/docx/images.docx',import.meta.url),file);
  const view=(await s.call('docx_open',{path:file})).view;
  // Omit explicit placement to exercise the general SDK roundtrip path. The
  // separately tested serialized proof requires an explicit body-end intent.
  const edited=await s.call('docx_apply',{view,operationId:'proof-edit',action:'insert_heading',arguments:{text:'PROOF_HEADING',level:2,changeMode:'direct'}});
  const c=s.service.get('candidate',edited.view.candidateId),d=s.service.doc(view.documentId);
  const task={input:c.path,original:d.snapshot,operations:c.operations,maxExpandedBytes:s.service.config.maxExpandedBytes};
  const noSdk=new Proxy({}, {get:()=>()=>{assert.fail('Already verified immutable bytes must not repeat SDK semantic extraction');}});
  const valid=await validateDocxCandidate(noSdk,task);assert(valid.passed);assert(valid.checks.includes('SDK-readback-proof-bound-to-exact-candidate-hash'));
  const missing=structuredClone(task);delete missing.operations.at(-1).engineEvidence.sdkSnapshotHash;await assert.rejects(validateDocxCandidate(noSdk,missing),{code:'VALIDATION_FAILED'});
  const mismatch=structuredClone(task);mismatch.operations.at(-1).engineEvidence.outputRevision='sha256:'+'0'.repeat(64);await assert.rejects(validateDocxCandidate(noSdk,mismatch),{code:'VALIDATION_FAILED'});
  const original=fs.readFileSync(c.path),parts=unzipSync(original),image=Object.keys(parts).find(name=>name.startsWith('word/media/')&&!name.endsWith('/'));assert(image);delete parts[image];const damaged=zipSync(parts);fs.writeFileSync(c.path,damaged);
  const forged=structuredClone(task);forged.operations.at(-1).engineEvidence.outputRevision=digest(damaged);
  await assert.rejects(validateDocxCandidate(noSdk,forged));fs.writeFileSync(c.path,original);
 }finally{s.service.close();const deadline=Date.now()+10000;while(s.service.docxPool.workers.size){assert(Date.now()<deadline);await pause(20);}fs.rmSync(s.root,{recursive:true,force:true});}
});
