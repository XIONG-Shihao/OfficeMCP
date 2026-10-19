import test from './support/test.mjs';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import {tmpdir} from 'node:os';
import {publishArtifactManifest} from '../src/infrastructure/artifact-manifest.mjs';

test('portable artifact metadata replays exact bytes and refuses foreign descriptors and links',()=>{
 const root=fs.mkdtempSync(path.join(tmpdir(),'office-artifact-descriptor-'));
 try{
  const output=path.join(root,'saved.docx');fs.writeFileSync(output,'document');
  const metadata={schema:'mcp-artifact/1',sha256:'1'.repeat(64),inputs:[{uri:'file:///source.docx',sha256:'2'.repeat(64)}]};
  const descriptor=publishArtifactManifest(output,metadata);
  assert.deepEqual(JSON.parse(fs.readFileSync(descriptor,'utf8')),metadata);
  assert.equal(publishArtifactManifest(output,metadata),descriptor);
  assert.throws(()=>publishArtifactManifest(output,{...metadata,inputs:[]}),{code:'INTEGRITY_FAILED'});
  fs.unlinkSync(descriptor);fs.symlinkSync(output,descriptor);
  assert.throws(()=>publishArtifactManifest(output,metadata),{code:'INTEGRITY_FAILED'});
  assert.equal(fs.readFileSync(output,'utf8'),'document');
  fs.unlinkSync(descriptor);
  assert.throws(()=>publishArtifactManifest(output,{...metadata,inputs:[{uri:'x'.repeat(65536)}]}),{code:'LIMIT_EXCEEDED'});
  assert.deepEqual(fs.readdirSync(root),['saved.docx']);
 }finally{fs.rmSync(root,{recursive:true,force:true})}
});
