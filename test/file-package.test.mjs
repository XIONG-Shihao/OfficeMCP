import test from './support/test.mjs';import assert from 'node:assert/strict';
import fs from 'node:fs';import os from 'node:os';import path from 'node:path';
import {zipSync,strToU8,unzipSync} from 'fflate';
import {decodeFilePackage,decodePackage} from '../src/formats/opc/ooxml.mjs';
import {packageEvidence} from '../src/formats/opc/engine-evidence.mjs';
const limit=20*1024*1024;
const fixture=level=>Buffer.from(zipSync({'[Content_Types].xml':strToU8('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/octet-stream"/></Types>'),'word/document.xml':strToU8('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>original</w:t></w:r></w:p></w:body></w:document>'),'word/media/a.bin':new Uint8Array(1024*1024).fill(37)}, {level}));
test('file-backed compressed/stored members preserve exact package evidence and close resources',()=>{
 const root=fs.mkdtempSync(path.join(os.tmpdir(),'file-package-'));
 try {for(const level of [0,6]){
  const bytes=fixture(level),file=path.join(root,level+'.docx');fs.writeFileSync(file,bytes);
  const before=decodeFilePackage(file,limit,limit),after=decodeFilePackage(file,limit,limit),expected=unzipSync(bytes);
  try {assert.deepEqual(Object.keys(before.parts),Object.keys(expected));for(const name of Object.keys(expected))assert.deepEqual(before.parts[name],Buffer.from(expected[name]));
   assert.notEqual(before.parts['word/media/a.bin'],before.parts['word/media/a.bin'],'Binary members are not retained in the package context');
   assert.deepEqual(packageEvidence(before,after,limit),packageEvidence(bytes,bytes,limit));
   assert(!after.xml.has('word/document.xml'));
   assert.throws(()=>decodePackage(before,1),{code:'LIMIT_EXCEEDED'});
  } finally {before.close();after.close();}
 }} finally {fs.rmSync(root,{recursive:true,force:true});}
});
test('file-backed proof rejects member tampering, malformed ZIP headers and source replacement',()=>{
 const root=fs.mkdtempSync(path.join(os.tmpdir(),'file-package-guard-')),file=path.join(root,'input.docx');
 try {
  const bytes=fixture(0);fs.writeFileSync(file,bytes);const context=decodeFilePackage(file,limit,limit);
  try {fs.writeFileSync(file,Buffer.from(bytes).fill(0,40,41));assert.throws(()=>context.filePackage.assertUnchanged(),{code:'INTEGRITY_FAILED'});}finally{context.close();}
  fs.writeFileSync(file,bytes);const replaced=decodeFilePackage(file,limit,limit);
  try {fs.renameSync(file,file+'.old');fs.writeFileSync(file,bytes);assert.throws(()=>replaced.filePackage.assertUnchanged(),{code:'INTEGRITY_FAILED'});}finally{replaced.close();}
  const bad=Buffer.from(bytes);bad[30]^=1;fs.writeFileSync(file,bad);assert.throws(()=>decodeFilePackage(file,limit,limit),{code:'INVALID_DOCUMENT'});
  fs.writeFileSync(file,bytes);assert.throws(()=>decodeFilePackage(file,10,limit),{code:'LIMIT_EXCEEDED'});
 }finally{fs.rmSync(root,{recursive:true,force:true});}
});
