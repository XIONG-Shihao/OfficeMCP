import test from './support/test.mjs'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import {DocxSnapshotRevisions} from '../src/infrastructure/docx-snapshot-revisions.mjs'
import {fileRevision} from '../src/infrastructure/files.mjs'
test('private DOCX hash reuse rejects same-size edits with restored mtime, replacement and unsafe/oversized files',()=>{
 const directory=fs.mkdtempSync(path.join(os.tmpdir(),'docx-snapshot-revisions-')),file=path.join(directory,'snapshot.docx'),cache=new DocxSnapshotRevisions(2)
 try {
  fs.writeFileSync(file,'original');const stamp=fs.statSync(file),first=cache.revision(file,100)
  assert.equal(first,fileRevision(file,100));assert.equal(cache.revision(file,100),first)
  if(process.platform!=='win32')assert.equal(cache.hits,1)
  fs.writeFileSync(file,'mutated!');fs.utimesSync(file,stamp.atime,stamp.mtime)
  assert.notEqual(cache.revision(file,100),first);assert.equal(cache.revision(file,100),fileRevision(file,100))
  const replacement=path.join(directory,'new');fs.writeFileSync(replacement,'replaced');fs.renameSync(replacement,file)
  assert.equal(cache.revision(file,100),fileRevision(file,100));assert.throws(()=>cache.revision(file,2),{code:'LIMIT_EXCEEDED'})
  if(process.platform!=='win32'){const link=path.join(directory,'link');fs.symlinkSync(file,link);assert.throws(()=>cache.revision(link,100))}
  for(let i=0;i<5;i++){const other=path.join(directory,String(i));fs.writeFileSync(other,String(i));cache.revision(other,100)}
  assert(cache.entries.size<=2)
 } finally {fs.rmSync(directory,{recursive:true,force:true})}
})
