/** Bounded hashes for private DOCX snapshots; live user sources still hash fresh. */
import fs from 'node:fs'
import {createHash} from 'node:crypto'
import {fileRevision} from './files.mjs'
import {requireThat} from './errors.mjs'
export class DocxSnapshotRevisions {
  constructor(maximumEntries=64){this.maximumEntries=maximumEntries;this.entries=new Map();this.hits=0;this.misses=0}
  revision(file,maxBytes) {
    // Unix ctime detects same-size edits even when mtime is restored. Keep the
    // conservative existing path on platforms without that qualified contract.
    if(process.platform==='win32')return fileRevision(file,maxBytes)
    const fd=fs.openSync(file,fs.constants.O_RDONLY|(fs.constants.O_NOFOLLOW??0))
    try {
      const before=fs.fstatSync(fd,{bigint:true})
      requireThat(before.isFile()&&before.size<=BigInt(maxBytes),'LIMIT_EXCEEDED','Input must be a regular file within the configured size limit.')
      const identity=m=>[m.dev,m.ino,m.size,m.mtimeNs,m.ctimeNs].join(':')
      const key=identity(before),cached=this.entries.get(key)
      if(cached!==undefined){this.hits++;return cached}
      this.misses++
      const hash=createHash('sha256'),buffer=Buffer.allocUnsafe(1024*1024);let total=0,count
      while((count=fs.readSync(fd,buffer,0,buffer.length,null))>0){total+=count;requireThat(total<=maxBytes,'LIMIT_EXCEEDED','Input grew beyond the limit.');hash.update(buffer.subarray(0,count))}
      requireThat(BigInt(total)===before.size&&identity(fs.fstatSync(fd,{bigint:true}))===key,'INTEGRITY_FAILED','Private DOCX snapshot changed while hashing.')
      const revision='sha256:'+hash.digest('hex')
      if(this.entries.size>=this.maximumEntries)this.entries.delete(this.entries.keys().next().value)
      this.entries.set(key,revision);return revision
    } finally {fs.closeSync(fd)}
  }
}
