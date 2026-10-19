/** Always-fresh bounded SHA; no cache, and detect mutation during the read. */
import fs from 'node:fs';import {createHash} from 'node:crypto';import {requireThat} from './errors.mjs';
export const fileIdentity=m=>Object.fromEntries(['dev','ino','size','mtimeNs','ctimeNs'].map(k=>[k,String(m[k])]));
export function freshRevisionProof(file,maxBytes) {
 const fd=fs.openSync(file,fs.constants.O_RDONLY|(fs.constants.O_NOFOLLOW??0));
 try{
  const before=fs.fstatSync(fd,{bigint:true}),identity=fileIdentity(before);
  requireThat(before.isFile()&&before.size<=BigInt(maxBytes),'LIMIT_EXCEEDED','Input must be a regular file within the configured size limit.');
  const hash=createHash('sha256'),buffer=Buffer.allocUnsafe(1024*1024);let total=0,count;
  while((count=fs.readSync(fd,buffer,0,buffer.length,null))>0){total+=count;requireThat(total<=maxBytes,'LIMIT_EXCEEDED','Input grew beyond the limit.');hash.update(buffer.subarray(0,count));}
  requireThat(BigInt(total)===before.size&&JSON.stringify(fileIdentity(fs.fstatSync(fd,{bigint:true})))===JSON.stringify(identity),
   'INTEGRITY_FAILED','File changed while hashing.');
  const current=fs.lstatSync(file,{bigint:true});requireThat(current.isFile()&&JSON.stringify(fileIdentity(current))===JSON.stringify(identity),
   'INTEGRITY_FAILED','File identity changed while hashing.');
  return{revision:'sha256:'+hash.digest('hex'),identity};
 }finally{fs.closeSync(fd);}
}
