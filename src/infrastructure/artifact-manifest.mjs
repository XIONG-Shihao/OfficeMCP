/** Portable output provenance; contains no host, harness, model or session IDs. */
import fs from 'node:fs';
import path from 'node:path';
import {randomUUID} from 'node:crypto';
import {OfficeError,requireThat} from './errors.mjs';
import {syncDirectory} from './files.mjs';

export function publishArtifactManifest(destination,manifest){
 const file=destination+'.artifact.json',bytes=Buffer.from(JSON.stringify(manifest)+'\n');
 requireThat(bytes.length<=65536,'LIMIT_EXCEEDED','Artifact provenance exceeds its bounded descriptor size.');
 if(fs.existsSync(file)){
  const stat=fs.lstatSync(file);
  requireThat(stat.isFile()&&!stat.isSymbolicLink()&&stat.size<=65536,'INTEGRITY_FAILED','Artifact descriptor is not a bounded regular file.');
  requireThat(fs.readFileSync(file).equals(bytes),'INTEGRITY_FAILED','Artifact descriptor belongs to different output bytes or inputs.');
  return file;
 }
 const temporary=file+'.'+randomUUID()+'.tmp';
 try{
  const fd=fs.openSync(temporary,'wx',0o600);
  try{fs.writeSync(fd,bytes);fs.fsyncSync(fd)}finally{fs.closeSync(fd)}
  try{fs.linkSync(temporary,file)}catch(error){if(error.code!=='EEXIST')throw error;throw new OfficeError('INTEGRITY_FAILED','Artifact descriptor was created concurrently.');}
  syncDirectory(path.dirname(file));return file;
 }finally{try{fs.unlinkSync(temporary)}catch{}}
}
