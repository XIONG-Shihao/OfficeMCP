/** Reusable bounded publication checker; every file check remains inside its job. */
import fs from 'node:fs';
import path from 'node:path';
import {parentPort} from 'node:worker_threads';
import {execFileSync} from 'node:child_process';
import {syncDirectory} from './files.mjs';
import {freshRevisionProof} from './streaming-file-revision.mjs';
parentPort.on('message',data=>{
 const {id,file,tmp,expectedRevision,maxBytes,stageDirectory}=data;
 try{
  if(data.operation==='source'){
   const proof=freshRevisionProof(file,maxBytes);
   if(proof.revision!==expectedRevision)throw Error('Source changed since open.');
   parentPort.postMessage({id,result:{file,identity:proof.identity}});
  }else{
   let cloned=false;
   if(process.platform==='darwin'){
    try{execFileSync('/bin/cp',['-c',file,tmp],{stdio:'ignore',timeout:10000});cloned=true;}
    catch{fs.rmSync(tmp,{force:true});}
   }
   if(!cloned)fs.copyFileSync(file,tmp,fs.constants.COPYFILE_EXCL|fs.constants.COPYFILE_FICLONE);
   fs.chmodSync(tmp,0o600);
   const proof=freshRevisionProof(tmp,maxBytes);
   if(proof.revision!==expectedRevision)throw Error('Publication artifact changed before staging.');
   const fd=fs.openSync(tmp,'r');try{fs.fsyncSync(fd);}finally{fs.closeSync(fd);}syncDirectory(path.dirname(tmp));
   parentPort.postMessage({id,result:{tmp,identity:proof.identity,stageDirectory}});
  }
 }catch(error){if(tmp)try{fs.unlinkSync(tmp);}catch{}if(stageDirectory)try{fs.rmdirSync(stageDirectory);}catch{}parentPort.postMessage({id,error:String(error.message)});}
});
