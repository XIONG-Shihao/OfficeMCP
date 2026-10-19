/** Bounded file-backed OPC members for read-only preservation proofs. */
import fs from 'node:fs';
import {inflateRawSync} from 'node:zlib';
import {checkArchiveLimits} from './archive-limits.mjs';
import {requireThat,OfficeError} from '../../infrastructure/errors.mjs';
import {freshRevisionProof,fileIdentity} from '../../infrastructure/streaming-file-revision.mjs';

export function openFilePackage(file,limit,maxBytes) {
  const proof=freshRevisionProof(file,maxBytes);
  const fd=fs.openSync(file,fs.constants.O_RDONLY|(fs.constants.O_NOFOLLOW??0));
  const read=(at,size)=>{
    requireThat(Number.isSafeInteger(at)&&Number.isSafeInteger(size)&&at>=0&&size>=0&&at+size<=Number(proof.identity.size),'INVALID_DOCUMENT','ZIP member is out of bounds.');
    const bytes=Buffer.allocUnsafe(size);let total=0;
    while(total<size){const count=fs.readSync(fd,bytes,total,size-total,at+total);requireThat(count>0,'INTEGRITY_FAILED','ZIP input changed during reading.');total+=count;}
    return bytes;
  };
  try {
    requireThat(JSON.stringify(fileIdentity(fs.fstatSync(fd,{bigint:true})))===JSON.stringify(proof.identity),'INTEGRITY_FAILED','Package identity changed.');
    const size=Number(proof.identity.size),tailStart=Math.max(0,size-65557),tail=read(tailStart,size-tailStart);
    let end=tail.length-22;
    while(end>=0&&tail.readUInt32LE(end)!==0x06054b50)end--;
    requireThat(end>=0,'INVALID_DOCUMENT','Invalid ZIP directory.');
    const start=tail.readUInt32LE(end+16),length=tail.readUInt32LE(end+12),endAt=tailStart+end;
    requireThat(length<=16*1024*1024&&start+length<=endAt,'LIMIT_EXCEEDED','ZIP directory exceeds its bounds.');
    const directory=read(start,length),synthetic=Buffer.concat([directory,tail.subarray(end,end+22)]);
    synthetic.writeUInt32LE(0,directory.length+16);
    const bounds=checkArchiveLimits(synthetic,limit),entries=new Map(),parts=Object.create(null),xmlMembers=new Map();
    let at=0;
    for(let i=0;i<bounds.entries;i++) {
      const method=directory.readUInt16LE(at+10),compressed=directory.readUInt32LE(at+20),expanded=directory.readUInt32LE(at+24),n=directory.readUInt16LE(at+28),extra=directory.readUInt16LE(at+30),comment=directory.readUInt16LE(at+32),local=directory.readUInt32LE(at+42);
      requireThat(at+46+n+extra+comment<=length&&local+30<=start,'INVALID_DOCUMENT','Invalid ZIP member header.');
      const name=directory.subarray(at+46,at+46+n).toString('utf8'),header=read(local,30),localName=header.readUInt16LE(26),localExtra=header.readUInt16LE(28),data=local+30+localName+localExtra;
      requireThat(header.readUInt32LE(0)===0x04034b50&&header.readUInt16LE(8)===method&&!(header.readUInt16LE(6)&1)&&data<=start&&compressed<=start-data,'INVALID_DOCUMENT','Invalid ZIP member payload.');
      requireThat(read(local+30,localName).toString('utf8')===name,'INVALID_DOCUMENT','ZIP member names disagree.');
      requireThat(method===0||method===8,'UNSUPPORTED_FEATURE','Unsupported ZIP compression method.');
      entries.set(name,{method,compressed,expanded,data,local,crc:directory.readUInt32LE(at+16),central:Buffer.from(directory.subarray(at,at+46+n+extra+comment))});at+=46+n+extra+comment;
      Object.defineProperty(parts,name,{enumerable:true,get(){
        if(xmlMembers.has(name))return xmlMembers.get(name);
        const entry=entries.get(name),payload=read(entry.data,entry.compressed);let bytes;
        try {bytes=entry.method===0?payload:inflateRawSync(payload,{maxOutputLength:Math.max(1,entry.expanded)});}
        catch {throw new OfficeError('INVALID_DOCUMENT','Could not decode ZIP members.');}
        requireThat(bytes.length===entry.expanded,'INVALID_DOCUMENT','ZIP member size mismatch.');
        if(/\.(xml|rels|vml)$/i.test(name))xmlMembers.set(name,bytes);
        return bytes;
      }});
    }
    requireThat(Object.hasOwn(parts,'[Content_Types].xml'),'UNSUPPORTED_FORMAT','OOXML content types are missing.');
    let closed=false;
    return {parts,expandedBytes:bounds.expandedBytes,revision:proof.revision,
      entries:()=>Array.from(entries,([name,entry])=>({name,...entry,central:Buffer.from(entry.central)})),
      assertUnchanged(){requireThat(!closed,'INTEGRITY_FAILED','Package is closed.');const current=freshRevisionProof(file,maxBytes);requireThat(current.revision===proof.revision&&JSON.stringify(current.identity)===JSON.stringify(proof.identity)&&JSON.stringify(fileIdentity(fs.fstatSync(fd,{bigint:true})))===JSON.stringify(proof.identity),'INTEGRITY_FAILED','Package changed during preservation checks.');},
      close(){if(!closed){closed=true;fs.closeSync(fd);xmlMembers.clear();}}};
  }catch(error){fs.closeSync(fd);throw error;}
}
