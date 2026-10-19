/** DOCX validation fast path: bounded native inflate and views of stored members. */
import {inflateRawSync} from 'node:zlib'
import {checkArchiveLimits} from './archive-limits.mjs'
import {requireThat,OfficeError} from '../../infrastructure/errors.mjs'
export function nativeUnpack(bytes,maxExpandedBytes) {
 const {entries}=checkArchiveLimits(bytes,maxExpandedBytes)
 let end=bytes.length-22
 while(end>=Math.max(0,bytes.length-65557)&&bytes.readUInt32LE(end)!==0x06054b50)end--
 requireThat(end>=0,'INVALID_DOCUMENT','Invalid ZIP directory.')
 const start=bytes.readUInt32LE(end+16),parts=Object.create(null);let at=start
 for(let i=0;i<entries;i++) {
  const method=bytes.readUInt16LE(at+10),compressed=bytes.readUInt32LE(at+20),expanded=bytes.readUInt32LE(at+24),n=bytes.readUInt16LE(at+28),extra=bytes.readUInt16LE(at+30),comment=bytes.readUInt16LE(at+32),local=bytes.readUInt32LE(at+42)
  requireThat(at+46+n+extra+comment<=end&&local+30<=start&&bytes.readUInt32LE(local)===0x04034b50,'INVALID_DOCUMENT','Invalid ZIP member header.')
  const name=bytes.subarray(at+46,at+46+n).toString('utf8'),localName=bytes.readUInt16LE(local+26),localExtra=bytes.readUInt16LE(local+28),data=local+30+localName+localExtra
  requireThat(bytes.readUInt16LE(local+8)===method&&!(bytes.readUInt16LE(local+6)&1)&&data<=start&&compressed<=start-data,'INVALID_DOCUMENT','Invalid ZIP member payload.')
  requireThat(bytes.subarray(local+30,local+30+localName).toString('utf8')===name,'INVALID_DOCUMENT','ZIP local and directory member names disagree.')
  const payload=bytes.subarray(data,data+compressed);let decoded
  if(method===0)decoded=payload
  else if(method===8){try{decoded=inflateRawSync(payload,{maxOutputLength:Math.max(1,expanded)})}catch{throw new OfficeError('INVALID_DOCUMENT','Could not decode ZIP members.')}}
  else throw new OfficeError('UNSUPPORTED_FEATURE','Unsupported ZIP compression method.')
  requireThat(decoded.byteLength===expanded,'INVALID_DOCUMENT','ZIP member size mismatch.')
  parts[name]=decoded;at+=46+n+extra+comment
 }
 requireThat(Object.keys(parts).length===entries,'INVALID_DOCUMENT','ZIP inventory mismatch.')
 requireThat(parts['[Content_Types].xml'],'UNSUPPORTED_FORMAT','OOXML content types are missing.')
 return parts
}
