import test from './support/test.mjs'
import assert from 'node:assert/strict'
import {zipSync,strToU8,unzipSync} from 'fflate'
import {nativeUnpack} from '../src/formats/opc/native-unpack.mjs'
const fixture=level=>Buffer.from(zipSync({'[Content_Types].xml':strToU8('<Types/>'),'word/document.xml':strToU8('<document>body</document>'),'word/media/a.bin':new Uint8Array(10000).fill(42)},{level}))
test('native DOCX decoding preserves exact stored/deflated bytes and stored members share immutable input',()=>{
 for(const level of [0,6]){const bytes=fixture(level),expected=unzipSync(bytes),actual=nativeUnpack(bytes,20000);assert.deepEqual(Object.keys(actual),Object.keys(expected));for(const name of Object.keys(actual))assert.deepEqual(Buffer.from(actual[name]),Buffer.from(expected[name]));if(level===0)assert.equal(actual['word/media/a.bin'].buffer,bytes.buffer)}
})
test('native DOCX decoding rejects expansion overflow, malformed names/offsets/payload sizes and unsafe members',()=>{
 const bytes=fixture(0);assert.throws(()=>nativeUnpack(bytes,100),{code:'LIMIT_EXCEEDED'})
 const end=bytes.length-22,central=bytes.readUInt32LE(end+16)
 for(const mutate of [b=>b.writeUInt32LE(b.length,central+42),b=>b.writeUInt32LE(1,central+24),b=>b[30]^=1,b=>b.writeUInt16LE(1,central+8)]){const bad=Buffer.from(bytes);mutate(bad);assert.throws(()=>nativeUnpack(bad,20000))}
 const unsafe=Buffer.from(zipSync({'[Content_Types].xml':strToU8('<Types/>'),'../escape':strToU8('bad')}));assert.throws(()=>nativeUnpack(unsafe,20000),{code:'INVALID_DOCUMENT'})
})
