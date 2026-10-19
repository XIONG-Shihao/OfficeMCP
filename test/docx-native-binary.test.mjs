import test from './support/test.mjs';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {fileURLToPath,pathToFileURL} from 'node:url';
import {installedSdkBinary} from '../src/formats/docx/docx-native-binary.mjs';
import {clientFor} from '../src/formats/docx/docx-sdk.mjs';

test('installed SDK binary uses only supported native packages and checks execution permission',()=>{
 const resolved=[],checked=[];
 const options={platform:'darwin',arch:'arm64',resolve:name=>{resolved.push(name);return '/installed/superdoc'},stat:()=>({isFile:()=>true}),access:(file,mode)=>checked.push({file,mode})};
 assert.equal(installedSdkBinary(options),'/installed/superdoc');
 assert.deepEqual(resolved,['@superdoc/sdk-darwin-arm64/bin/superdoc']);assert.deepEqual(checked,[{file:'/installed/superdoc',mode:fs.constants.X_OK}]);
 assert.equal(installedSdkBinary({...options,platform:'win32',arch:'x64'}),'/installed/superdoc');assert.equal(checked.at(-1).mode,fs.constants.F_OK);
});
test('missing, non-file, inaccessible and unsupported executables retain SDK fallback',()=>{
 const base={platform:'linux',arch:'x64',resolve:()=>'/installed/superdoc',stat:()=>({isFile:()=>true}),access:()=>{}};
 assert.equal(installedSdkBinary({...base,resolve:()=>{throw Error('missing')}}),undefined);
 assert.equal(installedSdkBinary({...base,stat:()=>({isFile:()=>false})}),undefined);
 assert.equal(installedSdkBinary({...base,access:()=>{throw Error('EACCES')}}),undefined);
 assert.equal(installedSdkBinary({...base,platform:'unknown',resolve:()=>{throw Error('must not resolve')}}),undefined);
});
test('actual SDK client construction leaves executable metadata and bytes untouched',async()=>{
 const file=installedSdkBinary();assert(file,'test runtime requires the installed SDK binary');
 const before=fs.statSync(file,{bigint:true});
 const client=clientFor({scratch:'/unused-sdk-state',timeoutMs:120000});
 try{const after=fs.statSync(file,{bigint:true});assert.equal(after.ctimeNs,before.ctimeNs);assert.equal(after.mode,before.mode);assert.equal(after.size,before.size);}
 finally{await client.dispose();}
});
test('an unrelated host SDK platform package cannot override the Office SDK dependency',async()=>{
 const root=fs.realpathSync(fs.mkdtempSync(path.join(os.tmpdir(),'office-sdk-dependency-')));
 try{
  const office=path.join(root,'node_modules/officemcp'),sdk=path.join(office,'node_modules/@superdoc/sdk'),platform=`@superdoc/sdk-${process.platform==='win32'?'windows':process.platform}-${process.arch}`;
  const makeBinary=(directory,label)=>{const file=path.join(directory,platform,'bin',process.platform==='win32'?'superdoc.exe':'superdoc');fs.mkdirSync(path.dirname(file),{recursive:true});fs.writeFileSync(file,label,{mode:0o755});return file;};
  fs.mkdirSync(path.join(sdk,'dist'),{recursive:true});fs.writeFileSync(path.join(sdk,'package.json'),JSON.stringify({name:'@superdoc/sdk',version:'2.2.1',main:'dist/index.js'}));fs.writeFileSync(path.join(sdk,'dist/index.js'),'');
  const expected=makeBinary(path.join(sdk,'node_modules'),'SDK owned'),host=makeBinary(path.join(root,'node_modules'),'host owned');
  const helper=path.join(office,'src/formats/docx/docx-native-binary.mjs');fs.mkdirSync(path.dirname(helper),{recursive:true});fs.copyFileSync(fileURLToPath(new URL('../src/formats/docx/docx-native-binary.mjs',import.meta.url)),helper);
  const isolated=await import(pathToFileURL(helper).href);assert.equal(isolated.installedSdkBinary(),expected);assert.notEqual(expected,host);
 }finally{fs.rmSync(root,{recursive:true,force:true});}
});
