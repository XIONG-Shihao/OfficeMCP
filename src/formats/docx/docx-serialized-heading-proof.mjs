/** Narrow serialized proof: one direct heading at body end, no other package change. */
import fs from 'node:fs';
import path from 'node:path';
import {randomUUID} from 'node:crypto';
import {decodePackage,decodeFilePackage,parseXml} from '../opc/ooxml.mjs';
import {packageEvidence} from '../opc/engine-evidence.mjs';
import {digest} from '../../infrastructure/files.mjs';
import {requireThat} from '../../infrastructure/errors.mjs';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
// Only host-owned objects can carry a prepared check. Retain bounded XML
// members/file handles until proof consumption; never retain all media buffers.
const preparedBaselines=new WeakMap();
const volatileWordAttributes=new Set(['w:rsid','w:rsidR','w:rsidRPr','w:rsidRDefault','w:rsidP','w:rsidDel','w:rsidTr','w:rsidSect']);
const equal=(a,b)=>a?.length===b?.length&&Buffer.from(a.buffer,a.byteOffset,a.byteLength).equals(Buffer.from(b.buffer,b.byteOffset,b.byteLength));
// Compare the XML infoset conservatively without building a second huge DOM.
// Only declaration spelling, attribute order/quote spacing, empty-tag spelling
// and Word's explicitly volatile revision/paragraph identities may normalize.
// Text, authored formatting, domain IDs and every other attribute stay exact.
export function normalizedSerializedXml(bytes) {
 const xml=Buffer.from(bytes.buffer,bytes.byteOffset,bytes.byteLength).toString('utf8');
 requireThat(equal(Buffer.from(xml,'utf8'),bytes),'UNSUPPORTED_FEATURE','Non-UTF8 XML needs full semantic verification.');
 requireThat(!/<!/.test(xml),'UNSUPPORTED_FEATURE','XML declarations/comments need the full verification path.');
 const out=[],scopes=[];let at=0;
 for(const token of xml.matchAll(/<(?:"[^"]*"|'[^']*'|[^'">])*>/g)){
  const text=xml.slice(at,token.index);if(text)out.push(JSON.stringify(['text',text]));const raw=token[0];at=token.index+raw.length;
  if(raw.startsWith('<?xml ')){
   requireThat(/version\s*=\s*["']1\.0["']/.test(raw)&&(!/encoding\s*=/.test(raw)||/encoding\s*=\s*["']utf-8["']/i.test(raw)),
    'UNSUPPORTED_FEATURE','XML encoding/version needs full semantic verification.');continue;
  }
  if(raw.startsWith('<?')){out.push(JSON.stringify(['pi',raw]));continue;}
  if(raw.startsWith('</')){scopes.pop();out.push(JSON.stringify(['close',raw.slice(2,-1).trim()]));continue;}
  const match=raw.match(/^<([^\s/>]+)/);requireThat(match,'INVALID_DOCUMENT','Malformed XML token.');
  const name=match[1],end=raw.endsWith('/>')?raw.length-2:raw.length-1,attrs=[];
  let cursor=match[0].length;
  const pattern=/\s+([^\s=]+)\s*=\s*("[^"]*"|'[^']*')/gy;
  while(cursor<end&&raw.slice(cursor,end).trim()){
   pattern.lastIndex=cursor;const a=pattern.exec(raw);requireThat(a,'INVALID_DOCUMENT','Malformed serialized attribute.');cursor=pattern.lastIndex;
   attrs.push([a[1],a[2].slice(1,-1)]);
  }
  let scope=scopes.at(-1)??Object.create(null);
  const declarations=attrs.filter(a=>a[0]==='xmlns'||a[0].startsWith('xmlns:'));
  if(declarations.length){scope=Object.create(scope);for(const [key,value]of declarations)scope[key==='xmlns'?'':key.slice(6)]=value;}
  const wordElement=name.startsWith('w:')&&scope.w===W;
  const kept=attrs.filter(([key])=>!(wordElement&&volatileWordAttributes.has(key))&&
    !(name==='w:p'&&scope.w===W&&scope.w14==='http://schemas.microsoft.com/office/word/2010/wordml'&&/^w14:(paraId|textId)$/.test(key)));
  kept.sort((a,b)=>a[0].localeCompare(b[0]));out.push(JSON.stringify(['open',name,kept]));
  if(raw.endsWith('/>'))out.push(JSON.stringify(['close',name]));else scopes.push(scope);
 }
 if(xml.slice(at))out.push(JSON.stringify(['text',xml.slice(at)]));return out.join('');
}
function verifyBaselineSerialization(original,baseline) {
 const names=Object.keys(original).filter(name=>!name.startsWith('docProps/')&&/\.(xml|rels|vml)$/i.test(name));
 const afterNames=Object.keys(baseline).filter(name=>!name.startsWith('docProps/')&&/\.(xml|rels|vml)$/i.test(name));
 requireThat(JSON.stringify(names.sort())===JSON.stringify(afterNames.sort()),'PRESERVATION_FAILED','Baseline serialization changed the document domain inventory.');
 for(const name of names)requireThat(normalizedSerializedXml(original[name])===normalizedSerializedXml(baseline[name]),
  'UNSUPPORTED_FEATURE',`SDK baseline normalization needs full semantic verification: ${name}.`);
 return['source-document-domains-preserved-through-baseline-serialization'];
}
export function serializedHeadingStep(task) {
 const steps=task.args.steps;
 if(steps?.length!==1||task.args.readback||task.args.semanticWorkflow)return null;
 const step=steps[0],a=step.arguments??{};
 return step.action==='insert_heading'&&a.changeMode==='direct'&&a.placement?.at==='document_end'&&
  typeof a.text==='string'&&a.text.length>0&&Number.isInteger(a.level)&&a.level>=1&&a.level<=6?step:null;
}
function bodyEnd(xml) {
 requireThat(!/<!/.test(xml),'UNSUPPORTED_FEATURE','Serialized proof requires ordinary SDK XML without declarations or comments.');
 const stack=[];let insertion=-1,bodyEnd=-1,rootStart='';
 const tokens=/<(?:"[^"]*"|'[^']*'|[^'">])*>/g;
 for(const token of xml.matchAll(tokens)){
  const raw=token[0];if(raw.startsWith('<?'))continue;
  const name=raw.match(/^<\/?([^\s/>]+)/)?.[1];requireThat(name,'INVALID_DOCUMENT','Malformed serialized XML token.');
  if(raw.startsWith('</')){
   requireThat(stack.at(-1)===name,'INVALID_DOCUMENT','Serialized XML nesting differs.');
   if(stack.length===2&&name==='w:body')bodyEnd=token.index;
   stack.pop();
  }else{
   if(!stack.length){requireThat(name==='w:document','UNSUPPORTED_FEATURE','Unsupported document prefix.');rootStart=raw;}
   if(stack.length===2&&stack[0]==='w:document'&&stack[1]==='w:body')insertion=name==='w:sectPr'?token.index:-1;
   if(!raw.endsWith('/>'))stack.push(name);
  }
 }
 requireThat(stack.length===0&&bodyEnd>=0,'INVALID_DOCUMENT','Incomplete serialized document.');
 return{offset:insertion<0?bodyEnd:insertion,rootStart};
}
function serializedHeadingContent(beforeBytes,afterBytes,step,limit) {
 const decode=bytes=>decodePackage(ArrayBuffer.isView(bytes)?Buffer.from(bytes.buffer,bytes.byteOffset,bytes.byteLength):bytes,limit,{nativeInflate:true});
 const before=decode(beforeBytes).parts,after=decode(afterBytes).parts;
 requireThat(JSON.stringify(Object.keys(before).sort())===JSON.stringify(Object.keys(after).sort()),'PRESERVATION_FAILED','Heading changed the package inventory.');
 for(const name of Object.keys(before))if(name!=='word/document.xml')requireThat(equal(before[name],after[name]),'PRESERVATION_FAILED',`Heading changed unrelated package part: ${name}.`);
 const a=Buffer.from(before['word/document.xml']),b=Buffer.from(after['word/document.xml']);
 const old=a.toString('utf8'),{offset,rootStart}=bodyEnd(old),byteOffset=Buffer.byteLength(old.slice(0,offset));
 const extra=b.length-a.length;
 requireThat(extra>0&&extra<=1024*1024&&equal(a.subarray(0,byteOffset),b.subarray(0,byteOffset))&&equal(a.subarray(byteOffset),b.subarray(byteOffset+extra)),
  'PRESERVATION_FAILED','Heading changed existing document XML or was inserted outside body end.');
 const fragment=b.subarray(byteOffset,byteOffset+extra).toString('utf8'),doc=parseXml(`${rootStart}<w:body>${fragment}</w:body></w:document>`);
 const body=doc.getElementsByTagNameNS(W,'body')[0];requireThat(body,'UNSUPPORTED_FEATURE','Missing WordprocessingML body.');
 const nodes=Array.from(body.childNodes).filter(n=>n.nodeType===1);
 requireThat(nodes.length===1&&nodes[0].namespaceURI===W&&nodes[0].localName==='p','VALIDATION_FAILED','Heading proof requires exactly one new paragraph.');
 const paragraph=nodes[0],text=Array.from(paragraph.getElementsByTagNameNS(W,'t')).map(n=>n.textContent).join('');
 requireThat(text===step.arguments.text,'VALIDATION_FAILED','Serialized heading text differs from the requested edit.');
 const childKinds={p:new Set(['pPr','r']),pPr:new Set(['pStyle','numPr']),numPr:new Set(['ilvl','numId']),r:new Set(['t']),pStyle:new Set(),ilvl:new Set(),numId:new Set(),t:new Set()};
 for(const n of [paragraph,...Array.from(paragraph.getElementsByTagName('*'))]){
  requireThat(n.namespaceURI===W&&childKinds[n.localName],'UNSUPPORTED_FEATURE','Rich heading requires full SDK verification.');
  for(const child of Array.from(n.childNodes).filter(c=>c.nodeType===1))requireThat(child.namespaceURI===W&&childKinds[n.localName].has(child.localName),
   'UNSUPPORTED_FEATURE','Heading nesting requires full SDK verification.');
  if(['ilvl','numId'].includes(n.localName))requireThat(n.getAttributeNS(W,'val')==='0','UNSUPPORTED_FEATURE','Numbered heading needs full verification.');
 }
 const styleId=paragraph.getElementsByTagNameNS(W,'pStyle')[0]?.getAttributeNS(W,'val');
 const styles=parseXml(Buffer.from(after['word/styles.xml']).toString('utf8'));
 const style=Array.from(styles.getElementsByTagNameNS(W,'style')).find(n=>n.getAttributeNS(W,'styleId')===styleId);
 const name=style?.getElementsByTagNameNS(W,'name')[0]?.getAttributeNS(W,'val');
 const outline=paragraph.getElementsByTagNameNS(W,'outlineLvl')[0]??style?.getElementsByTagNameNS(W,'outlineLvl')[0];
 requireThat(name?.replace(/\s/g,'').toLowerCase()===`heading${step.arguments.level}`||Number(outline?.getAttributeNS(W,'val'))===step.arguments.level-1,
  'VALIDATION_FAILED','Serialized heading level differs from the requested edit.');
 return{profile:'docx.serialized-heading.v1',text,level:step.arguments.level,
  checks:['exact-existing-document-XML-preserved','exact-unrelated-package-members-preserved','one-requested-heading-at-body-end']};
}
export function proveSerializedHeading(beforeBytes,afterBytes,step,limit){
 return{...serializedHeadingContent(beforeBytes,afterBytes,step,limit),baselineRevision:digest(beforeBytes),outputRevision:digest(afterBytes)};
}
export async function captureSerializedBaseline(doc,task) {
 const file=path.join(task.scratch,`heading-baseline-${randomUUID()}.docx`);
 try{await doc.save({out:file,force:false});return{file};}catch(error){fs.rmSync(file,{force:true});throw error;}
}
export function releaseSerializedBaseline(baseline) {
 const prepared=preparedBaselines.get(baseline);
 if(prepared){preparedBaselines.delete(baseline);prepared.original.close();prepared.baselinePackage.close();}
}
/** Decide normalization eligibility before mutation, without accepting a candidate. */
export function preflightSerializedBaseline(task,baseline,inputRevision) {
 releaseSerializedBaseline(baseline);
 const contexts=[];
 try {
  const open=file=>{const context=decodeFilePackage(file,task.maxExpandedBytes,task.maxFileBytes);contexts.push(context);return context;};
  const original=open(task.input),baselinePackage=open(baseline.file);
  requireThat(original.filePackage.revision===inputRevision,'INTEGRITY_FAILED','Source changed before serialized proof preparation.');
  const checks=verifyBaselineSerialization(original.parts,baselinePackage.parts);
  preparedBaselines.set(baseline,{input:task.input,file:baseline.file,maxFileBytes:task.maxFileBytes,
   maxExpandedBytes:task.maxExpandedBytes,original,baselinePackage,checks});
  return true;
 } catch(error) {
  try {for(const context of contexts)context.filePackage.assertUnchanged();}
  finally {for(const context of contexts)context.close();}
  if(['PRESERVATION_FAILED','UNSUPPORTED_FEATURE','VALIDATION_FAILED'].includes(error.code))return false;
  throw error;
 }
}
export function serializedHeadingEvidence(task,baseline,step) {
 const prepared=preparedBaselines.get(baseline);
 // Consume once. Every later file-stability check still runs before acceptance.
 preparedBaselines.delete(baseline);
 const contexts=prepared?[prepared.original,prepared.baselinePackage]:[];
 try {
  const open=file=>{const context=decodeFilePackage(file,task.maxExpandedBytes,task.maxFileBytes);contexts.push(context);return context;};
  if(prepared)requireThat(prepared.input===task.input&&prepared.file===baseline.file&&
   prepared.maxFileBytes===task.maxFileBytes&&prepared.maxExpandedBytes===task.maxExpandedBytes,
   'INTEGRITY_FAILED','Serialized proof preparation belongs to different inputs or limits.');
  const original=prepared?.original??open(task.input),baselinePackage=prepared?.baselinePackage??open(baseline.file),candidate=open(task.output);
  const initialChecks=prepared?.checks??verifyBaselineSerialization(original.parts,baselinePackage.parts);
  const proof=serializedHeadingContent(baselinePackage,candidate,step,task.maxExpandedBytes);
  proof.checks.push(...initialChecks);proof.baselineRevision=baselinePackage.filePackage.revision;
  const independent=packageEvidence(original,candidate,task.maxExpandedBytes,{nativeInflate:true});
  baselinePackage.filePackage.assertUnchanged();
  proof.outputRevision=independent.outputRevision;
  return{proof,independent};
 } catch(error) {
  // A candidate-only incompatibility may retry semantically. Changed prepared
  // inputs must never be reclassified as an ordinary proof incompatibility.
  if(prepared){prepared.original.filePackage.assertUnchanged();prepared.baselinePackage.filePackage.assertUnchanged();}
  throw error;
 } finally {for(const context of contexts)context.close();}
}
/** Host-owned ledger only; revisions must have been checked against actual files. */
export function boundSerializedValidation(operations,inputRevision,outputRevision) {
 if(!operations?.length||!operations.every(op=>op.engineEvidence?.verification==='superdoc-serialized-heading-v1'))return null;
 let revision=inputRevision;const checks=new Set();
 for(const op of operations){
  const e=op.engineEvidence,p=e.serializedProof,v=e.packageEvidence;
  requireThat(e.inputRevision===revision&&e.outputRevision===p?.outputRevision&&p?.profile==='docx.serialized-heading.v1'&&
   /^sha256:[a-f0-9]{64}$/.test(p.baselineRevision??'')&&v?.inputRevision===revision&&v.outputRevision===e.outputRevision&&
   Array.isArray(p.checks)&&p.checks.includes('source-document-domains-preserved-through-baseline-serialization')&&Array.isArray(v.checks),
   'VALIDATION_FAILED','Serialized proof does not bind this exact candidate chain.');
  for(const check of [...p.checks,...v.checks])checks.add(check);
  revision=e.outputRevision;
 }
 requireThat(revision===outputRevision,'VALIDATION_FAILED','Serialized candidate changed after verification.');
 return{passed:true,revision:outputRevision,profile:'docx.serialized-heading.v1',checks:[...checks,'serialized-proof-bound-to-exact-candidate-hash'],
  evidence:operations.at(-1).engineEvidence.packageEvidence,notVerified:['volatile-Word-revision-identities','original-package-metadata-not-exposed-by-SDK','visual-layout']};
}
