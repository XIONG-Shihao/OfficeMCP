import test from './support/test.mjs';import assert from 'node:assert/strict';
import {zipSync,strToU8,unzipSync} from 'fflate';
import {proveSerializedHeading,serializedHeadingStep,boundSerializedValidation,normalizedSerializedXml} from '../src/formats/docx/docx-serialized-heading-proof.mjs';
const step={action:'insert_heading',arguments:{text:'New & exact',level:2,placement:{at:'document_end'},changeMode:'direct'}};
const root='<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>',end='</w:body></w:document>';
const old='<w:p><w:r><w:t>Original</w:t></w:r></w:p>',section='<w:sectPr><w:pgSz w:w="100"/></w:sectPr>';
const heading='<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>New &amp; exact</w:t></w:r></w:p>';
function file(body,{media='exact binary'}={}){return zipSync({'[Content_Types].xml':strToU8('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'),'word/document.xml':strToU8(root+body+end),'word/styles.xml':strToU8('<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:styleId="Heading2"><w:name w:val="heading 2"/></w:style></w:styles>'),'word/media/image.png':strToU8(media)});}
const baseline=file(old+section),candidate=file(old+heading+section);
test('one exact body-end heading proves existing XML bytes and unrelated binary unchanged',()=>{
 const proof=proveSerializedHeading(baseline,candidate,step,1024*1024);assert.equal(proof.level,2);assert.equal(proof.text,'New & exact');assert.equal(serializedHeadingStep({args:{steps:[step]}}),step);
 for(const args of [{steps:[step,step]},{steps:[{...step,arguments:{...step.arguments,changeMode:'tracked'}}]},{steps:[step],readback:{mode:'body'}}])assert.equal(serializedHeadingStep({args}),null);
});
test('proof refuses wrong text/level/location, unrelated changes, extra paragraphs and tracked domains',()=>{
 for(const bad of [file(old.replace('Original','Changed')+heading+section),file(old+heading+section,{media:'changed'}),file(heading+old+section),file(old+heading+heading+section),file(old+heading.replace('exact','different')+section),file(old+heading.replace('<w:r>','<w:ins><w:r>').replace('</w:r>','</w:r></w:ins>')+section)])assert.throws(()=>proveSerializedHeading(baseline,bad,step,1024*1024));
 assert.throws(()=>proveSerializedHeading(baseline,candidate,{...step,arguments:{...step.arguments,level:3}},1024*1024));
 const parts=unzipSync(candidate);parts['word/unrelated.xml']=strToU8('<x/>');assert.throws(()=>proveSerializedHeading(baseline,zipSync(parts),step,1024*1024));
});
test('cached evidence must bind the complete host-owned revision chain',()=>{
 const proof=proveSerializedHeading(baseline,candidate,step,1024*1024),input='sha256:'+'a'.repeat(64);
 proof.checks.push('source-document-domains-preserved-through-baseline-serialization');
 const op={engineEvidence:{verification:'superdoc-serialized-heading-v1',inputRevision:input,outputRevision:proof.outputRevision,serializedProof:proof,packageEvidence:{inputRevision:input,outputRevision:proof.outputRevision,checks:['checked']}}};
 assert(boundSerializedValidation([op],input,proof.outputRevision).passed);assert.throws(()=>boundSerializedValidation([op],input,'changed'));
 const bad=structuredClone(op);bad.engineEvidence.packageEvidence.outputRevision='changed';assert.throws(()=>boundSerializedValidation([bad],input,proof.outputRevision));assert.equal(boundSerializedValidation([{engineEvidence:{verification:'superdoc-sdk-readback-v1'}}],input,proof.outputRevision),null);
});

test('baseline XML normalization keeps authored values, significant spaces and namespace-specific attributes',()=>{
 const x='<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:rsidR="a"><w:r><w:t> exact </w:t></w:r></w:p>';
 const norm=s=>normalizedSerializedXml(Buffer.from(s));
 assert.equal(norm(x),norm(x.replace('w:rsidR="a"','w:rsidR="b"')));
 assert.notEqual(norm(x),norm(x.replace(' exact ','exact')));
 assert.notEqual(norm(x),norm(x.replace(' exact ','["open","w:r",[]]')));
 const other=x.replace('http://schemas.openxmlformats.org/wordprocessingml/2006/main','other');
 assert.notEqual(norm(other),norm(other.replace('w:rsidR="a"','w:rsidR="b"')));
 assert.equal(norm(x.replace('<w:r>','<w:r a="x" b="y">')),norm(x.replace('<w:r>','<w:r b="y" a="x">')));
 assert.throws(()=>norm(x.replace('<w:r>','<!--x--><w:r>')));
});
