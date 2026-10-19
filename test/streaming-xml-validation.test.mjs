import test from './support/test.mjs';
import assert from 'node:assert/strict';
import {zipSync,strToU8} from 'fflate';
import {validateXml,decodePackage} from '../src/formats/opc/ooxml.mjs';
import {packageEvidence} from '../src/formats/opc/engine-evidence.mjs';
test('streamed XML preserves syntax, namespace, entity and UTF-8 safety across chunk boundaries',()=>{
  for(const xml of ['<a xmlns="urn:x"><b x="&amp;">é文🙂</b></a>','<?xml version="1.0"?><a><!--valid comment--><![CDATA[plain < text]]></a>'])assert.doesNotThrow(()=>validateXml(Buffer.from(xml)));
  const prefix='<a>'+ 'x'.repeat(65530);
  assert.doesNotThrow(()=>validateXml(Buffer.from(prefix+'🙂é文</a>')));
  for(const xml of ['<a><b></a>','<a x="1" x="2"/>','<p:a/>','<a>&unknown;</a>','<a/><b/>',prefix+'<!DOCTYPE a></a>','<!DOCTYPE a [<!ENTITY x "content">]><a>&x;</a>'])assert.throws(()=>validateXml(Buffer.from(xml)),{code:'INVALID_DOCUMENT'});
  assert.throws(()=>validateXml(Buffer.from([60,97,62,255,60,47,97,62])),{code:'INVALID_DOCUMENT'});
});
test('package preservation rejects malformed large document XML without retaining a body DOM',()=>{
  const w='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const parts={'[Content_Types].xml':strToU8('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>'),
    'word/document.xml':strToU8(`<w:document xmlns:w="${w}"><w:body>${'<w:p><w:r><w:t>unchanged</w:t></w:r></w:p>'.repeat(3000)}</w:body></w:document>`)};
  const bytes=Buffer.from(zipSync(parts)),before=decodePackage(bytes,20*1024*1024),after=decodePackage(bytes,20*1024*1024);
  const evidence=packageEvidence(before,after,20*1024*1024);assert(evidence.checks.includes('output-XML-well-formed'));
  assert(!after.xml.has('word/document.xml'),'Syntax-only validation must not retain the body model');
  parts['word/document.xml']=strToU8(`<w:document xmlns:w="${w}"><w:body><w:p></w:body></w:document>`);
  assert.throws(()=>packageEvidence(before,Buffer.from(zipSync(parts)),20*1024*1024),{code:'INVALID_DOCUMENT'});
});
