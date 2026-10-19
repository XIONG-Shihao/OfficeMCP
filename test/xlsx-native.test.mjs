import {fixturePath} from '../scripts/engine-source.mjs';
import {zipSync,unzipSync,strToU8,strFromU8} from 'fflate';
import test from './support/test.mjs';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import {createWorkbook,readWorkbook,editWorkbook} from '#document-core/xlsx/workbook-package.mjs';
import {createServiceFixture} from './service-fixture.mjs';
import {nativeWorkbookBinary} from '../src/formats/xlsx/native-binary.mjs';

test('native MCP profile reads, guards, retries, validates and saves compact bulk edits',async()=>{
 assert(nativeWorkbookBinary(),'Prepare this platform native helper before testing');
 const fixture=createServiceFixture({prefix:'xlsx-native-profile-',limits:{workerTimeoutMs:60000}});
 try{
  const source=path.join(fixture.docs,'source.xlsx');fs.writeFileSync(source,createWorkbook());const original=fs.readFileSync(source);
  const opened=await fixture.call('office_open',{path:source}),view={documentId:opened.document.id,revision:opened.document.sourceRevision};
  const inventory=await fixture.call('xlsx_query',{view});assert.equal(inventory.data.storage,'rust-indexed-v1');
  const request={view,operationId:'bulk-native',commands:[{sheet:'Sheet1',cell:'A1',value:4},{type:'formatRange',sheet:'Sheet1',range:'A1:A12000',format:{bold:true}}]};
  let edited=await fixture.call('xlsx_apply',request);assert.equal(fixture.service.view({documentId:edited.view.documentId,candidateId:edited.view.candidateId,expectedRevision:edited.view.revision}).c.operations[0].preservationProfile,'xlsx.native.v1');assert.equal(edited.data.postconditions.affectedCellCount,12000);
  const retry=await fixture.call('xlsx_apply',request);assert.equal(retry.view.revision,edited.view.revision);assert.equal(retry.view.candidateId,edited.view.candidateId);
  const tail=await fixture.call('xlsx_query',{view:edited.view,sheet:'Sheet1',range:'A12000'});assert.equal(tail.data.items[0].bold,true);
  edited=await fixture.call('xlsx_apply',{view:edited.view,operationId:'formula-then-paste',commands:[{sheet:'Sheet1',cell:'B1',formula:'2+3'},{type:'pasteValues',sheet:'Sheet1',source:'B1',range:'C1'}]});
  const pasted=await fixture.call('xlsx_query',{view:edited.view,sheet:'Sheet1',range:'C1'});assert.equal(pasted.data.items[0].value,5);
  const refused=await fixture.service.call('xlsx_apply',{view:edited.view,operationId:'conflict',commands:[{sheet:'Sheet1',cell:'A1',value:9,expectedValue:999}]});assert.equal(refused.isError,true);
  const saved=await fixture.call('office_save',{view:edited.view,operationId:'save-native',destination:{rootId:'output',relativePath:'saved.xlsx'}});
  const book=readWorkbook(fs.readFileSync(saved.data.path));assert.equal(book.sheets[0].cells.A1.value,4);assert.equal(book.sheets[0].cells.A12000.bold,true);assert.deepEqual(fs.readFileSync(source),original);
 }finally{fixture.close();}
});
test('native MCP keyed reads and guarded updates reuse only immutable revision indexes',async()=>{
 const fixture=createServiceFixture({prefix:'xlsx-native-keyed-',limits:{workerTimeoutMs:60000}});
 try{
  const source=path.join(fixture.docs,'source.xlsx');fs.writeFileSync(source,createWorkbook());
  const opened=await fixture.call('office_open',{path:source}),view={documentId:opened.document.id,revision:opened.document.sourceRevision};
  const populated=await fixture.call('xlsx_apply',{view,operationId:'populate',commands:[{sheet:'Sheet1',cell:'A1',value:'SKU'},{sheet:'Sheet1',cell:'B1',value:'Price'},{sheet:'Sheet1',cell:'A2',value:'alpha'},{sheet:'Sheet1',cell:'B2',value:10},{sheet:'Sheet1',cell:'A3',value:'beta'},{sheet:'Sheet1',cell:'B3',value:20}]});
  const query={view:populated.view,sheet:'Sheet1',range:'A2:B3',keys:{column:'A',values:['beta','missing']},limit:1,offset:1};
  const first=await fixture.call('xlsx_query',query),second=await fixture.call('xlsx_query',query);
  assert.equal(first.data.storage,'rust-indexed-v1');assert.equal(second.data.indexCacheHit,true);assert.equal(second.data.items[0].cell,'B3');assert.equal(second.data.items[0].value,20);assert.deepEqual(second.data.missingKeys,['missing']);assert.equal(second.data.total,2);
  const updated=await fixture.call('xlsx_apply',{view:populated.view,operationId:'keyed',commands:[{type:'updateRows',sheet:'Sheet1',range:'A2:B3',keyColumn:'A',updates:[{key:'beta',cells:[{column:'B',value:25,expectedValue:20}]}]}]});
  const state=fixture.service.view({documentId:updated.view.documentId,candidateId:updated.view.candidateId,expectedRevision:updated.view.revision});assert.equal(state.c.operations.at(-1).preservationProfile,'xlsx.native.v1');
  const saved=await fixture.call('office_save',{view:updated.view,operationId:'save-keyed',destination:{rootId:'output',relativePath:'keyed.xlsx'}});assert.equal(readWorkbook(fs.readFileSync(saved.data.path)).sheets[0].cells.B3.value,25);
  const old=await fixture.call('xlsx_query',{view,sheet:'Sheet1',range:'A1'});assert.equal(old.data.items[0].value,null);
 }finally{fixture.close();}
});
test('native MCP shares fill, sort, filter and structural owner with bounded replay receipts',async()=>{
 const fixture=createServiceFixture({prefix:'xlsx-native-structure-',limits:{workerTimeoutMs:60000}});
 try{
  const source=path.join(fixture.docs,'source.xlsx');fs.writeFileSync(source,createWorkbook());const original=fs.readFileSync(source);
  const opened=await fixture.call('office_open',{path:source});let view={documentId:opened.document.id,revision:opened.document.sourceRevision};
  const apply=async(id,commands)=>{const request={view,operationId:id,commands},result=await fixture.call('xlsx_apply',request);const retry=await fixture.call('xlsx_apply',request);assert.equal(retry.view.revision,result.view.revision);const state=fixture.service.view({documentId:result.view.documentId,candidateId:result.view.candidateId,expectedRevision:result.view.revision});assert.equal(state.c.operations.at(-1).preservationProfile,'xlsx.native.v1');view=result.view;return result};
  await apply('populate',[{sheet:'Sheet1',cell:'A1',value:'Key'},{sheet:'Sheet1',cell:'B1',value:'Value'},{sheet:'Sheet1',cell:'A2',value:30},{sheet:'Sheet1',cell:'B2',value:3},{sheet:'Sheet1',cell:'A3',value:10},{sheet:'Sheet1',cell:'B3',value:1}]);
  await apply('fill',[{type:'fillRange',sheet:'Sheet1',source:'B3',range:'B4:B1200'}]);
  const sorted=await apply('sort',[{operation:'sortRows',sheet:'Sheet1',range:'A2:B3',column:'A',direction:'asc'}]);assert.equal(sorted.data.postconditions.affectedCellCount,4);assert.equal(sorted.data.postconditions.transformations[0].operation,'sortRows');
  const deleted=await apply('delete',[{operation:'deleteRow',sheet:'Sheet1',index:1199}]);assert.equal(deleted.data.postconditions.affectedCellCount,1);assert.equal(deleted.data.affectedTargets[0].cell,'B1200');
  await apply('insert',[{operation:'insertRow',sheet:'Sheet1',index:1}]);
  await apply('column-insert',[{operation:'insertColumn',sheet:'Sheet1',index:1}]);
  await apply('column-delete',[{operation:'deleteColumn',sheet:'Sheet1',index:1}]);
  // Clearing a filter keeps it switched on (as Excel does). The filter leaves
  // extension records the engine's structural writer cannot keep correct, so
  // structural edits on that sheet are refused; the filter steps follow them.
  const filtered=await apply('filter',[{operation:'filterRows',sheet:'Sheet1',range:'A1:B4',column:'A',values:[30]}]);assert.equal(filtered.data.postconditions.affectedCellCount,0);assert.equal(filtered.data.postconditions.transformations[0].operation,'filterRows');
  await apply('unfilter',[{operation:'filterRows',sheet:'Sheet1',range:'A1:B4',column:'A',values:[]}]);
  // The refusal is OfficeMCP's unsupported-feature result; its sentence is the engine's.
  const beforeRefusal=view;
  await assert.rejects(apply('refused-while-filtered',[{operation:'deleteRow',sheet:'Sheet1',index:1300}]),/"code":"UNSUPPORTED_FEATURE"/);
  assert.deepEqual(view,beforeRefusal,'A refused structural edit leaves the accepted view unchanged');
  // Removing the filter switches it off entirely; structure edits work again.
  const removed=await apply('remove-filter',[{operation:'removeFilter',sheet:'Sheet1'}]);assert.equal(removed.data.postconditions.transformations[0].operation,'removeFilter');
  await assert.rejects(apply('remove-missing-filter',[{operation:'removeFilter',sheet:'Sheet1'}]),/no filter to remove/);
  await apply('delete-after-remove',[{operation:'deleteRow',sheet:'Sheet1',index:1300}]);
  await apply('freeze',[{operation:'freeze',sheet:'Sheet1',rows:1,columns:1}]);
  const saved=await fixture.call('office_save',{view,operationId:'save-structure',destination:{rootId:'output',relativePath:'structure.xlsx'}});const book=readWorkbook(fs.readFileSync(saved.data.path));assert.equal(book.sheets[0].cells.A3.value,10);assert.equal(book.sheets[0].cells.A4.value,30);assert.equal(book.sheets[0].cells.B1200.value,1);assert.deepEqual(fs.readFileSync(source),original);
 }finally{fixture.close();}
});

test('native MCP reference-aware columns preserve names and replay saved calculations',async()=>{
 const fixture=createServiceFixture({prefix:'xlsx-native-column-names-',limits:{workerTimeoutMs:60000}});
 try{
  const source=path.join(fixture.docs,'source.xlsx'),parts=unzipSync(createWorkbook());
  parts['xl/workbook.xml']=strToU8(strFromU8(parts['xl/workbook.xml']).replace('</workbook>','<definedNames><definedName name="Numbers">Sheet1!$A$1:$B$1</definedName></definedNames></workbook>'));
  parts['xl/worksheets/sheet1.xml']=strToU8('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>2</v></c><c r="B1"><v>3</v></c><c r="C1"><f>SUM(Numbers)</f></c></row></sheetData></worksheet>');
  fs.writeFileSync(source,zipSync(parts));const original=fs.readFileSync(source),opened=await fixture.call('office_open',{path:source});
  const moved=await fixture.call('xlsx_apply',{view:{documentId:opened.document.id,revision:opened.document.sourceRevision},operationId:'column',commands:[{sheet:'Sheet1',operation:'insertColumn',index:1}]});
  const changed=await fixture.call('xlsx_apply',{view:moved.view,operationId:'input',commands:[{sheet:'Sheet1',cell:'A1',expectedValue:2,value:7}]});
  const read=await fixture.call('xlsx_query',{view:changed.view,sheet:'Sheet1',range:'D1'});assert.equal(read.data.items[0].calculation.value,10);
  const saved=await fixture.call('office_save',{view:changed.view,operationId:'save-column',destination:{rootId:'output',relativePath:'column.xlsx'}});const book=readWorkbook(fs.readFileSync(saved.data.path));assert.equal(book.sheets[0].cells.D1.formula,'SUM(Numbers)');assert.equal(book.sheets[0].cells.D1.cachedValue,10);assert.equal(book.namedRanges[0].formula,'Sheet1!$A$1:$C$1');assert.deepEqual(fs.readFileSync(source),original);
 }finally{fixture.close();}
});

test('native MCP canonicalizes table shorthand across guarded candidates, retries and save replay',async()=>{
 const fixture=createServiceFixture({prefix:'xlsx-native-shorthand-',limits:{workerTimeoutMs:60000}});
 try{
  let bytes=editWorkbook(createWorkbook(),[['A1','Value'],['B1','Result'],['A2',2],['A3',3]].map(([cell,value])=>({sheet:'Sheet1',cell,value})));
  bytes=editWorkbook(bytes,[{sheet:'Sheet1',operation:'createTable',range:'A1:B3',tableName:'Prices'}]);
  const source=path.join(fixture.docs,'source.xlsx');fs.writeFileSync(source,bytes);
  const opened=await fixture.call('office_open',{path:source});
  const request={view:{documentId:opened.document.id,revision:opened.document.sourceRevision},operationId:'shorthand',commands:[{sheet:'Sheet1',cell:'B2',formula:'[@Value]*2'},{sheet:'Sheet1',cell:'B3',formula:'prices[@value]+1'}]};
  const result=await fixture.call('xlsx_apply',request),retry=await fixture.call('xlsx_apply',request);assert.equal(result.view.revision,retry.view.revision);
  const read=await fixture.call('xlsx_query',{view:result.view,sheet:'Sheet1',range:'B2:B3'});
  assert.equal(read.data.items[0].formula,'Prices[[#This Row],[Value]]*2');assert.equal(read.data.items[0].calculation.value,4);
  const guarded=await fixture.call('xlsx_apply',{view:result.view,operationId:'guarded-shorthand',commands:[{sheet:'Sheet1',cell:'B2',expectedFormula:'Prices[[#This Row],[Value]]*2',formula:'Prices[@Value]*3'}]});
  const invalid=await fixture.service.call('xlsx_apply',{view:guarded.view,operationId:'invalid-shorthand',commands:[{sheet:'Sheet1',cell:'B2',formula:'[@Missing]'}]});assert.equal(invalid.isError,true);
  const saved=await fixture.call('office_save',{view:guarded.view,operationId:'save-shorthand',destination:{rootId:'output',relativePath:'shorthand.xlsx'}});
  const disk=readWorkbook(fs.readFileSync(saved.data.path));assert.equal(disk.sheets[0].cells.B2.formula,'Prices[[#This Row],[Value]]*3');assert.equal(disk.sheets[0].cells.B2.cachedValue,6);assert.equal(disk.sheets[0].cells.B3.cachedValue,4);assert.deepEqual(new Uint8Array(fs.readFileSync(source)),new Uint8Array(bytes));
 }finally{fixture.close();}
});

test('native MCP formula fill, formula-only paste and series share compact candidate replay and guard atomicity',async()=>{
 const fixture=createServiceFixture({prefix:'xlsx-native-bulk-formulas-',limits:{workerTimeoutMs:60000}});
 try{
  const source=path.join(fixture.docs,'source.xlsx');fs.writeFileSync(source,createWorkbook());const original=fs.readFileSync(source);
  const opened=await fixture.call('office_open',{path:source});let view={documentId:opened.document.id,revision:opened.document.sourceRevision};
  const request={view,operationId:'native-formula-fill',commands:[{sheet:'Sheet1',cell:'A1',value:1},{sheet:'Sheet1',cell:'A2',value:2},{sheet:'Sheet1',cell:'B1',formula:'POWER(A1,2)'},{type:'fillSeries',sheet:'Sheet1',range:'A1:A100'},{type:'fillRange',sheet:'Sheet1',source:'B1',range:'B2:B100'},{type:'pasteFormulas',sheet:'Sheet1',source:'B1:B3',range:'C1:C3'}]};
  const result=await fixture.call('xlsx_apply',request);assert.equal((await fixture.call('xlsx_apply',request)).view.revision,result.view.revision);view=result.view;
  const state=fixture.service.view({documentId:view.documentId,candidateId:view.candidateId,expectedRevision:view.revision});assert.equal(state.c.operations.at(-1).preservationProfile,'xlsx.native.v1');assert.equal(state.c.operations.at(-1).commands.length,6);
  const read=await fixture.call('xlsx_query',{view,sheet:'Sheet1',range:'B100'});assert.equal(read.data.items[0].calculation.value,10000);
  const refused=await fixture.service.call('xlsx_apply',{view,operationId:'series-overwrite',commands:[{sheet:'Sheet1',cell:'D1',value:99},{type:'fillSeries',sheet:'Sheet1',range:'A1:A100'}]});assert.equal(refused.isError,true);
  const guarded=await fixture.call('xlsx_apply',{view,operationId:'input-edit',commands:[{sheet:'Sheet1',cell:'A3',expectedValue:3,value:4}]});
  const saved=await fixture.call('office_save',{view:guarded.view,operationId:'save-bulk',destination:{rootId:'output',relativePath:'bulk.xlsx'}});
  const disk=readWorkbook(fs.readFileSync(saved.data.path));assert.equal(disk.sheets[0].cells.B3.cachedValue,16);assert.equal(disk.sheets[0].cells.C3.cachedValue,256);assert.equal(disk.sheets[0].cells.D1,undefined);assert.deepEqual(fs.readFileSync(source),original);
 }finally{fixture.close();}
});

test('native MCP selection summaries are complete, bounded and revision-scoped',async()=>{
 const fixture=createServiceFixture({prefix:'xlsx-native-summary-',limits:{workerTimeoutMs:60000}});
 try{
  const source=path.join(fixture.docs,'source.xlsx');fs.writeFileSync(source,createWorkbook());const opened=await fixture.call('office_open',{path:source}),original={documentId:opened.document.id,revision:opened.document.sourceRevision};
  const result=await fixture.call('xlsx_apply',{view:original,operationId:'numbers',commands:[{sheet:'Sheet1',cell:'A1',value:1},{sheet:'Sheet1',cell:'A2',value:2},{type:'fillSeries',sheet:'Sheet1',range:'A1:A10000'}]});
  const query=await fixture.call('xlsx_query',{view:result.view,sheet:'Sheet1',range:'A1:A10000',summary:true});assert.equal(query.data.summary.sum,50005000);assert.equal(query.data.summary.numericCount,10000);assert.equal(query.data.items,undefined);assert(JSON.stringify(query).length<3000);
  assert.equal((await fixture.call('xlsx_query',{view:original,sheet:'Sheet1',range:'A1:A10000',summary:true})).data.summary.count,0);
  const invalid=await fixture.service.call('xlsx_query',{view:result.view,sheet:'Sheet1',summary:true,keys:{column:'A',values:[1]}});assert(invalid.isError);
 }finally{fixture.close()}
});

test('native MCP formula sort and multirow deletion preserve candidate replay and guards',async()=>{
 const fixture=createServiceFixture({prefix:'xlsx-native-advanced-',limits:{workerTimeoutMs:60000}});
 try{
  const parts=unzipSync(createWorkbook());parts['xl/worksheets/sheet1.xml']=strToU8('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'+[4,2,3,1].map((n,i)=>`<row r="${i+1}"><c r="A${i+1}"><v>${n}</v></c><c r="B${i+1}"><f>A${i+1}*2</f></c></row>`).join('')+'</sheetData></worksheet>');parts['custom/opaque.bin']=new Uint8Array([0,1,255]);const original=Buffer.from(zipSync(parts)),source=path.join(fixture.docs,'source.xlsx');fs.writeFileSync(source,original);
  const opened=await fixture.call('office_open',{path:source});let view={documentId:opened.document.id,revision:opened.document.sourceRevision};
  for(const [id,command] of [['sort',{operation:'sortRows',sheet:'Sheet1',range:'A1:B4',column:'B',direction:'asc'}],['delete',{operation:'deleteRow',sheet:'Sheet1',index:1,count:2}]]){
   const request={view,operationId:id,commands:[command]},result=await fixture.call('xlsx_apply',request);assert.equal((await fixture.call('xlsx_apply',request)).view.revision,result.view.revision);view=result.view;
   const state=fixture.service.view({documentId:view.documentId,candidateId:view.candidateId,expectedRevision:view.revision});assert.equal(state.c.operations.at(-1).preservationProfile,'xlsx.native.v1');
  }
  const query=await fixture.call('xlsx_query',{view,sheet:'Sheet1',range:'A1:B2'});assert.deepEqual(query.data.items.filter(c=>c.formula).map(c=>[c.formula,c.calculation.value]),[['A1*2',2],['A2*2',8]]);
  const refused=await fixture.service.call('xlsx_apply',{view,operationId:'bounds',commands:[{operation:'deleteRow',sheet:'Sheet1',index:1048575,count:2}]});assert(refused.isError);
  const saved=await fixture.call('office_save',{view,operationId:'save-advanced',destination:{rootId:'output',relativePath:'advanced.xlsx'}});const disk=readWorkbook(fs.readFileSync(saved.data.path));assert.equal(disk.sheets[0].cells.A2.value,4);assert.equal(disk.sheets[0].cells.B2.cachedValue,8);assert(!disk.sheets[0].cells.A3);assert.deepEqual(fs.readFileSync(source),original);assert.deepEqual(unzipSync(fs.readFileSync(saved.data.path))['custom/opaque.bin'],parts['custom/opaque.bin']);
 }finally{fixture.close()}
});

test('native MCP display search pages every match and shares table columns and sheet rename',async()=>{
 const fixture=createServiceFixture({prefix:'xlsx-native-query-metadata-',limits:{workerTimeoutMs:60000}});
 try{
  const source=path.join(fixture.docs,'source.xlsx');let bytes=editWorkbook(createWorkbook(),[['A1','Qty'],['B1','Price'],['A2',2],['B2',3],['A3',4],['B3',5]].map(([cell,value])=>({sheet:'Sheet1',cell,value})));fs.writeFileSync(source,bytes);const opened=await fixture.call('office_open',{path:source});let view={documentId:opened.document.id,revision:opened.document.sourceRevision};
  const result=await fixture.call('xlsx_apply',{view,operationId:'metadata-batch',commands:[{operation:'createTable',sheet:'Sheet1',range:'A1:B3',tableName:'Sales'},{sheet:'Sheet1',cell:'D1',formula:'SUM(Sales[Price])'},{operation:'insertColumn',sheet:'Sheet1',index:1},{type:'formatRange',sheet:'Sheet1',range:'C2:C3',format:{numberFormat:'currency'}},{operation:'renameSheet',sheet:'Sheet1',newName:'Budget'}]});view=result.view;
  const found=[];let after;do{const query=await fixture.call('xlsx_query',{view,table:'Sales',search:{text:'$',mode:'display',...(after?{after}:{})},limit:1});assert(query.data.items.length<=1);found.push(...query.data.items.map(c=>c.cell));after=query.data.nextCursor;}while(after);assert.deepEqual(found,['C2','C3']);
  const removed=await fixture.call('xlsx_apply',{view,operationId:'table-delete',commands:[{operation:'deleteColumn',sheet:'Budget',index:2}]});const read=await fixture.call('xlsx_query',{view:removed.view,sheet:'Budget',range:'D1'});assert.equal(read.data.items[0].formula,'SUM(#REF!)');
  const bad=await fixture.service.call('xlsx_query',{view,sheet:'Budget',search:{text:'$',mode:'display'},summary:true});assert(bad.isError);
  const saved=await fixture.call('office_save',{view:removed.view,operationId:'save-metadata',destination:{rootId:'output',relativePath:'metadata.xlsx'}});const disk=readWorkbook(fs.readFileSync(saved.data.path));assert.equal(disk.sheets[0].name,'Budget');assert.deepEqual(disk.tables[0].columns.map(c=>c.name),['Qty','Column1']);assert.deepEqual(new Uint8Array(fs.readFileSync(source)),new Uint8Array(bytes));
 }finally{fixture.close()}
});

test('MCP native runtime targets recalculate after retargeting and formula-sheet copy saves independent results',async()=>{
 const fixture=createServiceFixture({prefix:'xlsx-runtime-targets-',limits:{workerTimeoutMs:60000}});
 try{const source=path.join(fixture.docs,'source.xlsx');fs.writeFileSync(source,createWorkbook());const opened=await fixture.call('office_open',{path:source});let view={documentId:opened.document.id,revision:opened.document.sourceRevision};
 let changed=await fixture.call('xlsx_apply',{view,operationId:'runtime-create',commands:[{sheet:'Sheet1',cell:'A1',value:'B1'},{sheet:'Sheet1',cell:'B1',value:.1},{sheet:'Sheet1',cell:'C1',value:.2},{sheet:'Sheet1',cell:'D1',formula:'INDIRECT(A1)*C1'}]});view=changed.view;
 let read=await fixture.call('xlsx_query',{view,sheet:'Sheet1',range:'D1'});assert.equal(read.data.items[0].calculation.value,.1*.2);
 changed=await fixture.call('xlsx_apply',{view,operationId:'runtime-retarget',commands:[{sheet:'Sheet1',cell:'A1',value:'C1'}]});view=changed.view;read=await fixture.call('xlsx_query',{view,sheet:'Sheet1',range:'D1'});assert.equal(read.data.items[0].calculation.value,.2*.2);
 changed=await fixture.call('xlsx_apply',{view,operationId:'runtime-copy',commands:[{sheet:'Sheet1',cell:'D1',formula:'Sheet1!B1/C1'},{operation:'duplicateSheet',sheet:'Sheet1',newName:'Copy'},{sheet:'Copy',cell:'B1',value:.3}]});view=changed.view;
 const saved=await fixture.call('office_save',{view,operationId:'runtime-save',destination:{rootId:'output',relativePath:'saved.xlsx'}});const disk=readWorkbook(fs.readFileSync(saved.data.path));assert.equal(disk.sheets.find(s=>s.name==='Sheet1').cells.D1.cachedValue,.1/.2);assert.equal(disk.sheets.find(s=>s.name==='Copy').cells.D1.cachedValue,.3/.2);assert.equal(disk.sheets.find(s=>s.name==='Copy').cells.D1.formula,"'Copy'!B1/C1");
 }finally{fixture.close()}
});

test('MCP uses typed native selectors, B7 cancellation and complete special-formula replacement',async()=>{
 const fixture=createServiceFixture({prefix:'xlsx-typed-special-',limits:{workerTimeoutMs:60000}});
 try{const source=path.join(fixture.docs,'source.xlsx'),parts=unzipSync(createWorkbook());parts['xl/worksheets/sheet1.xml']=strToU8('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f t="array" ref="A1:A2">1+2</f><v>3</v></c><c r="D1"><v>9</v></c></row><row r="2"><c r="A2"><v>3</v></c></row></sheetData></worksheet>');fs.writeFileSync(source,zipSync(parts));const original=fs.readFileSync(source),opened=await fixture.call('office_open',{path:source});let view={documentId:opened.document.id,revision:opened.document.sourceRevision};
 const rejected=await fixture.service.call('xlsx_apply',{view,operationId:'partial',commands:[{type:'pasteValues',sheet:'Sheet1',source:'D1',range:'A1'}]});assert(rejected.isError);
 let changed=await fixture.call('xlsx_apply',{view,operationId:'replace-and-formulas',commands:[{type:'pasteValues',sheet:'Sheet1',source:'D1',range:'A1:A2'},{sheet:'Sheet1',cell:'B1',formula:'LEFT("Axxx",1)&2'},{sheet:'Sheet1',cell:'C1',formula:'INDIRECT(B1)'},{sheet:'Sheet1',cell:'F1',formula:'0.1+0.2-0.3'},{sheet:'Sheet1',cell:'F2',formula:'(0.1+0.2-0.3)'}]});view=changed.view;
 changed=await fixture.call('xlsx_apply',{view,operationId:'change-target',commands:[{sheet:'Sheet1',cell:'A2',value:11}]});view=changed.view;const query=await fixture.call('xlsx_query',{view,sheet:'Sheet1',range:'C1'});assert.equal(query.data.items[0].calculation.value,11);
 const saved=await fixture.call('office_save',{view,operationId:'save-typed',destination:{rootId:'output',relativePath:'typed.xlsx'}}),disk=readWorkbook(fs.readFileSync(saved.data.path));assert.equal(disk.sheets[0].cells.A1.formulaType,undefined);assert.equal(disk.sheets[0].cells.B1.cachedValue,'A2');assert.equal(disk.sheets[0].cells.C1.cachedValue,11);assert.equal(disk.sheets[0].cells.F1.cachedValue,0);assert.equal(disk.sheets[0].cells.F2.cachedValue,2**-54);assert.deepEqual(fs.readFileSync(source),original);
 }finally{fixture.close()}
});

test('native MCP partial shared edits, explicit coercion and empty-table deletion preserve revision guards',async()=>{
 const fixture=createServiceFixture({prefix:'xlsx-shared22-',limits:{workerTimeoutMs:60000}});
 try{for(const [file,sheet,commands,expected]of [
 ['shared-edit22-source.xlsx','Sheet1',[{sheet:'Sheet1',cell:'A2',value:50,expectedFormula:null,expectedRawValue:'999'},{sheet:'Sheet1',cell:'E1',formula:'N(TRUE)+LEN(REPT(T("ab"),3))'}],74],
 ['delete-table-source.xlsx','Data',[{operation:'deleteRow',sheet:'Data',index:1,count:2}],0]
 ]){const source=path.join(fixture.docs,file);fs.copyFileSync(fixturePath(file),source);const original=fs.readFileSync(source);const opened=await fixture.call('office_open',{path:source}),view={documentId:opened.document.id,revision:opened.document.sourceRevision};const changed=await fixture.call('xlsx_apply',{view,operationId:'change-'+file,commands});const saved=await fixture.call('office_save',{view:changed.view,operationId:'save-'+file,destination:{rootId:'output',relativePath:'saved-'+file}}),book=readWorkbook(fs.readFileSync(saved.data.path));assert.equal(book.sheets[0].cells.D1.cachedValue,expected);if(sheet==='Sheet1'){assert.equal(book.sheets[0].cells.A3.formula,'$B3+C$1');assert.equal(book.sheets[0].cells.E1.cachedValue,7)}else assert.equal(book.tables[0].ref,'A1:B2');assert.deepEqual(fs.readFileSync(source),original);}
 }finally{fixture.close()}
});

test('native MCP retires complete What-If result groups and evaluates qualified integer coercion',async()=>{
 const fixture=createServiceFixture({prefix:'xlsx-data-table23-',limits:{workerTimeoutMs:60000}});
 try{const source=path.join(fixture.docs,'table.xlsx');fs.copyFileSync(fixturePath('excel-data-table23.xlsx'),source);const original=fs.readFileSync(source);const opened=await fixture.call('office_open',{path:source}),view={documentId:opened.document.id,revision:opened.document.sourceRevision};
 const partial=await fixture.service.call('xlsx_apply',{view,operationId:'partial23',commands:[{sheet:'Data',cell:'D3',value:99}]});assert(partial.isError);
 const result=await fixture.call('xlsx_apply',{view,operationId:'full23',commands:[{type:'pasteValues',sheet:'Data',source:'F1',range:'D2:D4'},{sheet:'Data',cell:'H1',value:'12'},{sheet:'Data',cell:'I1',formula:'H1+TRUE'}]});const saved=await fixture.call('office_save',{view:result.view,operationId:'save23',destination:{rootId:'output',relativePath:'saved-table.xlsx'}});const book=readWorkbook(fs.readFileSync(saved.data.path));for(const ref of ['D2','D3','D4']){assert.equal(book.sheets[0].cells[ref].value,9);assert.equal(book.sheets[0].cells[ref].formulaType,undefined)}assert.equal(book.sheets[0].cells.I1.cachedValue,13);assert.equal(book.sheets[0].cells.A1.value,2);assert.deepEqual(fs.readFileSync(source),original);
 }finally{fixture.close()}
});

test('native MCP recalculates What-If scenarios and preserves Excel metadata and inputs',async()=>{
 const fixture=createServiceFixture({prefix:'xlsx-scenarios24-',limits:{workerTimeoutMs:60000}});
 try{const source=path.join(fixture.docs,'scenarios.xlsx');fs.copyFileSync(fixturePath('excel-scenarios-coercion24.xlsx'),source);const original=fs.readFileSync(source);const opened=await fixture.call('office_open',{path:source}),view={documentId:opened.document.id,revision:opened.document.sourceRevision};const changed=await fixture.call('xlsx_apply',{view,operationId:'scenario24',commands:[{sheet:'Two',cell:'E5',value:30},{type:'formatRange',sheet:'Two',range:'F5:H6',format:{bold:true}},{sheet:'Coercion',cell:'C1',formula:'("0.1"+"0.2"-"0.3")'}]});const saved=await fixture.call('office_save',{view:changed.view,operationId:'scenario-save24',destination:{rootId:'output',relativePath:'recalculated.xlsx'}});const book=readWorkbook(fs.readFileSync(saved.data.path)),two=book.sheets.find(s=>s.name==='Two');assert.equal(two.cells.J1.cachedValue,492);assert.equal(two.cells.F5.cachedValue,93);assert.equal(two.cells.F5.formulaType,'dataTable');assert.equal(two.cells.A1.value,2);assert.equal(two.cells.A2.value,5);assert.equal(book.sheets.find(s=>s.name==='Coercion').cells.C1.cachedValue,2**-54);assert.deepEqual(fs.readFileSync(source),original);
 }finally{fixture.close()}
});

test('explicit MCP recalculation stays on the native owner and preserves guarded replay',async()=>{
 const fixture=createServiceFixture({prefix:'xlsx-native-recalculate-',limits:{workerTimeoutMs:60000}});
 try{const source=path.join(fixture.docs,'source.xlsx');fs.writeFileSync(source,createWorkbook());const opened=await fixture.call('office_open',{path:source});const r=await fixture.call('xlsx_apply',{view:{documentId:opened.document.id,revision:opened.document.sourceRevision},operationId:'native-recalc',commands:[{sheet:'Sheet1',cell:'A1',value:3},{sheet:'Sheet1',cell:'B1',formula:'A1+2'},{type:'recalculate'}]});const state=fixture.service.view({documentId:r.view.documentId,candidateId:r.view.candidateId,expectedRevision:r.view.revision});assert.equal(state.c.operations[0].preservationProfile,'xlsx.native.v1');assert.equal((await fixture.call('xlsx_query',{view:r.view,sheet:'Sheet1',range:'B1'})).data.items[0].calculation.value,5);await fixture.call('office_save',{view:r.view,operationId:'save-recalc',destination:{rootId:'output',relativePath:'result.xlsx'}});}finally{fixture.close()}
});
