import fs from 'node:fs';
import {createHash} from 'node:crypto';
import {fileURLToPath} from 'node:url';
import path from 'node:path';
import {fileDigest} from '#document-core/xlsx/indexed-cache.mjs';
import {IndexedWorkbook} from '#document-core/xlsx/indexed-session.mjs';
import {calculateIndexedWorkbook} from '#document-core/xlsx/indexed-calculation.mjs';
import {createWorkbookNativeClient} from '#document-core/xlsx/native-client.mjs';
import {WorkbookSession,readWorkbookRange} from '#document-core/xlsx/workbook-service.mjs';
import {columnName,rangeBounds} from '#document-core/xlsx/workbook-commands.mjs';
import {digest,fileRevision} from '../../infrastructure/files.mjs';
import {requireThat} from '../../infrastructure/errors.mjs';
import {nativeWorkbookBinary} from './native-binary.mjs';

let cacheIdentity;
async function identity(binary){
 if(!cacheIdentity)cacheIdentity=(async()=>{
  const hash=createHash('sha256').update(await fileDigest(binary)).update(process.versions.bun ? `bun:${process.versions.bun}` : process.versions.node).update(await fileDigest(fileURLToPath(new URL('../../../npm-shrinkwrap.json',import.meta.url))));
  const root=path.dirname(fileURLToPath(import.meta.resolve('#document-core/xlsx/indexed-session.mjs')));
  for(const name of fs.readdirSync(root).filter(n=>n.endsWith('.mjs')).sort())hash.update(name).update(fs.readFileSync(path.join(root,name)));
  return hash.digest('hex');
 })();return cacheIdentity;
}
async function session(source,run) {
  const binary=nativeWorkbookBinary();requireThat(binary,'UNSUPPORTED_FEATURE','Native workbook helper is unavailable for this platform.');
  const client=createWorkbookNativeClient(binary);
  const book=new IndexedWorkbook(client.call,async(source,commands,output)=>{const session=new WorkbookSession(fs.readFileSync(source)),result=await session.apply({operationId:'compatibility',expectedRevision:digest(fs.readFileSync(source)),commands});await client.call({action:'workbook',operation:'reserve-candidate',root:path.dirname(output),bytes:result.bytes.byteLength});try{fs.writeFileSync(output,result.bytes)}finally{await client.call({action:'workbook',operation:'release-candidate',root:path.dirname(output)})}return {revision:result.view.revision,edits:result.edits}},input=>calculateIndexedWorkbook(input,client.call));
  try{const view=await book.open(source,{cacheIdentity:await identity(binary)});return await run(book,view);}finally{await book.close();client.close();}
}
export async function readNativeWorkbook(task) {
  return session(task.input,async(book,view)=>{
    let args=task.args;
    if(args.table){const table=view.tables.find(t=>t.name===args.table&&(!args.sheet||t.sheet===args.sheet));requireThat(table,'TARGET_CONFLICT','Named table was not found in this workbook view.');
      if(args.range){const a=rangeBounds(table.ref),b=rangeBounds(args.range);requireThat(b.first.row>=a.first.row&&b.first.col>=a.first.col&&b.last.row<=a.last.row&&b.last.col<=a.last.col,'TARGET_CONFLICT','Requested range is outside the named table.');}
      args={...args,sheet:table.sheet,range:args.range??table.ref};
    }
    const calculation=await book.calculationSummary('0');
    requireThat(!args.search||args.sheet,'INVALID_ARGUMENT','Search requires a sheet or table.');
    requireThat(!args.summary||args.sheet,'INVALID_ARGUMENT','Selection summary requires a sheet or table.');
    if(!args.sheet)return {sheets:view.sheets.map(({cells,...sheet})=>({...sheet,usedCells:sheet.storedCells})),namedRanges:view.namedRanges,tables:view.tables,limitations:view.limitations,readOnly:view.readOnly,calculation,storage:'rust-indexed-v1',indexCacheHit:book.cacheHit};
    const sheet=view.sheets.find(s=>s.name===args.sheet);requireThat(sheet,'TARGET_CONFLICT','Sheet not found');
    const range=args.range??'A1:J20',bounds=rangeBounds(range),offset=args.offset??0,limit=args.limit??100;
    requireThat(Number.isInteger(offset)&&offset>=0&&Number.isInteger(limit)&&limit>=1&&limit<=500,'INVALID_ARGUMENT','Invalid read window');
    if(args.search){
      requireThat(!args.summary&&!args.keys&&!offset,'INVALID_ARGUMENT','Search cannot combine with keys, summary or offset.');
      const page=await book.query('0',args.sheet,args.search.text,args.search.mode??'display',args.search.after,{range:args.range??'A1:XFD1048576',matchLimit:limit});
      sheet.cells=page.cells;const items=Object.keys(page.cells).flatMap(ref=>readWorkbookRange(view,{sheet:args.sheet,range:ref,limit:1}).items);
      return {items,matches:page.matches,scanned:page.scanned,nextCursor:page.nextCursor,storage:'rust-indexed-v1',indexCacheHit:book.cacheHit,calculation,coverage:{sheet:args.sheet,range:args.range??'A1:XFD1048576',mode:args.search.mode??'display',completeSearch:page.nextCursor===null,completeDocument:false}};
    }
    if(args.summary){
      requireThat(!args.keys,'INVALID_ARGUMENT','Selection summaries cannot be combined with keyed reads.');
      return {summary:await book.selectionSummary('0',args.sheet,range),coverage:{sheet:args.sheet,range,completeDocument:false},calculation,storage:'rust-indexed-v1',indexCacheHit:book.cacheHit};
    }
    if(args.keys){
      const rowOffset=Math.floor(offset/bounds.columns),page=await book.keys('0',args.sheet,range,args.keys.column,args.keys.values,{offset:rowOffset,limit:Math.ceil(limit/bounds.columns)+1});
      const total=page.count*bounds.columns,items=[];
      for(let at=offset;at<Math.min(total,offset+limit);){
        const row=page.rows[Math.floor(at/bounds.columns)-rowOffset],col=bounds.first.col+at%bounds.columns,n=Math.min(bounds.columns-at%bounds.columns,offset+limit-at,total-at);
        const readRange=`${columnName(col)}${row+1}:${columnName(col+n-1)}${row+1}`;
        const read=await book.read('0',args.sheet,readRange);sheet.cells=read.cells;
        items.push(...readWorkbookRange(view,{sheet:args.sheet,range:readRange,limit:n}).items);at+=n;
      }
      return {items,total,offset,nextOffset:offset+items.length<total?offset+items.length:null,
        missingKeys:args.keys.values.filter((_,i)=>page.matches[i]===0),duplicateKeys:args.keys.values.filter((_,i)=>page.matches[i]>1),calculation,storage:'rust-indexed-v1',indexCacheHit:book.cacheHit,
        coverage:{sheet:args.sheet,range,keys:args.keys,completeDocument:false,omissions:['unrequested-ranges','external-data-freshness','visual-layout']}};
    }
    // Fetch only output positions. Adjacent positions on one row share a request.
    for(let at=offset;at<Math.min(bounds.size,offset+limit);){const row=bounds.first.row+Math.floor(at/bounds.columns),col=bounds.first.col+at%bounds.columns,n=Math.min(bounds.columns-at%bounds.columns,offset+limit-at,bounds.size-at);
      const result=await book.read('0',args.sheet,`${columnName(col)}${row+1}:${columnName(col+n-1)}${row+1}`);Object.assign(sheet.cells,result.cells);at+=n;
    }
    return {...readWorkbookRange(view,args),calculation,storage:'rust-indexed-v1',indexCacheHit:book.cacheHit};
  });
}
export async function applyNativeWorkbook(task) {
  return session(task.input,async(book,view)=>{
    // The engine removes only the filter it is shown. Bind the caller's intent to
    // the filter in this exact workbook view; replay resolves it the same way.
    const commands=task.args.commands.map(command=>{
      if(command.operation!=='removeFilter')return command;
      const filter=view.sheets.find(sheet=>sheet.name===command.sheet)?.filter;
      requireThat(filter,'TARGET_CONFLICT','This worksheet has no filter to remove.');
      requireThat(filter.owned&&filter.revision,'UNSUPPORTED_FEATURE','This filter was not created by this tool, so its hidden rows cannot be restored safely. '+(filter.reason??''));
      return {...command,expectedFilterRevision:filter.revision};
    });
    const result=await book.apply('0',commands),state=result.view.storage.state;
    const {affected,total}=await book.affected('0',state),calculation=await book.calculationSummary(state);
    const saved=await book.export(state);fs.copyFileSync(saved.path,task.output,fs.constants.COPYFILE_EXCL);
    const revision=fileRevision(task.output,Infinity);calculation.revision=revision;
    for(const target of affected)if(target.after.calculation){const calculated=target.after.calculation;calculated.revision=revision;
      if(target.after.formula)target.after.value=calculated.status==='calculated'?(calculated.value??''):calculated.status==='error'?calculated.error:null;
    }
    return {revision,affectedTargets:affected,engine:'Rust indexed workbook + Univer OSS 0.25.2',evaluation:false,
      postconditions:{affectedCellCount:total,affectedCellsTruncated:total>affected.length,calculation,transformations:task.args.commands.filter(c=>c.operation).map(({operation,sheet,index,count,range})=>({operation,sheet,index,count,range}))}};
  });
}
export async function validateNativeWorkbook(task,bytes) {
  let source=task.original;const files=[];
  try{
    for(const [index,operation]of task.operations.entries()){
      const output=`${task.scratch}/native-replay-${index}.xlsx`;files.push(output);
      await applyNativeWorkbook({...task,input:source,output,args:{commands:operation.commands}});source=output;
    }
    requireThat(fileRevision(source,Infinity)===(bytes?digest(bytes):fileRevision(task.input,Infinity)),'VALIDATION_FAILED','Workbook differs from its accepted native command batch.');
    return {passed:true,revision:bytes?digest(bytes):fileRevision(task.input,Infinity),profile:'xlsx.native.v1',checks:['guarded-native-command-replay','saved-formulas-and-calculation','native-preserving-export'],notVerified:['visual-layout','external-data-freshness','opaque-object-semantics']};
  }finally{for(const file of files)fs.rmSync(file,{force:true});}
}
