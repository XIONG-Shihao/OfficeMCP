import {existsSync} from 'node:fs';
import {fileURLToPath} from 'node:url';
export function nativeWorkbookBinary() {
  const executable=process.platform==='win32'?'rpf-workbook.exe':'rpf-workbook';
  const file=fileURLToPath(new URL(`../../../.generated/native/${process.platform}-${process.arch}/${executable}`,import.meta.url));
  return existsSync(file)?file:null;
}
export function nativeCommands(commands) {
  return Array.isArray(commands)&&commands.length>0&&commands.every(c=>c.operation
    ? ['sortRows','filterRows','removeFilter','insertRow','deleteRow','insertColumn','deleteColumn','resizeRow','resizeColumn','freeze','mergeCells','unmergeCells','addSheet','renameSheet','duplicateSheet','createTable','addConditionalRule'].includes(c.operation)
    : [undefined,'recalculate','setCells','formatRange','clearRange','pasteValues','updateRows','fillRange','fillAcrossSheets','clearHyperlinks','fillSeries','pasteFormulas'].includes(c.type));
}
