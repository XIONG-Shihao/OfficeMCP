import { requireThat } from '../../infrastructure/errors.mjs';
import {nativeWorkbookBinary,nativeCommands} from './native-binary.mjs';

export async function queryWorkbook(service, args, signal) {
  const view = service.view({ documentId: args.view.documentId, candidateId: args.view.candidateId,
    expectedRevision: args.view.revision });
  requireThat(['xlsx', 'csv'].includes(view.d.format), 'UNSUPPORTED_FORMAT', 'This operation requires an XLSX workbook or CSV table.');
  requireThat(!args.search||view.d.format==='xlsx'&&nativeWorkbookBinary(),'UNSUPPORTED_FEATURE','Search requires the native XLSX helper.');
  requireThat(!args.summary||view.d.format==='xlsx'&&nativeWorkbookBinary(),'UNSUPPORTED_FEATURE','Selection summaries require the native XLSX helper.');
  const data = await service.worker(`${view.d.format}_oss_read`, view.file, args, { format: view.d.format }, signal);
  if (data.items) data.items = data.items.map(item => ({ ...item, target: { ...item.target, ...args.view },
    ...(item.calculation ? { calculation: { ...item.calculation, revision: args.view.revision } } : {}) }));
  if (data.calculation) data.calculation.revision = args.view.revision;
  return { view: args.view, data };
}

export async function applyWorkbook(service, args, signal) {
  const document = service.doc(args.view.documentId);
  requireThat(['xlsx', 'csv'].includes(document.format), 'UNSUPPORTED_FORMAT', 'This operation requires an XLSX workbook or CSV table.');
  requireThat(args.view.candidateId || args.view.revision === document.revision, 'REVISION_CONFLICT', 'Read the current workbook view before editing.');
  const hasClearHyperlinks=args.commands.some(command=>command.type==='clearHyperlinks');
  let commands = (hasClearHyperlinks?structuredClone(args.commands):args.commands).map(command => {
    if (!command.target) return command;
    const { target, ...edit } = command;
    requireThat(target.documentId === args.view.documentId && target.candidateId === args.view.candidateId,
      'TARGET_CONFLICT', 'Target belongs to another document or candidate; query this view.');
    requireThat(target.revision === args.view.revision, 'TARGET_CONFLICT', 'Target belongs to another workbook revision; query again.');
    return { ...edit, sheet: target.sheet, cell: target.cell };
  });
  if(hasClearHyperlinks){
    requireThat(document.format==='xlsx'&&nativeWorkbookBinary()&&nativeCommands(commands),'UNSUPPORTED_FEATURE','Clear Hyperlinks requires the native XLSX helper; CSV and compatibility writers are unsupported.');
    const {normalizeClearHyperlinksCommands}=await import('#document-core/xlsx/clear-hyperlinks.mjs');
    try{commands=normalizeClearHyperlinksCommands(commands);}catch(error){requireThat(false,'INVALID_ARGUMENT',error.message);}
  }
  if(commands.some(command=>command.operation==='removeFilter'))requireThat(document.format==='xlsx'&&nativeWorkbookBinary()&&nativeCommands(commands),'UNSUPPORTED_FEATURE','Removing a filter requires the native XLSX helper; CSV and compatibility writers are unsupported.');
  if(commands.some(command=>command.type==='fillAcrossSheets'))requireThat(document.format==='xlsx'&&nativeWorkbookBinary()&&nativeCommands(commands),'UNSUPPORTED_FEATURE','Across Worksheets requires the native XLSX helper; CSV and compatibility writers are unsupported.');
  const result = await service.operation(`${document.format}_oss_apply`, {
    documentId: args.view.documentId, expectedSourceRevision: document.revision,
    ...(args.view.candidateId ? { candidateId: args.view.candidateId, expectedCandidateRevision: args.view.revision } : {}),
    operationId: args.operationId, preservationProfile: document.format === 'csv' ? 'csv.literals.v1' : nativeWorkbookBinary()&&nativeCommands(commands)?'xlsx.native.v1':'xlsx.oss.v1', commands,
  }, signal, { transition: true });
  return { ...result, data: { ...result.data, affectedTargets: result.data.affectedTargets.map(item => ({ ...item,
    target: { sheet: item.sheet, cell: item.cell, documentId: args.view.documentId,
      candidateId: result.data.candidateId, revision: result.data.candidateRevision } })) },
    view: { documentId: args.view.documentId, candidateId: result.data.candidateId,
    revision: result.data.candidateRevision } };
}
