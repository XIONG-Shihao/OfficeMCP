import fs from 'node:fs';
import path from 'node:path';
import { digest, fingerprint, newId, sourcePath, regularFile } from '../../infrastructure/files.mjs';
import { requireThat } from '../../infrastructure/errors.mjs';
import { viewOf, withTargets } from './docx-service.mjs';
import { renderWordReference, inspectLayoutPdf, compareLayoutPixels } from './docx-layout-provider.mjs';
const normalized = value => String(value).normalize('NFKC').replace(/\s+/gu, '');

export async function layoutDocument(service, args, signal) {
  const current = viewOf(service, args.view);
  const provider = args.provider === 'configured' ? (service.config.engines.word ? 'word' : 'libreoffice') : args.provider;
  requireThat(provider === 'word' ? service.config.engines.word : service.config.engines.libreOffice,
    'ENGINE_UNAVAILABLE', `Configure engines.${provider === 'word' ? 'word' : 'libreOffice'} before requesting layout.`);
  const output = service.internal('exports', newId('layout') + '.pdf');
  try {
    let identity;
    if (provider === 'word') identity = await renderWordReference(current.file, output, { application: service.config.engines.word, scratchRoot: service.config.engines.wordScratch ?? service.internal("exports", "word-renderer"), signal, timeoutMs: service.config.workerTimeoutMs });
    else {
      const result = await service.worker('office_export_pdf', current.file, {}, { format: 'docx', output, bypassDocx: true }, signal);
      identity = { provider: 'libreoffice', version: result.engine, localReference: false };
    }
    viewOf(service, args.view);
    requireThat(!signal?.aborted, 'CANCELLED', 'Layout cancelled.');
    const layout = inspectLayoutPdf(output, service.config.maxFileBytes);
    const id = 'layout_' + fingerprint({ view: args.view, identity, pdfRevision: layout.pdfRevision }).slice(7);
    const record = { documentId: args.view.documentId, view: args.view, identity, ...layout };
    requireThat(layout.pages.reduce((sum, page) => sum + page.text.length, 0) <= 2000000, 'LIMIT_EXCEEDED', 'Layout text exceeds the two-million-character budget.');
    service.put('docxLayout', id, record);
    const retained = service.rows('docxLayout').filter(row => row.documentId === args.view.documentId).sort((a, b) => b.updatedAt - a.updatedAt);
    for (const expired of retained.slice(4)) service.delete('docxLayout', expired.id);
    let appearance;
    if (args.reference) {
      requireThat(args.reference.documentRevision === args.view.revision, 'REVISION_CONFLICT', 'Word reference belongs to a different DOCX revision.');
      const reference = sourcePath(service.config, args.reference.path);
      requireThat(path.extname(reference).toLowerCase() === '.pdf', 'INVALID_ARGUMENT', 'Word reference must be a PDF.');
      appearance = { reference: { producer: 'operator-supplied Microsoft Word reference', version: args.reference.producerVersion, pdfRevision: digest(regularFile(reference, service.config.maxFileBytes)) },
        ...compareLayoutPixels(reference, output, { maxBytes: service.config.maxFileBytes }) };
    }
    const pages = layout.pages.filter(page => !args.page || page.number === args.page);
    requireThat(!args.page || pages.length, 'TARGET_CONFLICT', 'Page does not exist in this layout.');
    return { view: args.view, data: { layoutId: id, provider: identity, documentRevision: args.view.revision, pdfRevision: layout.pdfRevision, pageCount: layout.pageCount,
      pages: pages.slice(args.offset, args.offset + args.limit).map(page => ({ number: page.number, bounds: page.bounds, text: page.text.slice(0, args.maxChars), textTruncated: page.text.length > args.maxChars })),
      nextOffset: args.offset + args.limit < pages.length ? args.offset + args.limit : null, ...(appearance ? { appearance } : {}),
      nextAction: 'Query a page with this layoutId, page number and a sufficiently precise literal pattern. Edits use returned semantic targets; changed revisions require new layout.',
      notVerified: ['arbitrary-Word-layout-equivalence', 'page-targeting-of-opaque-objects', 'font-equivalence-across-machines'] } };
  } finally { fs.rmSync(output, { force: true }); }
}

export async function queryPage(service, args, signal) {
  viewOf(service, args.view);
  const layout = service.get('docxLayout', args.select.layoutId);
  requireThat(fingerprint(layout.view) === fingerprint(args.view), 'TARGET_CONFLICT', 'Layout target is stale or belongs to another document. Render the current view again.');
  const page = layout.pages.find(page => page.number === args.select.number);
  requireThat(page, 'TARGET_CONFLICT', 'Page is absent from the recorded layout.');
  const pattern = normalized(args.select.pattern);
  requireThat(pattern && normalized(page.text).includes(pattern), 'TARGET_CONFLICT', 'Literal text is absent from the selected rendered page.');
  const pages = layout.pages.filter(page => normalized(page.text).includes(pattern));
  requireThat(pages.length === 1, 'TARGET_CONFLICT', 'Rendered text occurs on multiple pages. Use a longer unique phrase or a semantic heading/table target.');
  const current = viewOf(service, args.view);
  // Match the PDF observation to one real SDK range. A page label or a guessed
  // positional offset is never sufficient to authorize a mutation.
  const result = await service.worker('docx_native_agent_query', current.file, { ...args,
    select: { type: 'text', pattern: args.select.pattern, mode: 'contains', caseSensitive: true }, require: 'exactlyOne' }, {}, signal);
  viewOf(service, args.view);
  return { view: args.view, data: { ...withTargets(service, args.view, result), layoutId: args.select.layoutId, page: args.select.number, provider: layout.identity },
    coverage: { completeDocument: false, notChecked: ['cross-page-ranges', 'opaque-object-targets', 'ambiguous-rendered-text'] } };
}
