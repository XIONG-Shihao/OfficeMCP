/** Optional layout providers; document edits remain in the native SDK. */
import fs from 'node:fs';
import path from 'node:path';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { fileURLToPath } from 'node:url';
import * as mupdf from 'mupdf';
import { regularFile, digest } from '../../infrastructure/files.mjs';
import { requireThat, OfficeError } from '../../infrastructure/errors.mjs';
import { openFilePackage } from '../opc/file-package.mjs';
import { parseXml } from '../opc/ooxml.mjs';
const execute = promisify(execFile);
const script = fileURLToPath(new URL('word-render.applescript', import.meta.url));
let wordQueue = Promise.resolve();
export async function renderWordReference(input, output, { application, scratchRoot = path.dirname(output), signal, timeoutMs = 120000 } = {}) {
  requireThat(process.platform === 'darwin' && application && fs.existsSync(application), 'ENGINE_UNAVAILABLE', 'Configure engines.word with an installed Microsoft Word.app on macOS.');
  requireThat(!fs.existsSync(output), 'DESTINATION_CONFLICT', 'Layout output already exists.');
  const archive = openFilePackage(input, 1024 * 1024 * 1024, 512 * 1024 * 1024);
  try {
    for (const name of Object.keys(archive.parts)) {
      requireThat(!/vbaProject/i.test(name), 'UNSUPPORTED_FEATURE', 'Word reference rendering does not execute macro-enabled packages.');
      if (name.endsWith('.rels')) {
        const rels = parseXml(Buffer.from(archive.parts[name]).toString('utf8'));
        requireThat(!Array.from(rels.getElementsByTagNameNS('*', 'Relationship')).some(r => r.getAttribute('TargetMode') === 'External' && !r.getAttribute('Type').endsWith('/hyperlink')), 'UNSUPPORTED_FEATURE', 'Word reference rendering refuses external templates, media and data sources.');
      }
      if (/^word\/.*\.xml$/.test(name)) {
        const tree = parseXml(Buffer.from(archive.parts[name]).toString('utf8'));
        const instructions = [...Array.from(tree.getElementsByTagNameNS('*', 'instrText')).map(n => n.textContent), ...Array.from(tree.getElementsByTagNameNS('*', 'fldSimple')).map(n => n.getAttributeNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'instr'))];
        requireThat(!instructions.some(s => /\b(?:DDEAUTO|DDE|LINK|INCLUDETEXT|INCLUDEPICTURE|DATABASE)\b/i.test(s)), 'UNSUPPORTED_FEATURE', 'Word reference rendering refuses fields that fetch or execute external content.');
      }
    }
  } finally { archive.close(); }
  const operation = wordQueue.catch(() => {}).then(async () => {
    requireThat(!signal?.aborted, 'CANCELLED', 'Word rendering cancelled.');
    // The operator grants Word access only to this private render directory.
    // Never expose the source root or change the original document.
    const root = path.resolve(scratchRoot); fs.mkdirSync(root, { recursive: true, mode: 0o700 });
    const marker = path.join(root, '.officemcp-word-renderer.json');
    const copy = path.join(root, 'OfficeMCP-word-input.docx'), pdf = path.join(root, 'OfficeMCP-word-reference.pdf');
    if (!fs.existsSync(marker)) {
      requireThat(!fs.existsSync(copy) && !fs.existsSync(pdf), 'DESTINATION_CONFLICT', 'Word scratch files already exist without OfficeMCP ownership. Choose a new private render directory.');
      fs.writeFileSync(marker, JSON.stringify({ kind: 'officemcp.word-renderer.v1' }), { flag: 'wx', mode: 0o600 });
    }
    requireThat(JSON.parse(fs.readFileSync(marker)).kind === 'officemcp.word-renderer.v1', 'ACCESS_DENIED', 'Word scratch ownership is invalid.');
    const lock = path.join(root, '.word-render.lock');
    try { fs.mkdirSync(lock, { mode: 0o700 }); } catch { throw new OfficeError('BUSY', 'Another Word render owns this private directory.'); }
    try {
      for (const file of [copy, pdf]) requireThat(!fs.existsSync(file) || fs.lstatSync(file).isFile(), 'ACCESS_DENIED', 'Word render files must be owned regular files.');
      await execute('/usr/bin/osascript', [script, copy, pdf, 'close', application], { signal, timeout: timeoutMs, maxBuffer: 65536 });
      fs.copyFileSync(input, copy); fs.chmodSync(copy, 0o600); fs.rmSync(pdf, { force: true });
      await execute('/usr/bin/osascript', [script, copy, pdf, 'open', application], { signal, timeout: timeoutMs, maxBuffer: 65536 });
      let stdout;
      const deadline = Date.now() + timeoutMs;
      while (stdout === undefined) {
        try { ({ stdout } = await execute('/usr/bin/osascript', [script, copy, pdf, 'export', application], { signal, timeout: Math.max(1, deadline - Date.now()), maxBuffer: 65536 })); }
        catch (error) {
          if (!error.stderr?.includes('(-2701)') || signal?.aborted || Date.now() >= deadline) throw error;
          await new Promise(resolve => setTimeout(resolve, 100));
        }
      }
      requireThat(fs.existsSync(pdf), 'ENGINE_FAILED', 'Word produced no PDF; check local file permissions, activation and dialogs.');
      fs.copyFileSync(pdf, output, fs.constants.COPYFILE_EXCL); fs.chmodSync(output, 0o600);
      return { provider: 'microsoft-word-macos', version: stdout.trim(), localReference: true };
    } catch (error) {
      if (error instanceof OfficeError) throw error;
      const failure = new OfficeError(signal?.aborted ? 'CANCELLED' : 'ENGINE_FAILED', 'Microsoft Word reference rendering failed or timed out. Check local activation, Automation/file permissions and dialogs; no document edit was accepted.');
      failure.cause = error; throw failure;
    } finally {
      await execute('/usr/bin/osascript', [script, copy, pdf, 'close', application], { timeout: 10000, maxBuffer: 65536 }).catch(() => {});
      for (const file of [copy, pdf]) if (fs.existsSync(file) && fs.lstatSync(file).isFile()) fs.rmSync(file);
      fs.rmSync(lock, { recursive: true, force: true });
    }
  });
  wordQueue = operation; return operation;
}
export function inspectLayoutPdf(file, maxBytes) {
  const bytes = regularFile(file, maxBytes), document = new mupdf.PDFDocument(bytes);
  try {
    document.disableJS(); const count = document.countPages();
    requireThat(count > 0 && count <= 2000, 'LIMIT_EXCEEDED', 'Layout exceeds the 2000-page budget.');
    const pages = [];
    for (let index = 0; index < count; index++) {
      const page = document.loadPage(index), text = page.toStructuredText('preserve-whitespace');
      try {
        const blocks = JSON.parse(text.asJSON()).blocks ?? [];
        pages.push({ number: index + 1, bounds: page.getBounds(), text: blocks.flatMap(b => b.lines ?? []).map(l => l.text ?? '').join('\n') });
      } finally { text.destroy(); page.destroy(); }
    }
    return { pdfRevision: digest(bytes), pageCount: count, pages };
  } finally { document.destroy(); }
}
export function compareLayoutPixels(referenceFile, candidateFile, { maxBytes, dpi = 96 } = {}) {
  requireThat(Number.isFinite(dpi) && dpi > 0 && dpi <= 300, 'INVALID_ARGUMENT', 'Comparison DPI must be within 0–300.');
  const reference = new mupdf.PDFDocument(regularFile(referenceFile, maxBytes));
  const candidate = new mupdf.PDFDocument(regularFile(candidateFile, maxBytes));
  try {
    reference.disableJS(); candidate.disableJS();
    requireThat(reference.countPages() <= 2000 && candidate.countPages() <= 2000, 'LIMIT_EXCEEDED', 'Appearance exceeds the page budget.');
    const pages = [], count = Math.max(reference.countPages(), candidate.countPages());
    for (let i = 0; i < count; i++) {
      if (i >= reference.countPages() || i >= candidate.countPages()) { pages.push({ page: i + 1, equal: false, reason: 'missing-page' }); continue; }
      const left = reference.loadPage(i), right = candidate.loadPage(i);
      let a, b;
      try {
        for (const page of [left, right]) { const [x0,y0,x1,y1] = page.getBounds(); requireThat((x1-x0)*(y1-y0)*(dpi/72)**2 <= 16000000, 'LIMIT_EXCEEDED', 'Appearance page exceeds the sixteen-million-pixel budget.'); }
        const matrix = [dpi / 72, 0, 0, dpi / 72, 0, 0];
        a = left.toPixmap(matrix, mupdf.ColorSpace.DeviceRGB, false, true); b = right.toPixmap(matrix, mupdf.ColorSpace.DeviceRGB, false, true);
        const dimensionsEqual = a.getWidth() === b.getWidth() && a.getHeight() === b.getHeight();
        const x = a.getPixels(), y = b.getPixels(); let differentBytes = 0;
        if (dimensionsEqual && x.length === y.length) for (let n = 0; n < x.length; n++) if (x[n] !== y[n]) differentBytes++;
        pages.push({ page: i + 1, equal: dimensionsEqual && x.length === y.length && differentBytes === 0, dimensionsEqual, differentBytes: dimensionsEqual ? differentBytes : null });
      } finally { a?.destroy(); b?.destroy(); left.destroy(); right.destroy(); }
    }
    return { identical: pages.every(p => p.equal), dpi, referencePages: reference.countPages(), candidatePages: candidate.countPages(), pages,
      scope: 'Exact rendered pixels at the recorded DPI and MuPDF version for these two PDF artifacts; no universal Word or other-platform equivalence.' };
  } finally { reference.destroy(); candidate.destroy(); }
}
