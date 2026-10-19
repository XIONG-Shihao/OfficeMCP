/** External headless engine sources. No editor or application source is vendored here.
 * See docs/ENGINE_SOURCES.md for the checkout layout and qualified revision.
 */
import {existsSync} from 'node:fs';
import {execFileSync} from 'node:child_process';
import path from 'node:path';
import {fileURLToPath, pathToFileURL} from 'node:url';

const siblingRoot = fileURLToPath(new URL('../../', import.meta.url));
export const xlsxSource = path.resolve(process.env.OFFICEMCP_XLSX_ENGINE_SOURCE ?? path.join(siblingRoot, 'editor/xlsx'));
export const docxSource = path.resolve(process.env.OFFICEMCP_DOCX_ENGINE_SOURCE ?? path.join(siblingRoot, 'editor/docx'));
async function sourceModule(directory, name, setting) {
  const file = path.join(directory, name);
  if (!existsSync(file)) throw Error(`Missing external engine source: ${file}. Set ${setting} to the component source directory. See docs/ENGINE_SOURCES.md.`);
  return import(pathToFileURL(file).href);
}
const [nativeBuild, workbookExport, nativeExport, docxExport, fixtures, memory] = await Promise.all([
  sourceModule(xlsxSource, 'scripts/build-native.mjs', 'OFFICEMCP_XLSX_ENGINE_SOURCE'),
  sourceModule(xlsxSource, 'scripts/export-core.mjs', 'OFFICEMCP_XLSX_ENGINE_SOURCE'),
  sourceModule(xlsxSource, 'scripts/export-native.mjs', 'OFFICEMCP_XLSX_ENGINE_SOURCE'),
  sourceModule(docxSource, 'scripts/export-core.mjs', 'OFFICEMCP_DOCX_ENGINE_SOURCE'),
  sourceModule(xlsxSource, 'scripts/test-support.mjs', 'OFFICEMCP_XLSX_ENGINE_SOURCE'),
  sourceModule(xlsxSource, 'scripts/support/process-memory.mjs', 'OFFICEMCP_XLSX_ENGINE_SOURCE'),
]);
export const {exportWorkbookCore} = workbookExport;
export const {exportNativeWorkbook} = nativeExport;
export const {exportDocxCore} = docxExport;
export const {fixtureDirectory, fixturePath, oraclePath} = fixtures;
export const {sampleSpreadsheetMemory} = memory;
const exe = process.platform === 'win32' ? '.exe' : '';
export const nativeSource = path.join(xlsxSource, 'native/bin', `${process.platform}-${process.arch}`, 'rpf-workbook' + exe);

/** Build the native workbook helper and (on macOS) the DOCX process observer. */
export function buildEnginePrograms() {
  nativeBuild.buildNativeWorkbook();
  if (process.platform === 'darwin') execFileSync(process.execPath, [path.join(docxSource, 'scripts/prepare-observer.mjs')], {stdio: 'inherit'});
}
export function requireEnginePrograms() {
  if (!existsSync(nativeSource)) throw Error('The spreadsheet engine program is not built. Run: node scripts/prepare-core.mjs');
}
