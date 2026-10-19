/**
 * Build the engines from their sources and place the runtime files OfficeMCP
 * loads under .generated/ (git-ignored). Run it after any engine change;
 * `npm test`, `npm run build` and the repository test runner do it for you.
 *
 *   node scripts/prepare-core.mjs                   build and export
 *   node scripts/prepare-core.mjs --check           fail if .generated differs from the sources
 *   node scripts/prepare-core.mjs --into DIRECTORY  export somewhere else (the app build does this)
 */
import {fileURLToPath,pathToFileURL} from 'node:url';
import path from 'node:path';
import {buildEnginePrograms,requireEnginePrograms,exportWorkbookCore,exportNativeWorkbook,exportDocxCore} from './engine-source.mjs';
export const generated=fileURLToPath(new URL('../.generated/',import.meta.url));
export function prepareEngines({check=false,into=generated}={}) {
  if(check)requireEnginePrograms();else buildEnginePrograms();
  return {
    xlsxCore:exportWorkbookCore(path.join(into,'document-core'),{check}),
    native:exportNativeWorkbook(path.join(into,'native'),{check}),
    docx:exportDocxCore(path.join(into,'docx-core'),{check})
  };
}
if(process.argv[1]&&pathToFileURL(path.resolve(process.argv[1])).href===import.meta.url) {
  const argv=process.argv.slice(2),at=argv.indexOf('--into');
  prepareEngines({check:argv.includes('--check'),...(at<0?{}:{into:path.resolve(argv[at+1])})});
}
