/** Configure the installed SDK executable without rewriting its permissions. */
import fs from 'node:fs';
import {createRequire} from 'node:module';

const require=createRequire(import.meta.url);
const targets=new Set(['darwin-arm64','darwin-x64','linux-arm64','linux-x64','windows-x64']);

export function installedSdkBinary({platform=process.platform,arch=process.arch,
  resolve,stat=file=>fs.statSync(file),access=(file,mode)=>fs.accessSync(file,mode)}={}) {
  const target=`${platform==='win32'?'windows':platform}-${arch}`;
  if(!targets.has(target))return undefined;
  try {
    // Resolve from the SDK's dependency tree, rather than a host's hoisted
    // platform package. This also works with isolated dependency layouts.
    const resolveNative=resolve??createRequire(require.resolve('@superdoc/sdk')).resolve;
    const file=resolveNative(`@superdoc/sdk-${target}/bin/${platform==='win32'?'superdoc.exe':'superdoc'}`);
    if(!stat(file).isFile())return undefined;
    access(file,platform==='win32'?fs.constants.F_OK:fs.constants.X_OK);
    return file;
  }catch {
    // Preserve SDK resolution/permission recovery and its normal errors when
    // the packaged executable is unavailable or not executable yet.
    return undefined;
  }
}
