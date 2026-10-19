/** Preserve portable runtime settings without forwarding credentials to workers. */
export function workerEnvironment(environment=process.env){
 const result={PATH:environment.PATH??'',LANG:'C.UTF-8'}
 for(const name of ['TMPDIR','TMP','TEMP'])if(typeof environment[name]==='string'&&environment[name])result[name]=environment[name]
 if(environment.ELECTRON_RUN_AS_NODE==='1')result.ELECTRON_RUN_AS_NODE='1'
 // A host that embeds Bun in a multi-mode executable selects JavaScript mode this way.
 if(environment.BUN_BE_BUN==='1')result.BUN_BE_BUN='1'
 return result
}
/** Runtime flags for a forked worker. The V8 heap ceiling is a Node option; a Bun
 * host that forbids dependency auto-install keeps that rule in its workers. */
export function workerExecArgv({heapLimitMb}={}){
 if(process.versions.bun)return process.execArgv.includes('--no-install')?['--no-install']:[]
 return heapLimitMb?[`--max-old-space-size=${heapLimitMb}`]:[]
}
