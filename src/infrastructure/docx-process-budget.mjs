/** Process-tree RSS, distributed identically in the independent MCP packages. */
import { execFile } from 'node:child_process'
import { promisify } from 'node:util'
const execute=promisify(execFile)

export function selectProcessTree(rows,roots,groups=[]) {
  const ids=new Set(roots),ownedGroups=new Set(groups)
  for(const row of rows)if(ownedGroups.has(row.group))ids.add(row.pid)
  for(let changed=true;changed;){changed=false;for(const row of rows)if(ids.has(row.parent)&&!ids.has(row.pid)){ids.add(row.pid);changed=true}}
  const processes=rows.filter(row=>ids.has(row.pid))
  return{rssBytes:processes.reduce((sum,row)=>sum+row.rssBytes,0),processes}
}

/** Enumeration and liveness probes are separate snapshots. A worker can be
 * born/exiting between them; retry briefly, but never accept missing live RSS. */
export async function readOwnedNativeTable(roots,{readTable,probe=pid=>process.kill(pid,0),pause=()=>new Promise(resolve=>setTimeout(resolve,25))}) {
  for(let attempt=0;attempt<3;attempt++) {
    const rows=await readTable();
    let missingLive=false;
    for(const pid of roots)if(!rows.some(row=>row.pid===pid)) {
      try{probe(pid);missingLive=true;}catch(error){if(error.code!=='ESRCH')throw error;}
    }
    if(!missingLive)return rows;
    if(attempt<2)await pause();
  }
  throw Error('Owned process memory is unavailable');
}

export async function readProcessTree(roots,{groups=[]}={}) {
  if(!roots.length&&!groups.length)return{rssBytes:0,processes:[]}
  if([...roots,...groups].some(pid=>!Number.isSafeInteger(pid)||pid<=0))throw Error('Invalid owned process ID')
  let rows
  if(process.platform==='darwin') {
    const {readNativeProcessTable}=await import('#docx-observer');
    rows=await readOwnedNativeTable(roots,{readTable:readNativeProcessTable});
  }else if(process.platform==='win32') {
    const {stdout}=await execute('powershell.exe',['-NoProfile','-NonInteractive','-Command','@(Get-CimInstance Win32_Process | Select-Object ProcessId,ParentProcessId,WorkingSetSize) | ConvertTo-Json -Compress'],{timeout:5000,maxBuffer:8*1024*1024,windowsHide:true})
    const result=JSON.parse(stdout)
    rows=(Array.isArray(result)?result:[result]).map(row=>({pid:Number(row.ProcessId),parent:Number(row.ParentProcessId),rssBytes:Number(row.WorkingSetSize)}))
  } else {
    const {stdout}=await execute('/bin/ps',['-axo','pid=,ppid=,pgid=,rss='],{timeout:5000,maxBuffer:8*1024*1024})
    rows=stdout.trim().split('\n').filter(Boolean).map(line=>{const match=line.match(/^\s*(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s*$/);if(!match)throw Error('Invalid process memory response');return{pid:+match[1],parent:+match[2],group:+match[3],rssBytes:+match[4]*1024}})
  }
  if(rows.some(row=>[row.pid,row.parent,row.rssBytes].some(n=>!Number.isSafeInteger(n)||n<0)))throw Error('Invalid process memory counters')
  return selectProcessTree(rows,roots,groups)
}

export class ProcessBudget {
  constructor({roots,groups=()=>[],maximumBytes,intervalMs=500,onExceeded,sample=readProcessTree}) {
    if(!Number.isSafeInteger(maximumBytes)||maximumBytes<=0||!Number.isSafeInteger(intervalMs)||intervalMs<=0)throw Error('Invalid process budget')
    Object.assign(this,{roots,groups,maximumBytes,intervalMs,onExceeded,sample})
    this.maximumObservedBytes=0;this.last={rssBytes:0,processes:[]};this.pending=null;this.timer=null
  }
  check() {
    if(this.pending)return this.pending
    this.pending=(async()=>{
      const result=await this.sample(this.roots(),{groups:this.groups()})
      this.last=result;this.maximumObservedBytes=Math.max(this.maximumObservedBytes,result.rssBytes)
      if(result.rssBytes>this.maximumBytes){const error=Error('DOCX process-tree memory budget exceeded');error.code='LIMIT_EXCEEDED';error.observation=result;throw error}
      return result
    })().finally(()=>{this.pending=null})
    return this.pending
  }
  start() {
    if(this.timer)return
    this.timer=setInterval(()=>{void this.check().catch(error=>this.onExceeded?.(error))},this.intervalMs);this.timer.unref()
  }
  stop(){clearInterval(this.timer);this.timer=null}
}
