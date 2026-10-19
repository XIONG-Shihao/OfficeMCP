/** Opt-in numerical orchestration timings; never log document payloads. */
export function phaseEvent(name,start,extra={}) {
  if(process.env.RPF_DOCX_MCP_PHASE_TRACE==='1')process.stderr.write(`[docx-phase] ${JSON.stringify({name,at:Date.now(),ms:performance.now()-start,memory:process.memoryUsage(),...extra})}\n`)
}
export async function tracePhase(name,work) {
  if(process.env.RPF_DOCX_MCP_PHASE_TRACE!=='1')return work()
  const start=performance.now();try{return await work()}finally{phaseEvent(name,start)}
}
