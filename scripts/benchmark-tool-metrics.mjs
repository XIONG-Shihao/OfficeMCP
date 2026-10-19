// Repeated names, identical requests and unnecessary work are different measures.
function canonical(value) {
  if (Array.isArray(value)) return value.map(canonical);
  if (value && typeof value === 'object')
    return Object.fromEntries(Object.keys(value).sort().map(key => [key, canonical(value[key])]));
  return value;
}
export function toolMetrics(events) {
  const counts = {}, requests = new Set();
  let repeatedRequests = 0, requestsWithArguments = 0;
  let toolInputJsonBytes = 0, toolOutputTextBytes = 0;
  let toolOutputsWithText = 0, toolOutputsMissing = 0, toolOutputsNonText = 0, toolEventsWithoutNames = 0;
  for (const event of events) {
    const name = event.part?.tool;
    if (!name) { toolEventsWithoutNames++; continue; }
    counts[name] = (counts[name] ?? 0) + 1;
    const output = event.part?.state?.output;
    if (typeof output === 'string') {
      toolOutputTextBytes += Buffer.byteLength(output, 'utf8');
      toolOutputsWithText++;
    } else if (output === undefined) toolOutputsMissing++;
    else toolOutputsNonText++;
    const input = event.part?.state?.input;
    if (input === undefined) continue;
    requestsWithArguments++;
    toolInputJsonBytes += Buffer.byteLength(JSON.stringify(input), 'utf8');
    const key = JSON.stringify([name, canonical(input)]);
    if (requests.has(key)) repeatedRequests++;
    requests.add(key);
  }
  const repeatedToolNames = Object.values(counts).reduce((sum, count) => sum + count - 1, 0);
  return {
    metricVersion: 2,
    toolCalls: events.length,
    toolSequence: events.map(event => event.part?.tool).filter(Boolean),
    toolNameCounts: counts,
    repeatedToolNames,
    repeatedToolCalls: repeatedToolNames, // Retained legacy field; never a waste estimate.
    repeatedRequests,
    requestsWithArguments,
    toolPayloadMetricVersion: 1,
    toolInputJsonBytes,
    toolOutputTextBytes,
    toolOutputsWithText,
    toolOutputsMissing,
    toolOutputsNonText,
    toolEventsWithoutNames,
    toolInputJsonComplete: requestsWithArguments === events.length,
    toolOutputTextComplete: toolOutputsWithText === events.length,
    payloadMeaning: 'UTF-8 bytes of JSON.stringify(state.input) and captured string state.output. These are text bytes, not tokens, transport bytes or native image/binary content. Missing and non-text outputs are counted separately and contribute no measured output bytes.',
    repetitionMeaning: 'repeatedToolCalls is the legacy alias for repeatedToolNames. repeatedRequests compares tool name and exact arguments with sorted object keys, including view/revision and pagination; retries may be intentional. Neither counts unnecessary work.',
  };
}
