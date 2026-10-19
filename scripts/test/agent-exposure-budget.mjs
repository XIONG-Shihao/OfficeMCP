/** Caller-owned, format-independent boundary for model-visible MCP results. */
export function exposureBudget({ perReply = 32 * 1024, total = 256 * 1024 } = {}) {
  const state = { perReply, total, sentBytes: 0, refusedReplies: 0 };
  return { state, expose(output, result) {
    const bytes = Buffer.byteLength(output);
    const binary = result?.content?.some(c => c.type !== 'text' && c.type !== 'resource_link');
    if (binary || bytes > perReply || state.sentBytes + bytes > total) {
      state.refusedReplies++;
      output = JSON.stringify({ clientError: 'MODEL_EXPOSURE_BUDGET', reason: binary ? 'Binary results cannot be sent to the model.' : 'Tool result exceeds the per-reply or cumulative byte budget.',
        instruction: 'Request fewer items, narrower fields or a precise target. The original result was withheld, not truncated.', perReplyBytes: perReply, remainingBytes: total - state.sentBytes });
    }
    const sent = Buffer.byteLength(output);
    if (sent > perReply || state.sentBytes + sent > total) throw new Error('Model exposure budget exhausted; no further tool result can be sent');
    state.sentBytes += sent; return output;
  } };
}
