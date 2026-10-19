/** Accept a safe preflight refusal as well as an actual server refusal. */
export function ambiguityEvidence(steps, reportedText) {
  const serverRefusalObserved = steps.some(s => s.payload?.error?.code === 'TARGET_CONFLICT');
  const queried = steps.some(s => s.tool === 'docx_query' &&
    (s.payload?.data?.matchedTotal ?? s.payload?.data?.total) === 2 && s.arguments?.select?.pattern === 'Duplicate risk');
  const readIds = new Set(steps.filter(s => s.tool === 'docx_read' && !s.isError).flatMap(s =>
    (s.payload?.data?.items ?? []).filter(i => i.text === 'Duplicate risk' && i.itemId && ['paragraph', 'heading', 'listItem'].includes(i.kind))
      .map(i => JSON.stringify([i.story, i.itemId]))));
  const agentRecognizedAmbiguity = (queried || readIds.size >= 2) && /ambig|two|multiple|2 matching/i.test(reportedText);
  const acceptedEdit = steps.some(s => s.tool === 'docx_apply' && !s.isError);
  return { passed: !acceptedEdit && (serverRefusalObserved || agentRecognizedAmbiguity),
    serverRefusalObserved, agentRecognizedAmbiguity, distinctReadMatches: readIds.size, acceptedEdit };
}
