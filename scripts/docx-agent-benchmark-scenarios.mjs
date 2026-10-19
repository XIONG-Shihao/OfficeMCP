export const DOCX_AGENT_BENCHMARK_SCENARIOS = Object.freeze([
  {
    id: "contract-clause",
    category: "guarded formatted replacement",
    instruction:
      "Change the payment deadline from 30 days to 45 days. Preserve the bold formatting of the number and leave all other content intact. Read and report every legal qualification, including any outside the body.",
    expectedSignals: [
      "outputExists",
      "deadlineUpdated",
      "otherTextPreserved",
      "tablePreserved",
      "footerPreserved",
      "boldNumberPreserved",
      "footerReported",
    ],
  },
  {
    id: "proposal-template",
    category: "template filling",
    instruction:
      "Replace the literal [insert client name] placeholder with Meridian Labs, preserving the surrounding proposal text.",
    expectedSignals: [
      "outputExists",
      "placeholderFilled",
      "proposalContextPreserved",
    ],
  },
  {
    id: "report-table-caption",
    category: "table and reference maintenance",
    instruction:
      "Change the Service price in the existing table from 125 to 150 and insert the caption Service fees next to that table.",
    expectedSignals: ["outputExists", "tableUpdated", "unrelatedTableCellsPreserved", "captionAdded"],
  },
  {
    id: "header-footer-field",
    category: "non-body story editing",
    instruction:
      "Change the footer liability cap from USD 500 to USD 750 without changing the body.",
    expectedSignals: ["outputExists", "footerUpdated", "bodyPreserved"],
  },
  {
    id: "comment-read",
    category: "existing comment reading",
    instruction:
      "Read the existing comment on the Review clause paragraph, report its text and status, and save an unchanged copy. Preserve all comments and paragraph text.",
    expectedSignals: ["outputExists", "commentReported", "commentsPreserved", "commentAnchoredToReview", "reviewTextPreserved"],
  },
  {
    id: "image-layout",
    category: "image metadata and layout",
    instruction:
      "Set the existing image alternative text to Updated company mark and resize it to 24 by 24 points. Do not replace the image bytes.",
    expectedSignals: ["outputExists", "imageAltUpdated", "imageStillPresent", "imageDimensionsUpdated", "imageBytesPreserved"],
  },
  {
    id: "large-targeted-read",
    category: "bounded retrieval",
    instruction:
      "Find and report the exact paragraph containing Sentinel clause 287. Save an unchanged copy as reviewed.docx and do not rewrite unrelated paragraphs.",
    expectedSignals: ["outputExists", "sentinelReported", "largeBodyPreserved"],
  },
  {
    id: "ambiguous-target-refusal",
    category: "safe failure",
    instruction:
      "Try to rewrite exactly one occurrence of Duplicate risk without choosing between the two matching paragraphs. Do not save an output if the target is ambiguous; report the refusal.",
    expectedSignals: ["ambiguityRefused", "noOutputPublished"],
  },
]);

export function benchmarkScenario(id) {
  const scenario = DOCX_AGENT_BENCHMARK_SCENARIOS.find((item) => item.id === id);
  if (!scenario)
    throw new Error(
      `Unknown DOCX benchmark scenario ${id}. Choose: ${DOCX_AGENT_BENCHMARK_SCENARIOS.map((item) => item.id).join(", ")}.`,
    );
  return scenario;
}
