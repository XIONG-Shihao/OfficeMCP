import { requireThat } from "../../infrastructure/errors.mjs";

const body = { kind: "story", storyType: "body" };

function occurrences(value, pattern) {
  if (!pattern) return 0;
  let total = 0;
  for (let offset = 0; ; ) {
    const at = value.indexOf(pattern, offset);
    if (at < 0) return total;
    total++;
    offset = at + Math.max(1, pattern.length);
  }
}

export async function capturePostconditionBaseline(doc, expect) {
  if (!expect) return null;
  const text = String(await doc.getText({ in: body }));
  for (const [pattern, expected] of Object.entries(
    expect?.beforeOccurrences ?? {},
  ))
    requireThat(
      occurrences(text, pattern) === expected,
      "TARGET_CONFLICT",
      `Expected ${expected} occurrence(s) of ${JSON.stringify(pattern)} before mutation.`,
    );
  return text;
}

export async function verifyPostconditions(doc, expect, beforeText) {
  if (!expect) return {};
  const afterText = String(await doc.getText({ in: body }));
  const checked = [];
  for (const [pattern, expected] of Object.entries(
    expect.remainingOccurrences ?? {},
  )) {
    requireThat(
      occurrences(afterText, pattern) === expected,
      "VALIDATION_FAILED",
      `Expected ${expected} occurrence(s) of ${JSON.stringify(pattern)} after mutation.`,
    );
    checked.push(`remainingOccurrences:${pattern}`);
  }
  const changedOccurrences = Object.keys(expect.beforeOccurrences ?? {}).reduce(
    (total, pattern) =>
      total + occurrences(beforeText, pattern) - occurrences(afterText, pattern),
    0,
  );
  if (expect.changedOccurrences !== undefined) {
    requireThat(
      changedOccurrences === expect.changedOccurrences,
      "VALIDATION_FAILED",
      `Expected ${expect.changedOccurrences} changed occurrence(s); observed ${changedOccurrences}.`,
    );
    checked.push("changedOccurrences");
  }
  return {
    postconditions: { passed: true, checked },
    semanticDelta: { changedOccurrences },
  };
}
