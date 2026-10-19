import { createHash } from "node:crypto";
import { requireThat } from "../infrastructure/errors.mjs";

const jsonBytes = value => Buffer.byteLength(JSON.stringify(value));
const identityFields = new Set([
  "sheet", "cell", "address", "range", "documentId", "candidateId",
  "revision", "candidateRevision", "sourceRevision", "operationId", "status",
]);

function omittedValue(value) {
  const text = typeof value === "string" ? value : JSON.stringify(value);
  let chars = 0;
  for (const _ of text) chars++;
  return {
    omitted: true,
    type: typeof value === "string" ? "string" : Array.isArray(value) ? "array" : "object",
    chars,
    bytes: Buffer.byteLength(text),
    sha256: createHash("sha256").update(text).digest("hex"),
    encoding: typeof value === "string" ? "utf-8" : "json-utf-8",
  };
}

/** Bound evidence, never workbook content or command guards. The durable result
 * must fit before acceptance, including public target/view enrichment and replay.
 * Kept locations remain usable by xlsx-service to construct revision-bound targets. */
export function boundWorkbookMutationEvidence(result, { maxResponseBytes, documentId }) {
  const available = Math.floor((maxResponseBytes - 1) / 2) - 1024;
  const originalTargets = result.data.affectedTargets ?? [];
  const total = result.data.postconditions?.affectedCellCount ?? originalTargets.length;
  const fieldBudget = Math.max(128, Math.floor(available / Math.max(8, originalTargets.length * 8)));

  function compact(value) {
    if (Array.isArray(value)) return value.map(compact);
    if (!value || typeof value !== "object") return value;
    const output = {};
    for (const [name, field] of Object.entries(value)) {
      if (typeof field === "string" && !identityFields.has(name) && jsonBytes(field) > fieldBudget) {
        (output.omittedFields ??= {})[name] = omittedValue(field);
      } else output[name] = compact(field);
    }
    return output;
  }

  const bounded = { ...result, data: compact(result.data) };
  const targets = bounded.data.affectedTargets ?? [];
  const view = { documentId, candidateId: bounded.data.candidateId, revision: bounded.data.candidateRevision };
  // Cover the largest current public shape: enriched targets plus a duplicated
  // legacy workflow summary. The 1024-byte reserve covers the common envelope.
  function responseBytes() {
    const enriched = bounded.data.affectedTargets.map(item => ({ ...item,
      target: { sheet: item.sheet, cell: item.cell, ...view } }));
    const expanded = { ...bounded, view, data: { ...bounded.data, affectedTargets: enriched,
      summary: { action: "rows_columns", targets: enriched, originalUnchanged: true },
      nextAction: "Read the returned view or call office_save." } };
    return Math.max(jsonBytes(expanded), jsonBytes({ operationId: result.operationId, data: bounded }));
  }

  function evidence() {
    const returned = bounded.data.affectedTargets.length;
    const countOmitted = value => {
      if (!value || typeof value !== "object") return 0;
      return Object.entries(value).reduce((count, [name, field]) => count +
        (name === "omittedFields" ? Object.keys(field).length : countOmitted(field)), 0);
    };
    const omittedFields = countOmitted(bounded.data);
    if (omittedFields || returned < total) {
      bounded.data.evidence = {
        bounded: true,
        omittedFieldCount: omittedFields,
        affectedTargets: { total, returned, omitted: total - returned },
        omittedValues: "omittedFields records Unicode characters, UTF-8 bytes and SHA-256; saved content and command guards remain exact.",
      };
      if (bounded.data.postconditions) bounded.data.postconditions.affectedCellsTruncated = returned < total;
    }
  }
  bounded.data.affectedTargets = targets;
  evidence();
  while (responseBytes() > available && bounded.data.affectedTargets.length) {
    bounded.data.affectedTargets.pop();
    evidence();
  }
  // Optional engine readback can contain another large collection. Retain an
  // explicit digest instead of allowing it to prevent an otherwise valid edit.
  for (const name of ["readback", "workflowVerification", "workflow", "semanticDelta", "receipt", "engineEvidence"]) {
    if (responseBytes() <= available) break;
    if (bounded.data[name] === undefined) continue;
    (bounded.data.omittedFields ??= {})[name] = omittedValue(bounded.data[name]);
    delete bounded.data[name];
    evidence();
  }
  requireThat(responseBytes() <= available, "LIMIT_EXCEEDED",
    "Response budget cannot hold the required mutation receipt; the candidate was not accepted.");
  return bounded;
}
