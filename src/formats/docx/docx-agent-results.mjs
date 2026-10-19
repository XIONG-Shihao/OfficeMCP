const S = { type: "string" };
const N = { type: "number" };
const I = { type: "integer" };
const B = { type: "boolean" };
const O = { type: "object" };
const A = (items = {}) => ({ type: "array", items });
const nullable = (schema) => ({ anyOf: [schema, { type: "null" }] });

const view = {
  type: "object",
  required: ["documentId", "revision"],
  properties: { documentId: S, candidateId: S, revision: S },
};
const target = {
  type: "object",
  required: ["id"],
  properties: { id: S },
  additionalProperties: false,
};
const provenance = {
  type: "object",
  properties: { engine: S, source: S, revision: S, editable: B },
};
const item = {
  type: "object",
  required: ["kind", "itemId", "provenance"],
  properties: {
    kind: S,
    itemId: S,
    text: S,
    story: { anyOf: [S, O] },
    order: I,
    path: A({}),
    target,
    provenance,
    rows: A(A(O)),
    children: A({}),
    details: O,
    formattingSummary: O,
    reviewChanges: A(O),
    altText: S,
    packagePart: S,
    status: S,
    changeType: S,
    author: S,
    name: S,
    count: I,
    partName: S,
    textLength: I,
    truncated: B,
    contentSource: S,
    mutationTargetAvailable: B,
  },
};
const coverage = {
  type: "object",
  required: ["completeDocument", "windowComplete", "matrix", "remainingReads"],
  properties: {
    completeDocument: B,
    windowComplete: B,
    dimensions: O,
    matrix: O,
    scopes: A(O),
    remainingReads: A(O),
    notChecked: A(S),
  },
};
const profiles = {
  read: {
    items: A(item),
    total: I,
    offset: I,
    nextOffset: nullable(I),
    remainingReads: A(O),
    readSnapshotId: S,
    complete: B,
    limitations: A(S),
    nextAction: S,
  },
  inspect: {
    counts: O,
    outline: A(O),
    detectedDomains: A(S),
    readingPlan: A(O),
    diagnostics: A(O),
    availableActions: A(S),
    actionFamilies: A({ type: "object", required: ["id", "description", "actionCount"], properties: { id: S, description: S, actionCount: I }, additionalProperties: false }),
    selectedActionFamily: S,
    actionDescriptors: { type: "object", additionalProperties: O },
    accessibilityReport: O,
    storyInventory: A(O),
    objectInventory: A(O),
    textboxInventory: A(O),
    templateReport: O,
    documentInspected: B,
    availableWorkflows: A(S),
    workflowDescriptors: { type: "object", additionalProperties: O },
    usage: O,
  },
  apply: {
    candidateId: S,
    candidateRevision: S,
    applied: B,
    preview: B,
    wouldChange: B,
    action: S,
    steps: A(O),
    affectedTargets: A(S),
    verificationPassed: B,
    validationPassed: B,
    published: B,
    engine: S,
    nextAction: S,
    atomicity: S,
    workflow: S,
    postconditions: O,
    semanticDelta: O,
    workflowVerification: O,
    readback: { type: "object", properties: {
      scope: { enum: ["body", "objects"] }, total: I, truncated: B, textChars: I, reusableTargets: B,
      notChecked: A(S), nextAction: S,
      items: A({ type: "object", properties: {
        kind: S, change: { enum: ["modified", "inserted", "deleted"] },
        before: { type: "object", properties: { text: S, truncated: B } },
        after: { type: "object", properties: { text: S, truncated: B } },
        row: I, column: I, target, targetKind: { enum: ["block", "table", "object"] },
      } }),
    } },
  },
  validation: {
    passed: B,
    checks: A(S),
    notVerified: A(S),
    packageSummary: O,
  },
  compare: {
    hasChanges: B,
    changedComponents: A(S),
    summary: O,
    coverage: O,
    baseFingerprint: S,
    targetFingerprint: S,
  },
  export: {
    published: B,
    format: S,
    path: S,
    assetId: S,
    outputRevision: S,
    engine: S,
    notVerified: A(S),
  },
  media: {
    assetId: S,
    width: I,
    height: I,
    engine: S,
    source: S,
    size: I,
    notVerified: A(S),
  },
};

function resultEnvelope(fields) {
  return {
    $schema: "https://json-schema.org/draft/2020-12/schema",
    type: "object",
    required: ["contractVersion", "requestId", "status", "warnings"],
    properties: {
      contractVersion: { const: "docx/6" },
      requestId: S,
      status: { enum: ["completed", "failed"] },
      warnings: A(S),
      error: O,
      operationId: S,
      view,
      coverage,
      data: { type: "object", properties: fields },
    },
  };
}

export const docxAgentResultSchemas = Object.freeze(Object.fromEntries(
  Object.entries({ generic: {}, ...profiles }).map(([name, fields]) => [name, resultEnvelope(fields)]),
));
