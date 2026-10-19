const S = { type: "string" };
const N = { type: "number" };
const I = { type: "integer" };
const B = { type: "boolean" };
const O = { type: "object" };
const A = (items = {}) => ({ type: "array", items });
const nullable = (schema) => ({ anyOf: [schema, { type: "null" }] });
const described = (schema, description) => ({ ...schema, description });

const view = {
  type: "object",
  required: ["documentId", "revision"],
  properties: {
    documentId: S,
    candidateId: S,
    revision: { ...S, pattern: "^sha256:[a-f0-9]{64}$" },
  },
};
const document = {
  type: "object",
  properties: {
    id: S,
    format: S,
    sourceRevision: S,
    candidateId: S,
    revision: S,
  },
};
const item = {
  type: "object",
  properties: {
    kind: S,
    text: nullable(S),
    target: O,
    story: { anyOf: [S, O] },
    sheet: S,
    address: { anyOf: [S, O] },
    value: {},
    exactValue: {},
    formula: { anyOf: [S, O, { type: "null" }] },
    cachedValue: nullable(O),
    displayedText: nullable(S),
    expectedFormula: nullable(S),
    shapeId: nullable(S),
    slide: I,
    page: I,
    rows: A(A(O)),
    runs: A(O),
    reviewChanges: A(O),
  },
};
const coverage = described(
  {
    type: "object",
    properties: {
      completeDocument: B,
      completeForScope: B,
      completeForDocument: B,
      windowComplete: B,
      scope: O,
      features: A(O),
      omissions: A(S),
      notChecked: A(S),
    },
  },
  "Visited scope, completeness and explicit omissions.",
);
const mutation = {
  candidateId: described(S, "Private staged candidate; the source is unchanged."),
  candidateRevision: described(S, "Revision required by the next guarded call."),
  sourceRevision: S,
  applied: B,
  validationPassed: B,
  published: B,
  affectedTargets: A({}),
  engine: S,
  evaluation: B,
  engineEvidence: O,
  action: S,
  receipt: O,
  workflow: S,
  postconditions: O,
  semanticDelta: O,
};
const page = {
  items: described(A(item), "Only the requested result window."),
  total: described(N, "Total items in the selected scope."),
  offset: I,
  nextOffset: nullable(I),
  domain: S,
  omissions: A(S),
};
const profiles = {
  create: {
    candidateId: S,
    candidateRevision: S,
    created: B,
    published: B,
    warnings: A({}),
    nextAction: S,
  },
  save: {
    candidateId: S,
    path: S,
    artifactDescriptor: S,
    assetId: S,
    published: B,
    outputRevision: S,
    validationPassed: B,
  },
  asset: {
    assetId: S,
    path: S,
    mediaType: S,
    size: N,
    hash: S,
    encoding: S,
    bytes: S,
    offset: I,
    totalBytes: I,
    nextOffset: nullable(I),
  },
  validation: {
    validationId: S,
    passed: B,
    revision: S,
    profile: S,
    checks: A(S),
    notVerified: A(S),
    phases: A(O),
  },
};
function resultEnvelope(fields) {
  return {
    $schema: "https://json-schema.org/draft/2020-12/schema",
    type: "object",
    required: ["contractVersion", "requestId", "status", "warnings"],
    properties: {
      contractVersion: S,
      requestId: S,
      status: S,
      warnings: A({}),
      error: O,
      operationId: S,
      jobId: S,
      pollAfterMs: N,
      document,
      view,
      data: { type: "object", properties: fields },
      coverage,
      nextCursor: nullable(S),
      validation: O,
    },
  };
}

export const resultSchemas = Object.freeze(Object.fromEntries(
  Object.entries({ generic: {}, mutation, page, ...profiles }).map(([name, fields]) => [name, resultEnvelope(fields)]),
));
