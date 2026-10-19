import { executeDocxAgent } from "../formats/docx/docx-agent-service.mjs";
import { queryWorkbook, applyWorkbook } from "../formats/xlsx/xlsx-service.mjs";
import {
  createWorkflow,
  editWorkflow,
  saveWorkflow,
  asView,
} from "../formats/common/document-workflow.mjs";
import { operationDescriptor } from "../contracts/registry.mjs";
import { OfficeError, requireThat } from "../infrastructure/errors.mjs";
import {
  digest,
  fingerprint,
  newId,
  regularFile,
} from "../infrastructure/files.mjs";
import { serverInfo } from "../protocol/server-info.mjs";

const commands = Object.freeze({
  queryWorkbook: (service, _name, args, signal) => queryWorkbook(service, args, signal),
  applyWorkbook: (service, _name, args, signal) => applyWorkbook(service, args, signal),
  docxCommand: (service, name, args, signal) => executeDocxAgent(service, name, args, signal),
  createWorkflow: (service, _name, args, signal) => createWorkflow(service, args, signal),
  editWorkflow: (service, _name, args, signal) => editWorkflow(service, args, signal),
  saveWorkflow: (service, _name, args, signal) => saveWorkflow(service, args, signal),
  readWorkflow: readWorkflow,
  readDocxDetails: readDocxDetails,
  extractAsset: extractAsset,
  readImage: readImage,
  assetInfo: (service, _name, args) => readAsset(service, args, true),
  readAsset: (service, _name, args) => readAsset(service, args, false),
  extractPackage: extractPackage,
  openDocument: openDocument,
  closeDocument: closeDocument,
  queryDocument: queryDocument,
  discardCandidate: discardCandidate,
  validateCandidate: validateCandidate,
  mutateCandidate: (service, name, args, signal) => service.operation(name, args, signal),
  operationStatus: operationStatus,
  jobStatus: jobStatus,
  cancelJob: cancelJob,
  engineStatus: async (service, name, _args, signal) => ({ data: await service.worker(name, null, {}, {}, signal) }),
  stateInfo: (service) => ({ data: service.stateInfo() }),
  cleanup: (service) => ({ data: service.cleanup() }),
  serverInfo: (service) => serverInfo(service),
});

export function bindCommands(service) {
  return Object.freeze(Object.fromEntries(Object.entries(commands).map(([name, execute]) =>
    [name, (tool, args, signal) => execute(service, tool, args, signal)])));
}

export function operationBinding(preset, name) {
  const operation = operationDescriptor(preset, name);
  if (!operation) return null;
  const execute = operation.execute;
  requireThat(execute, "UNKNOWN_TOOL", "Operation has no application binding.");
  return { operation, execute };
}

export async function executeTool(service, name, args, signal) {
  // Public parsing is preset-bound. Legacy lifecycle adapters also use internal
  // complete-surface operations, whose bindings remain available here.
  const binding = operationBinding(service.config.toolset, name) ?? operationBinding("all", name);
  requireThat(binding, "UNKNOWN_TOOL", "Unknown tool.");
  return binding.execute(service, name, args, signal);
}

async function extractAsset(service, name, a, signal) {
  const v = service.view({ documentId: a.view.documentId, candidateId: a.view.candidateId, expectedRevision: a.view.revision });
  return { view: a.view, ...await service.artifacts.produce({
    kind: name, input: v.file, args: a, context: { format: v.d.format },
    documentId: a.view.documentId, prefix: "extracted", reservation: 20 * 1024 * 1024, includeImage: "auto",
  }, signal) };
}

async function readImage(service, _name, a) {
  const { asset, bytes, nativeImage } = service.artifacts.image(a.assetId);
  return { data: { assetId: a.assetId, mediaType: asset.mediaType, size: bytes.length }, nativeImage };
}

async function readWorkflow(service, name, a, signal) {
  requireThat(
    Boolean(a.documentId) !== Boolean(a.view) && (!a.view || !a.candidateId),
    "INVALID_ARGUMENT",
    "Pass either documentId (and optional candidateId) or the returned view object.",
  );
  requireThat(
    a.offset === 0 || a.view,
    "INVALID_ARGUMENT",
    "Pagination requires the returned view object so pages cannot silently mix revisions.",
  );
  const documentId = a.view?.documentId ?? a.documentId,
    candidateId = a.view?.candidateId ?? a.candidateId;
  const d = service.doc(documentId),
    c = candidateId ? service.get("candidate", candidateId) : null;
  const v = service.view({
    documentId,
    candidateId,
    expectedRevision: a.view?.revision ?? c?.revision ?? d.revision,
  });
  const result = await service.worker(
    "office_read",
    v.file,
    a,
    { format: d.format },
    signal,
  );
  return {
    view: asView(documentId, v.revision, candidateId),
    data: result,
    coverage: result.coverage,
  };
}

async function openDocument(service, _name, args) {
  return service.documents.open(args);
}

async function operationStatus(service, name, a, signal) {
  const op = service.get("operation", a.operationId);
  if (op.status === "expired")
    throw new OfficeError(
      "OPERATION_EXPIRED",
      "Receipt expired; service operation ID cannot be reused.",
    );
  if (op.status === "failed")
    throw new OfficeError(op.error.code, op.error.message);
  return {
    operationId: a.operationId,
    data: op.result ?? {
      state: op.status,
      recovery:
        "Reconcile the recorded operation; never use a new ID to repeat uncertain work.",
    },
  };
}

async function jobStatus(service, name, a, signal) {
  const j = service.get("job", a.jobId);
  if (j.status === "failed" || j.status === "cancelled")
    throw new OfficeError(
      j.result?.error?.code ?? "WORKER_FAILED",
      j.result?.error?.message ?? "Job interrupted.",
    );
  return {
    status: j.status,
    jobId: a.jobId,
    data: j.result ?? null,
    pollAfterMs: 500,
  };
}

async function cancelJob(service, name, a, signal) {
  const j = service.get("job", a.jobId);
  service.jobs.get(a.jobId)?.abort();
  return {
    jobId: a.jobId,
    data: {
      state: service.jobs.has(a.jobId) ? "cancelling" : j.status,
      result: j.result ?? null,
    },
  };
}

async function closeDocument(service, _name, args) {
  return service.documents.close(args.documentId, args.disposition);
}

async function discardCandidate(service, name, a, signal) {
  service.doc(a.documentId);
  return { data: service.discard(a.candidateId, a.documentId) };
}

async function readAsset(service, a, metadataOnly) {
  const asset = service.get("asset", a.assetId);
  service.doc(asset.documentId);
  const bytes = regularFile(asset.path, service.config.maxFileBytes);
  requireThat(
    digest(bytes) === asset.hash,
    "INTEGRITY_FAILED",
    "Artifact changed.",
  );
  if (metadataOnly)
    return { data: { ...asset, assetId: a.assetId } };
  requireThat(
    a.offset <= bytes.length,
    "INVALID_ARGUMENT",
    "Offset exceeds asset length.",
  );
  const slice = bytes.subarray(a.offset, a.offset + a.maxBytes);
  return {
    data: {
      assetId: a.assetId,
      encoding: "base64",
      offset: a.offset,
      totalBytes: bytes.length,
      bytes: slice.toString("base64"),
      hash: asset.hash,
      nextOffset:
        a.offset + slice.length < bytes.length
          ? a.offset + slice.length
          : null,
    },
  };
}

async function validateCandidate(service, name, a, signal) {
  const validation = await service.documents.validate(a, signal);
  return { data: { validationId: validation.id, ...validation.result } };
}

async function readDocxDetails(service, _name, a, signal) {
  return service.execute(
    "office_read",
    {
      view: a.view,
      docxDomain: a.domain,
      offset: a.offset,
      limit: a.limit,
    },
    signal,
  );
}

function formatView(service, name, a) {
  const v = service.view(a);
  requireThat(
    name.startsWith(v.d.format + "_") ||
      (name.startsWith("office_package_") && ["docx", "xlsx", "pptx"].includes(v.d.format)),
    "UNSUPPORTED_FORMAT",
    "Tool does not match document format.",
  );
  service.checkQuota();
  return v;
}

async function extractPackage(service, name, a, signal) {
  const v = formatView(service, name, a);
  return { document: service.viewIdentity(a, v), ...await service.artifacts.produce({
    kind: name, input: v.file, args: a, documentId: a.documentId,
    prefix: "part", reservation: service.config.maxFileBytes,
  }, signal) };
}

async function queryDocument(service, name, a, signal) {
  const v = formatView(service, name, a);
  const result = await service.worker(name, v.file, a, {}, signal);
  if (result.items) {
    let index = 0;
    const query = fingerprint({
      name,
      revision: v.revision,
      args: { ...a, cursor: undefined, background: undefined },
    });
    if (a.cursor) {
      const cursor = service.get("cursor", a.cursor);
      requireThat(
        cursor.query === query,
        "CURSOR_CONFLICT",
        "Cursor belongs to a different query, window or revision.",
      );
      index = cursor.index;
    }
    const items = result.items.slice(index, index + a.limit);
    let nextCursor = null;
    if (index + items.length < result.items.length) {
      service.checkQuota();
      requireThat(
        service.rows("cursor").length < 5000,
        "LIMIT_EXCEEDED",
        "Cursor quota reached.",
      );
      nextCursor = newId("cursor");
      service.put("cursor", nextCursor, { query, index: index + items.length });
    }
    return {
      document: service.viewIdentity(a, v),
      data: { ...result, items, total: result.items.length },
      coverage: {
        scope: {
          domain: result.domain,
          page: a.page,
          pageFrom: a.pageFrom,
          pageTo: a.pageTo,
          sheetId: a.sheetId,
          range: a.range,
          slide: a.slide,
          collection: a.collection,
          start: index,
          count: items.length,
        },
        completeForScope: true,
        completeForDocument: false,
        omissions: result.omissions ?? [
          "non-main-body-content",
          "layout",
          "review-semantics",
        ],
      },
      nextCursor,
    };
  }
  return {
    document: service.viewIdentity(a, v),
    data: result,
    coverage: {
      completeForScope: true,
      completeForDocument: false,
      omissions: result.omissions,
    },
  };
}
