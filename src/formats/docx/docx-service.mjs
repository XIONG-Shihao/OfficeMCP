import { createWorkflow, saveWorkflow, asView } from "../common/document-workflow.mjs";
import { OfficeError, requireThat } from "../../infrastructure/errors.mjs";

export function viewOf(service, view) {
  const current = service.view({
    documentId: view.documentId,
    candidateId: view.candidateId,
    expectedRevision: view.revision,
  });
  requireThat(
    current.d.format === "docx",
    "UNSUPPORTED_FORMAT",
    "This contract only accepts DOCX documents.",
  );
  return current;
}

export function withTargets(service, view, result) {
  return service.documents.targets(view, result);
}

async function openDocx(service, source) {
  const result = service.documents.open({ path: source }, "docx");
  return { ...result, view: asView(result.document.id, result.document.sourceRevision) };
}

function closeDocx(service, view, disposition) {
  return service.documents.close(view.documentId, disposition, { view, format: "docx" });
}

function createFromTemplate(service, args, signal) {
  return service.documents.fromTemplate(args, "docx", signal);
}

async function saveDocx(service, args, signal) {
  viewOf(service, args.view);
  return saveWorkflow(service, args, signal, "docx_save");
}

export async function executeDocx(service, name, args, signal) {
  if (name === "docx_session") {
    if (args.action === "open") {
      requireThat(args.path && !args.view, "INVALID_ARGUMENT", "Open requires a path only.");
      return openDocx(service, args.path, signal);
    }
    requireThat(
      args.action === "close" && args.view && !args.path,
      "INVALID_ARGUMENT",
      "Close requires a view.",
    );
    return closeDocx(service, args.view, args.disposition, signal);
  }
  if (name === "docx_create") {
    if (args.action === "from_template") {
      requireThat(args.path && !args.document, "INVALID_ARGUMENT", "Template creation requires path only.");
      return createFromTemplate(service, args, signal);
    }
    requireThat(args.document && !args.path, "INVALID_ARGUMENT", "Content creation requires a DOCX specification only.");
    return createWorkflow(
      service,
      { operationId: args.operationId, document: args.document },
      signal,
    );
  }
  if (name === "docx_save") return saveDocx(service, args, signal);
  throw new OfficeError("UNKNOWN_TOOL", "Unknown internal DOCX lifecycle command.");
}
