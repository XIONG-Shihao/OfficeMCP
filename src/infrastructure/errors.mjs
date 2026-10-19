export class OfficeError extends Error {
  constructor(code, message) {
    super(message);
    this.code = code;
  }
}

export function requireThat(value, code, message) {
  if (!value) throw new OfficeError(code, message);
}

export function errorPayload(error) {
  return {
    code: error instanceof OfficeError ? error.code : "INTERNAL_ERROR",
    message:
      error instanceof OfficeError
        ? error.message
        : "Operation failed; inspect local diagnostics without exposing document contents.",
    recovery:
      {
        TARGET_CONFLICT:
          "Call office_read on the current view, resolve the target and its exact value again, then submit a new edit operationId.",
        REVISION_CONFLICT:
          "The view is stale. Read the current document/candidate and use the newly returned view.",
        CANDIDATE_REVISION_CONFLICT:
          "Read the candidate again; do not reuse an older revision.",
        SOURCE_REVISION_CONFLICT:
          "The original file changed outside OfficeMCP. Reopen it and resolve the intended edit before saving.",
        OPERATION_ID_CONFLICT:
          "This ID belongs to different arguments. Inspect office_operation_status; do not repeat uncertain work under a new ID.",
        OPERATION_PENDING:
          "Poll office_operation_status using the same operationId.",
        PARTIAL_MUTATION:
          "Inspect office_operation_status using the same operationId before attempting another write.",
        ENGINE_LICENSE_REQUIRED:
          "Configure a valid local Aspose license. Open-source creation, inspection and supported OOXML edits do not require it.",
        ENGINE_UNAVAILABLE:
          "Use office_engine_status and install/configure the missing optional runtime; no candidate was accepted.",
        DESTINATION_CONFLICT:
          "Choose a different output filename. OfficeMCP never overwrites an existing destination.",
        HOST_PUBLICATION_REFUSED:
          "Save or discard manual edits and check the current host run before saving again with a new operation ID. No destination was published.",
        LIMIT_EXCEEDED:
          "Reduce the requested window or document size, or have the operator increase the relevant configured limit.",
        PROFILE_CONFLICT:
          "Save this candidate and reopen the output before switching between advanced preservation profiles.",
        UNSUPPORTED_FEATURE:
          "The feature is outside this operation's scope. Read coverage and choose a supported operation; missing data is not evidence of absent content.",
        INVALID_ARGUMENT:
          "Check the input schema and argumentIssues (when present), correct those fields, then retry.",
      }[error?.code] ??
      "Inspect the error code and operation status before retrying; the source is not modified in place.",
    ...(error instanceof OfficeError && error.argumentIssues
      ? { argumentIssues: error.argumentIssues }
      : {}),
  };
}
