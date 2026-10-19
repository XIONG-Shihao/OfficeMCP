import * as fs from "node:fs";
import path from "node:path";
import { randomUUID } from "node:crypto";

const MAX_LOG_BYTES = 4 * 1024 * 1024;

function redact(text, config) {
  let value = String(text);
  const roots = [
    config.stateDirectory,
    ...config.readRoots,
    ...Object.values(config.writeRoots),
  ].sort((a, b) => b.length - a.length);
  for (const root of roots) value = value.split(root).join("<configured-root>");
  return value
    .replace(/\/(?:Users|home|private|tmp|var)\/[^\s:)]+/g, "<local-path>")
    .replace(/[A-Za-z]:\\(?:[^\s:)]+\\)*[^\s:)]+/g, "<local-path>")
    .slice(0, 32_000);
}

/** Write a private, bounded JSON-lines record and return its correlation ID. */
export function writeDiagnostic(config, source, error) {
  const id = `diag_${randomUUID()}`;
  const directory = path.join(config.stateDirectory, "diagnostics");
  const file = path.join(directory, "officemcp.jsonl");
  try {
    fs.mkdirSync(directory, { recursive: true, mode: 0o700 });
    if (fs.existsSync(file) && fs.statSync(file).size >= MAX_LOG_BYTES)
      fs.renameSync(file, `${file}.1`);
    const detail = (error instanceof Error ? error.stack ?? error.message : String(error))
      + (error?.details ? `\nDetails: ${JSON.stringify(error.details)}` : '');
    fs.appendFileSync(
      file,
      `${JSON.stringify({ id, at: new Date().toISOString(), source, detail: redact(detail, config) })}\n`,
      { mode: 0o600 },
    );
    fs.chmodSync(file, 0o600);
  } catch {
    // A diagnostic sink must never replace the original failure.
  }
  return id;
}

/** Raw SDK messages can contain paths or document-derived text. */
export function safeEngineFailure(error, diagnosticId) {
  const code = typeof error?.code === "string" && /^[A-Z0-9_]{1,40}$/.test(error.code)
    ? error.code
    : "SDK_ERROR";
  return {
    code: "ENGINE_FAILED",
    message: `SuperDoc failed (${code}).${diagnosticId ? ` Diagnostic ID: ${diagnosticId}.` : " Inspect the coordinator diagnostic record."}`,
    ...(diagnosticId ? { diagnosticId } : {}),
  };
}
