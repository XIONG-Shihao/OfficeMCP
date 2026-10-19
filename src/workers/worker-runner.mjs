import * as fs from "node:fs";
import path from "node:path";
import { fork } from "node:child_process";
import { docxWorker } from "../formats/docx/docx-bridge.mjs";
import { DOCX_NATIVE_COMMANDS } from "../formats/docx/docx-host-process.mjs";
import { OfficeError, requireThat } from "../infrastructure/errors.mjs";
import {workerEnvironment,workerExecArgv} from '../infrastructure/worker-environment.mjs';
import {writeDiagnostic} from '../infrastructure/diagnostics.mjs';

export async function runWorker(
  service,
  kind,
  input,
  args = {},
  other = {},
  signal,
) {
  if (DOCX_NATIVE_COMMANDS.includes(kind))
    return service.docxPool.run(kind, input, args, other, signal);
  requireThat(!kind.startsWith("docx_native_"), "UNKNOWN_TOOL", "Unknown DOCX worker command.");
  if (
    !other.bypassDocx &&
    (kind.startsWith("docx") ||
      (other.format === "docx" ||
        args.document?.format === "docx" ||
        input?.toLowerCase().endsWith(".docx")))
  )
    return docxWorker(service, kind, input, args, other, signal);
  requireThat(
    service.activeWorkers.size < 2,
    "BUSY",
    "Worker limit reached; retry after the current job finishes.",
  );
  requireThat(!signal?.aborted, "CANCELLED", "Job cancelled.");
  return await new Promise((resolve, reject) => {
    const scratch = fs.mkdtempSync(
      path.join(service.config.stateDirectory, "worker-"),
    );
    let child;
    try {
      child = fork(new URL("./worker.mjs", import.meta.url), [], {
        stdio: ["ignore", "ignore", "pipe", "ipc"],
        // Spreadsheet size must not impose an artificial V8 memory ceiling.
        // Keep cancellation/deadlines and bounded responses; V8 uses its host defaults.
        // Spreadsheet size must not impose an artificial heap ceiling.
        execArgv: workerExecArgv(kind.startsWith("xlsx_") || kind.startsWith("csv_") || ["xlsx", "csv"].includes(other.format ?? args.document?.format)
          ? {} : {heapLimitMb: 768}),
        detached: process.platform !== "win32",
        env: workerEnvironment(),
      });
    } catch {
      fs.rmSync(scratch, { recursive: true, force: true });
      reject(new OfficeError("WORKER_FAILED", "Worker could not start."));
      return;
    }
    service.activeWorkers.add(child);
    let stderr='';child.stderr?.on('data',chunk=>{stderr=(stderr+String(chunk)).slice(-8192)});
    let done = false;
    const finish = (error, result) => {
      if (done) return;
      done = true;
      clearTimeout(timer);
      signal?.removeEventListener("abort", cancel);
      service.activeWorkers.delete(child);
      service.killWorker(child);
      error ? reject(error) : resolve(result);
    };
    const cancel = () =>
      finish(new OfficeError("CANCELLED", "Job cancelled before publication."));
    const timer = setTimeout(
      () =>
        finish(
          new OfficeError(
            "WORKER_TIMEOUT",
            "Worker exceeded its deadline. Source remains unchanged.",
          ),
        ),
      service.config.workerTimeoutMs,
    );
    signal?.addEventListener("abort", cancel, { once: true });
    child.once("error", () =>
      finish(new OfficeError("WORKER_FAILED", "Worker could not start.")),
    );
    child.once('exit',(code,signal)=>{
      if(done)return;
      const diagnosticId=writeDiagnostic(service.config,kind,Error(`Worker exit ${signal??code}\n${stderr}`));
      finish(new OfficeError('WORKER_FAILED',`Worker exited without a result. Diagnostic ID: ${diagnosticId}.`));
    });
    // Failed spawns emit close without exit; both paths must release scratch.
    child.once("close", () =>
      fs.rmSync(scratch, { recursive: true, force: true }),
    );
    child.once("message", (m) =>
      m.ok
        ? finish(null, m.result)
        : finish(new OfficeError(m.error.code, m.error.message)),
    );
    const sendFailed = () => finish(new OfficeError("WORKER_FAILED", "Worker could not receive the operation."));
    try {
      child.send({
        kind,
        input,
        args,
        maxExpandedBytes: service.config.maxExpandedBytes,
        maxFileBytes: service.config.maxFileBytes,
        engines: service.config.engines ?? {},
        timeoutMs: service.config.workerTimeoutMs,
        scratch,
        ...other,
      }, (error) => { if (error) sendFailed(); });
    } catch { sendFailed(); }
  });
}
