import { OfficeError, errorPayload } from "../../infrastructure/errors.mjs";
import { safeEngineFailure, writeDiagnostic } from "../../infrastructure/diagnostics.mjs";

const commands = Object.freeze({
  docx_native_create: { prepare: "dispose", mutating:true, run: (handlers, task) => handlers.create(task) },
  docx_native_read: { run: (handlers, task) => handlers.read(task.args) },
  docx_native_edit: { mutating:true, run: (handlers, task) => handlers.edit(task) },
  docx_native_agent_apply: { mutating:true, run: (handlers, task) => handlers.agentApply(task) },
  docx_native_agent_preview: { mutating:true, run: (handlers, task) => handlers.agentApply(task, true) },
  docx_native_validate: { prepare:"prepareValidation", run: (handlers, task) => handlers.validate(task) },
  docx_native_agent_inspect: { run: (handlers, task) => handlers.agentInspect(task.args, task) },
  docx_native_agent_read: { run: (handlers, task) => handlers.agentRead(task.args) },
  docx_native_agent_query: { run: (handlers, task) => handlers.agentQuery(task.args) },
  docx_native_agent_compare: { run: (handlers, task) => handlers.agentCompare(task) },
});
export const DOCX_NATIVE_COMMANDS = Object.freeze(Object.keys(commands));

export async function executeDocxHostTask(handlers, task) {
  const command = Object.hasOwn(commands, task.kind) ? commands[task.kind] : null;
  if (!command) throw new OfficeError("UNKNOWN_TOOL", "Unknown DOCX worker command.");
  await handlers[command.prepare ?? "load"](task);
  if(command.mutating)handlers.invalidateReads?.();
  return command.run(handlers, task);
}

/** Keep process/IPC/error policy separate from the SuperDoc domain adapter. */
export function installDocxHostProcess(handlers) {
  let busy = false;
  process.on("message", async (task) => {
    if (busy)
      return process.send({
        ok: false,
        error: { code: "BUSY", message: "SDK host busy." },
      });
    busy = true;
    try {
      const result = await executeDocxHostTask(handlers, task);
      process.send({
        ok: true,
        result: { ...result, engineSession: { pid: process.pid, loads: handlers.loads(), ...(handlers.cacheStats?{cache:handlers.cacheStats()}: {}) } },
      });
    } catch (error) {
      await handlers.dispose();
      process.send({
        ok: false,
        error: error instanceof OfficeError
          ? errorPayload(error)
          : safeEngineFailure(
              error,
              task.diagnosticConfig
                ? writeDiagnostic(task.diagnosticConfig, "superdoc-host", error)
                : undefined,
            ),
      });
    } finally {
      busy = false;
    }
  });
  process.on("disconnect", () => process.exit(0));
}
