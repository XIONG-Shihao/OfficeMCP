import { runFormatAdapter } from "./format-adapters.mjs";
import { OfficeError, requireThat, errorPayload } from "../infrastructure/errors.mjs";

async function run(task) {
  if (task.kind === "office_read" && task.args.docxDomain)
    requireThat(task.format === "docx", "UNSUPPORTED_FORMAT", "DOCX detail domains require a DOCX document.");
  const adapted = await runFormatAdapter(task);
  if (adapted.handled) return adapted.result;
  throw new OfficeError("UNSUPPORTED_FEATURE", "Unknown worker operation.");
}

process.once("message", async (task) => {
  try {
    process.send({ ok: true, result: await run(task) }, () => process.disconnect());
  } catch (error) {
    process.send({ ok: false, error: errorPayload(error) }, () => process.disconnect());
  }
});
process.once("disconnect", () => process.exit(0));
