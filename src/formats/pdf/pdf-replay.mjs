import { readFileSync } from "node:fs";
import { digest } from "../../infrastructure/files.mjs";
import { requireThat } from "../../infrastructure/errors.mjs";

export async function verifyPdfReplay(task, candidate, apply, message) {
  let expected = readFileSync(task.original);
  for (const operation of task.operations) expected = await apply(expected, operation);
  requireThat(digest(expected) === digest(candidate), "VALIDATION_FAILED", message);
}
