import { requireThat } from "../../infrastructure/errors.mjs";
import { success } from "./docx-sdk.mjs";

/** Both public routes decide the same supported tracked-change kinds by SDK ID. */
export async function decideTrackedChange(doc, locator, decision) {
  requireThat(locator?.kind === "change", "INVALID_ARGUMENT", "Use one tracked-change target.");
  requireThat(["insertion", "deletion", "replacement", "insert", "delete"].includes(
    locator.change?.changeType ?? locator.change?.type,
  ), "UNSUPPORTED_FEATURE", "Only explicitly selected text insertions, deletions and replacements are supported.");
  requireThat(["accept", "reject"].includes(decision), "INVALID_ARGUMENT", "Use accept or reject.");
  return success(await doc.trackChanges.decide({ target: { kind: "id", id: locator.nodeId }, decision }));
}
