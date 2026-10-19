import { insertParagraphsAt } from "./docx-paragraph-actions.mjs";
import { mutateTableRows } from "./docx-table-actions.mjs";
import { success } from './docx-sdk.mjs';
import { requireThat } from "../../infrastructure/errors.mjs";
import {
  DOCX_ACTION_REGISTRY,
  matchesDocxActionTarget,
} from "./docx-agent-actions.mjs";
import { applyExtendedAction } from "./docx-extended-actions.mjs";

function coreArguments(step, locator) {
  const args = { ...step.arguments };
  const locators = step.targetLocators?.length
    ? step.targetLocators
    : locator
      ? [locator]
      : [];
  if (step.action === "delete_blocks")
    args.selectors = locators.map((item) => item.selector);
  else if (locator?.kind === "table") {
    requireThat(!locator.address?.story || locator.address.story.storyType === "body",
      "UNSUPPORTED_FEATURE", "This core table action uses body table ordinals. Use an address-based action such as write_table_cells for another story.");
    args.tableOrdinal ??= (locator.index ?? 0) + 1;
    requireThat(args.tableOrdinal === (locator.index ?? 0) + 1, "TARGET_CONFLICT", "tableOrdinal differs from the selected table target.");
  }
  else if (locator?.selector) args.selector = locator.selector;
  if (args.placement && ["before", "after"].includes(args.placement.at)) {
    requireThat(
      locator?.selector,
      "INVALID_ARGUMENT",
      `${step.action} placement ${args.placement.at} requires a block target.`,
    );
    args.placement = { at: args.placement.at, selector: locator.selector };
  }
  return args;
}

function verifyReceipt(action, receipt) {
  requireThat(
    receipt?.status === "ok" && receipt.verificationPassed !== false,
    receipt?.status === "partial" ? "PARTIAL_MUTATION" : "ENGINE_FAILED",
    receipt?.error?.message ??
      receipt?.errors?.map((error) => error.message).join("; ") ??
      `SuperDoc action ${action} did not return a verified success receipt.`,
  );
  return receipt;
}

/** Execute one declarative action descriptor from validation through receipt. */
export async function executeDocxAction({
  doc,
  step,
  locator,
  toolkit,
  placementOf,
  directTextReplacement,
}) {
  const descriptor = DOCX_ACTION_REGISTRY[step.action];
  const requestedAction = step.action;
  requireThat(descriptor, "INVALID_ARGUMENT", `Unknown DOCX action: ${step.action}.`);
  const locators = step.targetLocators?.length
    ? step.targetLocators
    : locator
      ? [locator]
      : [];
  const optional = descriptor.target.endsWith("-optional") || descriptor.target === "none";
  requireThat(optional || locators.length, "INVALID_ARGUMENT", `${step.action} requires a ${descriptor.target} target.`);
  requireThat(descriptor.cardinality === "many" || locators.length <= 1, "INVALID_ARGUMENT", `${step.action} accepts one target.`);
  for (const target of locators)
    requireThat(matchesDocxActionTarget(descriptor.target, target), "INVALID_ARGUMENT", `${step.action} requires target kind ${descriptor.target}.`);

  const appliedStep = { ...step, targetLocator: locator };
  let insertionReceipt;
  if (step.action === 'append_list' && step.arguments.placement && !step.arguments.headingText) {
    coreArguments(step, locator);
    const at = step.arguments.placement.at;
    const results = await insertParagraphsAt(doc, { target: locator?.address, placement: at === 'document_end' ? 'documentEnd' : at === 'document_start' ? 'documentStart' : at, texts: step.arguments.items, changeMode: step.arguments.changeMode, paragraphStyleId: 'Normal' });
    const addresses = results.map(result => result.paragraph);
    const result = success(await doc.lists.create({ mode: 'fromParagraphs', target: { from: addresses[0], to: addresses.at(-1) }, preset: step.arguments.kind === 'ordered' ? 'decimal' : 'disc', level: 0 }));
    insertionReceipt = { status: 'ok', intent: step.action, verificationPassed: true, selectedTargets: addresses.map(a => a.nodeId), executedOperations: [{ operationId: 'doc.lists.create', result }] };
  }
  if (locator?.kind === 'table' && step.action === 'delete_table') {
    if (step.arguments.tableOrdinal !== undefined) coreArguments(step, locator);
    const result = await doc.tables.delete({ target: locator.address, changeMode: step.arguments.changeMode });
    requireThat(result?.success !== false, 'ENGINE_FAILED', 'SDK table deletion failed.');
    insertionReceipt = { status: 'ok', intent: step.action, verificationPassed: true, selectedTargets: [locator.address.nodeId], executedOperations: [{ operationId: 'doc.tables.delete', result }] };
  }
  if (locator?.kind === "table" && ["insert_table_row", "delete_table_row"].includes(step.action) && Number.isInteger(step.arguments.rowIndex)) {
    coreArguments(step, locator); // Enforce matching body/ordinal guards before using the address.
    const result = await mutateTableRows(doc, locator, {
      ...step.arguments, action: step.action === "delete_table_row" ? "delete" : "insert", values: step.arguments.cellTexts,
    });
    insertionReceipt = { status: "ok", intent: step.action, verificationPassed: true,
      selectedTargets: [locator.address.nodeId], executedOperations: [{ operationId: step.action, result }] };
  }
  if (step.action === "insert_paragraphs" && (step.paragraphStyleId || locator?.address && ["before", "after"].includes(step.arguments.placement?.at))) {
    const at = step.arguments.placement.at;
    const results = await insertParagraphsAt(doc, {
      target: locator?.address, placement: at === 'document_end' ? 'documentEnd' : at === 'document_start' ? 'documentStart' : at,
      texts: step.arguments.texts ?? [step.arguments.text],
      headingLevel: step.arguments.headingLevel, changeMode: step.arguments.changeMode,
      paragraphStyleId: step.paragraphStyleId,
    });
    insertionReceipt = { status: "ok", intent: step.action, verificationPassed: true,
      selectedTargets: locator?.address ? [locator.address.nodeId] : results.map(result => result.paragraph?.nodeId ?? result.heading?.nodeId), executedOperations: results.map(result => ({ operationId: step.action, result })) };
  }
  const directReceipt = insertionReceipt ?? await directTextReplacement(step, locator);
  const receipt = directReceipt ??
    (descriptor.handler === "extended"
      ? await applyExtendedAction(doc, appliedStep, placementOf)
      : await toolkit.dispatch(doc, "superdoc_perform_action", {
          action: step.action,
          ...coreArguments(step, locator),
        }));
  if(step.action==='replace_text'&&!locator?.selector&&Array.isArray(step.arguments.edits)&&step.arguments.edits.length>1) {
    const counts=(receipt.executedOperations??[]).filter(operation=>operation.operationId==='doc.mutations.apply').map(operation=>operation.result?.stepCount);
    const completed=counts.length&&counts.every(Number.isSafeInteger)?counts.reduce((sum,count)=>sum+count,0):null;
    requireThat(completed===step.arguments.edits.length,'PARTIAL_MUTATION',`SDK replacement completed ${completed??'an unverified number of'} of ${step.arguments.edits.length} requested edits. Discard this failed session; no candidate is accepted.`);
  }
  return { ...verifyReceipt(step.action, receipt), intent: requestedAction };
}
