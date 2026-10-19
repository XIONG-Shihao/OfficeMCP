import { OfficeError, requireThat } from "../../infrastructure/errors.mjs";
import { success } from "./docx-sdk.mjs";

/** Keep unchanged edge runs intact for both native and agent replacements. */
export async function replaceSelection(doc, { target, expectedText, text, changeMode = "direct", story }) {
  let selected = target;
  let replacement = text;
  if (target?.start?.kind === "text" && target?.end?.kind === "text" &&
      target.start.blockId === target.end.blockId &&
      target.end.offset - target.start.offset === expectedText.length) {
    let prefix = 0, suffix = 0;
    while (prefix < Math.min(expectedText.length, text.length) && expectedText[prefix] === text[prefix]) prefix++;
    while (suffix < Math.min(expectedText.length, text.length) - prefix && expectedText.at(-suffix - 1) === text.at(-suffix - 1)) suffix++;
    // SDK offsets are UTF-16; never split a surrogate pair at a trimmed edge.
    if (prefix && /[\uD800-\uDBFF]/.test(expectedText[prefix - 1]) && /[\uDC00-\uDFFF]/.test(expectedText[prefix] ?? "")) prefix--;
    if (suffix && /[\uDC00-\uDFFF]/.test(expectedText[expectedText.length - suffix])) suffix--;
    selected = {
      ...target,
      start: { ...target.start, offset: target.start.offset + prefix },
      end: { ...target.end, offset: target.end.offset - suffix },
    };
    replacement = text.slice(prefix, text.length - suffix);
    requireThat(selected.start.offset !== selected.end.offset || replacement.length, "INVALID_ARGUMENT", "Edit makes no change.");
  }
  try {
    return success(await doc.replace({
      ...(story ? { in: story } : {}), target: selected, text: replacement, changeMode,
    }));
  } catch (error) {
    if (error instanceof OfficeError) throw error;
    throw new OfficeError(error.message === "invalid-range" ? "UNSUPPORTED_FEATURE" : "ENGINE_FAILED",
      `SuperDoc 2.2.1 rejected this ${changeMode} text range: ${error.message}. No candidate was accepted.`);
  }
}

export async function replaceStoryText(doc, { story, find, replace, changeMode = "direct" }) {
  const match = await doc.query.match({
    in: story,
    select: { type: "text", pattern: find, mode: "contains", caseSensitive: true },
    require: "exactlyOne", offset: 0, limit: 2,
  });
  requireThat(match.total === 1 && match.items[0]?.target, "TARGET_CONFLICT", "Story text must match exactly once.");
  return replaceSelection(doc, { story, target: match.items[0].target, expectedText: find, text: replace, changeMode });
}
