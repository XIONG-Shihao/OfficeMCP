import { success } from "./docx-sdk.mjs";

/** Shared native insertion; adapters retain their own target and text guards. */
export async function insertParagraphsAt(doc, { target, placement, texts, headingLevel, changeMode, paragraphStyleId }) {
  let at = { kind: placement, target };
  const results = [];
  for (const [index, text] of texts.entries()) {
    const heading = index === 0 && headingLevel;
    const result = success(await (heading ? doc.create.heading : doc.create.paragraph)({
      at, text,
      ...(changeMode !== undefined ? { changeMode } : {}),
      ...(heading ? { level: headingLevel } : {}),
    }));
    results.push(result);
    if (!heading && paragraphStyleId) success(await doc.styles.paragraph.setStyle({ target: result.paragraph, styleId: paragraphStyleId }));
    at = { kind: "after", target: result.paragraph ?? result.heading };
  }
  return results;
}
