import { replaceSelection, replaceStoryText } from "./docx-text-actions.mjs";
import { decideTrackedChange } from "./docx-review-actions.mjs";
import { requireThat, OfficeError } from "../../infrastructure/errors.mjs";
import { paragraphs, textOf, list, success } from "./docx-sdk.mjs";
import { mutateTableRows, writeTableCells } from "./docx-table-actions.mjs";
import { insertParagraphsAt } from "./docx-paragraph-actions.mjs";
export async function mutate(doc, s) {
  const t = s.locator;
  let row, priorRows;
  if (["text", "paragraph"].includes(t.kind)) {
    priorRows = await paragraphs(doc);
    row = priorRows.find((r) => r.address.nodeId === t.nodeId);
    requireThat(
      row && textOf(row).slice(t.start, t.end) === t.text,
      "TARGET_CONFLICT",
      "Target text changed. Read/find again.",
    );
  }
  const target = row?.address ?? t.address;
  let r;
  if (s.type === "text") {
    requireThat(
      s.expectedText === t.text,
      "TARGET_CONFLICT",
      "Expected text differs from the selected target.",
    );
    let start = t.start,
      end = t.end,
      text = s.action === "delete" ? "" : s.content;
    if (s.action === "insert_before") end = start;
    else if (s.action === "insert_after") start = end;
    const before = textOf(row);
    requireThat(
      start !== end || text.length,
      "INVALID_ARGUMENT",
      "Edit makes no change.",
    );
    r = await replaceSelection(doc, {
      target: {
        kind: "selection",
        start: { kind: "text", blockId: t.nodeId, offset: start },
        end: { kind: "text", blockId: t.nodeId, offset: end },
      },
      expectedText: s.action.startsWith("insert_") ? "" : t.text,
      text,
      changeMode: s.changeMode,
    });
    const afterRows = await paragraphs(doc),
      expected = priorRows.map(textOf);
    expected[priorRows.indexOf(row)] =
      before.slice(0, start) + text + before.slice(end);
    requireThat(
      JSON.stringify(afterRows.map(textOf)) === JSON.stringify(expected),
      "VALIDATION_FAILED",
      "SDK body text differs from the requested edit; no candidate accepted.",
    );
  } else if (s.type === "format_text")
    r = await doc.format.apply({
      blockId: t.nodeId,
      start: t.start,
      end: t.end,
      inline: s.format,
      changeMode: "direct",
    });
  else if (s.type === "blocks") {
    requireThat(
      t.kind === "paragraph" && textOf(row) === s.expectedText,
      "TARGET_CONFLICT",
      "Use a full paragraph target and matching text.",
    );
    [r] = await insertParagraphsAt(doc, {
      target, placement: s.action === "insert_before" ? "before" : "after", texts: [s.text],
    });
  } else if (s.type === "format_paragraph") {
    const f = s.format;
    if (f.alignment !== undefined)
      success(
        await doc.format.paragraph.setAlignment({
          target,
          alignment: f.alignment,
        }),
      );
    if (f.before !== undefined || f.after !== undefined)
      success(
        await doc.format.paragraph.setSpacing({
          target,
          ...(f.before !== undefined ? { before: f.before * 20 } : {}),
          ...(f.after !== undefined ? { after: f.after * 20 } : {}),
        }),
      );
    const keep = Object.fromEntries(
      Object.entries(f).filter(([k]) =>
        ["keepNext", "keepLines", "widowControl"].includes(k),
      ),
    );
    if (Object.keys(keep).length)
      success(await doc.format.paragraph.setKeepOptions({ target, ...keep }));
  } else if (s.type === "styles")
    r = await doc.styles.paragraph.setStyle({ target, styleId: s.styleId });
  else if (s.type === "lists")
    r = await doc.lists.create({
      mode: "fromParagraphs",
      target,
      kind: s.kind,
      level: s.level,
    });
  else if (s.type === "table_layout")
    r = await doc.tables.setColumnWidth({
      target,
      columnIndex: s.columnIndex,
      widthPt: s.width,
    });
  else if (s.type === "tables" || s.type === "table_cells") {
    if (s.type === "tables")
      r = await mutateTableRows(doc, t, {
        action: s.action === "delete_rows" ? "delete" : "insert", rowIndex: s.rowIndex,
        position: s.position, count: s.count, values: s.values, changeMode: s.changeMode, requireSimple: true,
      });
    else r = await writeTableCells(doc, t, s.cells);
  } else if (s.type === "sections")
    r = await doc.sections.setBreakType({ target, breakType: s.breakType });
  else if (s.type === "page_setup") {
    const size = {},
      margins = {};
    for (const [k, v] of Object.entries(s.setup)) {
      if (["width", "height", "orientation"].includes(k))
        size[k] = k === "orientation" ? v : v / 72;
      else margins[k] = v / 72;
    }
    if (Object.keys(size).length)
      success(await doc.sections.setPageSetup({ target, ...size }));
    if (Object.keys(margins).length)
      success(await doc.sections.setPageMargins({ target, ...margins }));
  } else if (s.type === "images")
    r = await doc.images.setAltText({
      imageId: t.nodeId,
      description: s.description,
    });
  else if (s.type === "image_layout")
    r = await doc.images.setSize({
      imageId: t.nodeId,
      size: { width: s.width, height: s.height, unit: "pt" },
    });
  else if (s.type === "hyperlinks")
    r = await doc.hyperlinks.wrap({
      target: {
        kind: "text",
        blockId: t.nodeId,
        range: { start: t.start, end: t.end },
      },
      link: { destination: { href: s.url } },
    });
  else if (s.type === "bookmarks") {
    requireThat(
      t.kind === "paragraph",
      "UNSUPPORTED_FEATURE",
      "Bookmark insertion is qualified for full paragraphs only.",
    );
    requireThat(
      !(await list(doc, "bookmarks.list")).some((b) => b.name === s.name),
      "INVALID_ARGUMENT",
      "Bookmark name already exists.",
    );
    r = await doc.bookmarks.insert({
      name: s.name,
      at: {
        kind: "text",
        segments: [
          { blockId: t.nodeId, range: { start: t.start, end: t.end } },
        ],
      },
    });
  } else if (s.type === "notes")
    r = await doc.footnotes.insert({
      at: {
        kind: "text",
        segments: [{ blockId: t.nodeId, range: { start: t.end, end: t.end } }],
      },
      type: s.noteType,
      content: s.text,
    });
  else if (s.type === "fields")
    r = await doc.fields.insert({
      mode: "raw",
      at: {
        kind: "text",
        segments: [{ blockId: t.nodeId, range: { start: t.end, end: t.end } }],
      },
      instruction: s.fieldType,
      cachedResultText: s.cachedText,
    });
  else if (s.type === "toc")
    r = await doc.create.tableOfContents({
      at: { kind: "before", target },
      instruction: `TOC \\o "${s.fromLevel}-${s.toLevel}" \\h`,
    });
  else if (s.type === "content_controls") {
    const c = await doc.contentControls.get({ target });
    requireThat(
      c.controlType === "text" &&
        c.lockMode === "unlocked" &&
        c.text === s.expectedText,
      "TARGET_CONFLICT",
      "Select an unlocked plain text control with matching text.",
    );
    r = await doc.contentControls.text.setValue({ target, value: s.text });
    requireThat(
      (await doc.contentControls.get({ target })).text === s.text,
      "VALIDATION_FAILED",
      "SDK control value differs.",
    );
  } else if (s.type === "review") {
    r = await decideTrackedChange(doc, t, s.action);
  } else if (s.type === "headers_footers") {
    requireThat(
      t.story,
      "INVALID_ARGUMENT",
      "Use a resolved header/footer target.",
    );
    const part = (await list(doc, "headerFooters.parts.list")).find(
      (p) => p.refId === t.story.refId,
    );
    requireThat(
      !(await list(doc, "fields.list")).some((f) =>
        f.address?.storyId?.includes(part?.partPath ?? "__missing__"),
      ),
      "UNSUPPORTED_FEATURE",
      "Field-bearing header/footer editing is not qualified.",
    );
    const before = await doc.getText({ in: t.story });
    requireThat(
      before === s.expectedText,
      "TARGET_CONFLICT",
      "Header/footer text differs.",
    );
    requireThat(
      before.length,
      "UNSUPPORTED_FEATURE",
      "Use creation for empty header/footer content.",
    );
    r = await replaceStoryText(doc, { story: t.story, find: before, replace: s.text });
    requireThat(
      (await doc.getText({ in: t.story })) === s.text,
      "VALIDATION_FAILED",
      "SDK header/footer text differs.",
    );
  } else
    throw new OfficeError(
      "UNSUPPORTED_FEATURE",
      "This action has no SDK-only implementation.",
    );
  success(r);
  return r;
}
