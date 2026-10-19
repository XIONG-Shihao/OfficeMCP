import { OfficeError, requireThat } from "../../infrastructure/errors.mjs";
// Complete-surface tools are thin adapters over the same SDK host commands used
// by docx-agent. They keep their public result shapes without a second DOCX engine.
export async function docxWorker(service, kind, input, args, other, signal) {
  const documentId = args.documentId ?? args.view?.documentId ?? other.documentId;
  const run = (kind, nativeArgs, extra = other) =>
    service.docxPool.run(kind, input,
      { ...nativeArgs, ...(documentId ? { documentId } : {}) }, extra, signal);
  if (kind === "office_create") return run("docx_native_create", args);
  if (kind.endsWith("_validate") || kind === "office_validate_created")
    return run("docx_native_validate", args);
  if (
    [
      "office_package_parts",
      "office_package_extract",
      "office_extract",
    ].includes(kind)
  )
    throw new OfficeError(
      "UNSUPPORTED_FEATURE",
      "Raw DOCX package access is not exposed by this SDK-only implementation.",
    );
  if (args.docxDomain)
    requireThat(
      !args.query &&
        !args.ocr &&
        !args.sheet &&
        !args.range &&
        !args.page &&
        !args.slide,
      "INVALID_ARGUMENT",
      "DOCX detail domains do not accept unrelated selectors.",
    );
  if (
    [
      "office_read",
      "docx_read",
      "docx_search",
      "docx_inspect",
      "docx_structure",
    ].includes(kind)
  ) {
    const r = await run("docx_native_read", {
      ...args,
      mode: kind === "docx_inspect" ? "inspect" : "read",
      domain:
        kind === "office_read"
          ? (args.docxDomain ??
            (args.scope === "overview" ? "paragraphs" : "content"))
          : "paragraphs",
      offset: ["office_read", "docx_structure"].includes(kind)
        ? (args.offset ?? 0)
        : 0,
      limit: ["office_read", "docx_structure"].includes(kind)
        ? (args.limit ?? 30)
        : 10000,
    });
    if (r.items) {
      r.items = r.items
        .filter((i) => !args.query || i.text?.includes(args.query))
        .map((i, n) => ({
          ...i,
          paragraphIndex: i.nativeTarget?.paragraphIndex ?? n,
          paragraphId: i.nativeTarget?.nodeId,
          id: i.nativeTarget?.nodeId,
          ...(i.nativeTarget
            ? {
                target:
                  i.nativeTarget.kind === "table"
                    ? {
                        partName: "word/document.xml",
                        tableIndex: i.nativeTarget.index,
                      }
                    : {
                        partName: "word/document.xml",
                        paragraphIndex: i.nativeTarget.paragraphIndex,
                      },
              }
            : {}),
        }));
    }
    if (kind === "docx_structure") {
      r.nodes = r.items;
      delete r.items;
    }
    return r;
  }
  if (kind === "docx_review_changes") {
    const r = await run("docx_native_read", {
        mode: "read",
        domain: "tracked_changes",
        offset: 0,
        limit: 100,
      }),
      matches = r.items.filter((i) =>
        Object.values(i.wordRevisionIds ?? {}).some((id) =>
          args.changeIds.includes(id),
        ),
      );
    requireThat(
      matches.length,
      "TARGET_CONFLICT",
      "No matching SDK tracked changes.",
    );
    return run("docx_native_edit", {
      steps: matches.map((i) => ({
        type: "review",
        action: args.decision,
        target: { id: i.id },
        locator: i.nativeTarget,
      })),
    });
  }
  if (["docx_table_cell", "docx_table_row"].includes(kind)) {
    const tables = await run("docx_native_read", {
        mode: "read",
        domain: "tables",
        offset: 0,
        limit: 100,
      }),
      item = tables.items[args.tableIndex ?? args.target?.tableIndex];
    requireThat(item, "TARGET_CONFLICT", "Table target does not exist.");
    const step =
      kind === "docx_table_cell"
        ? {
            type: "table_cells",
            action: "write",
            cells: [
              {
                row: args.row ?? args.target.row,
                column: args.column ?? args.target.column,
                expectedText: args.expectedText,
                text: args.text ?? args.replacement,
              },
            ],
          }
        : {
            type: "tables",
            action: args.action === "delete" ? "delete_rows" : "insert_rows",
            rowIndex: args.row ?? args.target?.row,
            position: args.action === "insert_before" ? "before" : "after",
            count: 1,
            values: args.values,
          };
    return run("docx_native_edit", {
      steps: [
        { ...step, target: { id: "legacy-table" }, locator: item.nativeTarget },
      ],
    });
  }
  const p = await run("docx_native_read", {
      mode: "read",
      domain: "paragraphs",
      offset: 0,
      limit: 10000,
    }),
    index = args.paragraphIndex ?? args.target?.paragraphIndex;
  let item = args.paragraphId
    ? (p.items.find((i) => i.nativeTarget.nodeId === args.paragraphId) ??
      (/^p_\d+$/.test(args.paragraphId)
        ? p.items[Number(args.paragraphId.slice(2))]
        : undefined))
    : p.items[index];
  requireThat(
    item?.nativeTarget,
    "TARGET_CONFLICT",
    "SDK paragraph target is unavailable. Read again.",
  );
  requireThat(
    !args.target?.partName || args.target.partName === "word/document.xml",
    "UNSUPPORTED_FEATURE",
    "Use docx_headers_footers for SDK header/footer editing.",
  );
  const locator = { ...item.nativeTarget };
  let step;
  if (kind === "docx_insert_paragraph")
    step = {
      type: "blocks",
      action: args.placement === "before" ? "insert_before" : "insert_after",
      expectedText: args.expectedText,
      text: args.text,
    };
  else if (kind === "docx_format_text") {
    requireThat(
      item.text === args.expectedText,
      "TARGET_CONFLICT",
      "Paragraph text differs.",
    );
    locator.start = args.start;
    locator.end = args.end;
    locator.text = item.text.slice(args.start, args.end);
    step = { type: "format_text", format: args.format };
  } else if (
    [
      "docx_replace_text",
      "docx_edit_paragraph",
      "docx_replace_span",
      "docx_format_span",
    ].includes(kind)
  ) {
    const start = item.text.indexOf(args.expectedText);
    requireThat(
      start >= 0 && item.text.indexOf(args.expectedText, start + 1) < 0,
      "TARGET_CONFLICT",
      "Expected text must occur exactly once in the selected paragraph.",
    );
    locator.start = start;
    locator.end = start + args.expectedText.length;
    locator.text = args.expectedText;
    step =
      kind === "docx_format_span"
        ? {
            type: "format_text",
            format: {
              ...args.format,
              ...(args.format.color ? { color: "#" + args.format.color } : {}),
            },
          }
        : {
            type: "text",
            action: "replace",
            expectedText: args.expectedText,
            content: args.replacement,
            changeMode: args.tracked ? "tracked" : "direct",
            author: args.author ?? "OfficeMCP",
          };
  } else
    throw new OfficeError(
      "UNSUPPORTED_FEATURE",
      "Use the DOCX preset for this SDK operation.",
    );
  return run("docx_native_edit", {
    steps: [{ ...step, target: { id: "legacy-paragraph" }, locator }],
  });
}
