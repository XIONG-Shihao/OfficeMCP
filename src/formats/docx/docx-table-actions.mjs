import { requireThat } from "../../infrastructure/errors.mjs";
import { success, tables, textOf } from "./docx-sdk.mjs";

export async function requireSimpleTable(doc, locator) {
  requireThat(locator?.kind === "table", "INVALID_ARGUMENT", "Use a table selectorTarget.");
  const target = locator.address;
  const cells = await doc.tables.getCells({ target });
  requireThat(
    cells.cells.every((cell) => cell.colspan === 1 && cell.rowspan === 1),
    "UNSUPPORTED_FEATURE",
    "Merged-cell mutations are not qualified.",
  );
  return target;
}

export async function writeTableCells(doc, locator, edits) {
  requireThat(edits.length > 0 && edits.length <= 100, "INVALID_ARGUMENT", "Provide 1-100 table cells.");
  const target = await requireSimpleTable(doc, locator);
  const model = (await tables(doc)).find(
    (item) => item.address.nodeId === target.nodeId,
  );
  requireThat(model, "TARGET_CONFLICT", "Table is unavailable.");
  const seen = new Set();
  for (const edit of edits) {
    const key = `${edit.row}:${edit.column}`;
    requireThat(!seen.has(key), "INVALID_ARGUMENT", "Duplicate cell coordinates.");
    seen.add(key);
    const blocks = model.node.table.rows[edit.row]?.cells[edit.column]?.content;
    requireThat(
      blocks?.length === 1 && blocks[0].kind === "paragraph",
      "UNSUPPORTED_FEATURE",
      "Only one plain paragraph per target cell is qualified.",
    );
    requireThat(textOf(blocks[0]) === edit.expectedText, "TARGET_CONFLICT", "Table cell text changed.");
  }
  for (const edit of edits)
    success(await doc.tables.setCellText({
      target,
      rowIndex: edit.row,
      columnIndex: edit.column,
      text: edit.text,
    }));
  return { changed: edits.map(({ row, column }) => ({ row, column })) };
}

export async function mutateTableRows(doc, locator, { action, rowIndex, position = "after", count = 1, values, changeMode = "direct", dryRun = false, requireSimple = false }) {
  requireThat(locator?.kind === "table", "INVALID_ARGUMENT", "Use a table selectorTarget.");
  const target = requireSimple ? await requireSimpleTable(doc, locator) : locator.address;
  const before = (await tables(doc)).find(item => item.address.nodeId === target.nodeId)?.node.table.rows;
  requireThat(before && Number.isInteger(rowIndex) && rowIndex >= 0 && rowIndex < before.length,
    "TARGET_CONFLICT", "Table row does not exist.");
  requireThat(["insert", "delete"].includes(action), "INVALID_ARGUMENT", "Use insert or delete rows.");
  const above = ["before", "above"].includes(position);
  const result = success(action === "delete"
    ? await doc.tables.deleteRow({ target, rowIndex, changeMode, dryRun })
    : await doc.tables.insertRow({ target, rowIndex, position: above ? "above" : "below", count, changeMode, dryRun }));
  if (action === "insert" && values && !dryRun)
    for (const [columnIndex, text] of values.entries())
      success(await doc.tables.setCellText({ target, rowIndex: rowIndex + (above ? 0 : 1), columnIndex, text, changeMode }));
  const after = (await tables(doc)).find(item => item.address.nodeId === target.nodeId)?.node.table.rows ?? [];
  requireThat(after.length === before.length + (dryRun ? 0 : action === "delete" ? -1 : count),
    "VALIDATION_FAILED", "Table row count differs from the requested edit.");
  return result;
}
