import { fingerprint } from "../../infrastructure/files.mjs";
import { parseXml } from "../opc/ooxml.mjs";
import { openFilePackage } from "../opc/file-package.mjs";
import { requireThat } from "../../infrastructure/errors.mjs";
import { clean, textOf, tables, list } from "./docx-sdk.mjs";
import { queryLocator } from "./docx-agent-query.mjs";
import { bodyModel, nodeId, blockLocator, cellText } from "./docx-change-readback.mjs";

const same = (left, right) => fingerprint(clean(left)) === fingerprint(clean(right));
const level = node => Number(node.heading?.level ?? node.attrs?.level ?? node.level ?? 1);

const children = node => Array.from(node.childNodes ?? []).filter(child => child.nodeType === 1);
function packageBody(task) {
  const archive = openFilePackage(task.input, task.maxExpandedBytes, task.maxFileBytes);
  let xml;
  try { xml = parseXml(Buffer.from(archive.parts['word/document.xml']).toString('utf8')); }
  finally { archive.close(); }
  const body = xml.getElementsByTagNameNS('*', 'body')[0];
  requireThat(body, "INVALID_DOCUMENT", "DOCX body is missing.");
  return children(body).filter(node => node.localName !== 'sectPr');
}

export async function semanticWorkflow(doc, workflow, locator, task) {
  if (workflow.type === "update_table_by_key") {
    requireThat(locator?.kind === "table", "INVALID_ARGUMENT", "Use a table selectorTarget.");
    const target = locator.address;
    const before = (await tables(doc)).find(entry => entry.address.nodeId === target.nodeId)?.node;
    requireThat(before, "TARGET_CONFLICT", "Selected table is unavailable.");
    const rows = before.table.rows, header = rows[0]?.cells.map(cellText);
    const bodyBefore = await bodyModel(doc);
    const bodyTables = bodyBefore.filter(node => node.kind === "table");
    const xmlTables = packageBody(task).filter(node => node.localName === 'tbl');
    requireThat(bodyTables.length === xmlTables.length, "UNSUPPORTED_FEATURE", "SDK and package table inventories differ; use explicit edits.");
    const xmlTable = xmlTables[bodyTables.findIndex(node => nodeId(node) === target.nodeId)];
    requireThat(xmlTable, "UNSUPPORTED_FEATURE", "The selected table has no matching package boundary.");
    const xmlRows = children(xmlTable).filter(node => node.localName === 'tr');
    const merges = [...Array.from(xmlTable.getElementsByTagNameNS('*', 'gridSpan')), ...Array.from(xmlTable.getElementsByTagNameNS('*', 'vMerge'))].length > 0;
    requireThat(xmlRows.length === rows.length && xmlRows.every((row, index) => children(row).filter(node => node.localName === 'tc').length === rows[index].cells.length),
      "UNSUPPORTED_FEATURE", "SDK and package table dimensions differ; use explicit edits.");
    requireThat(bodyBefore.some(node => nodeId(node) === target.nodeId), "UNSUPPORTED_FEATURE", "Keyed updates currently require a top-level body table; nested content outside the chosen cells is preserved.");
    requireThat(header?.length, "UNSUPPORTED_FEATURE", "Keyed updates require a first header row.");
    const attr = (node, name, fallback) => node ? node.getAttributeNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', name) || fallback : fallback;
    const xmlCells = row => children(row).filter(node => node.localName === 'tc');
    const spanOf = cell => Number(attr(cell.getElementsByTagNameNS('*', 'gridSpan')[0], 'val', '1'));
    const verticalOf = cell => cell.getElementsByTagNameNS('*', 'vMerge').length > 0;
    requireThat(xmlCells(xmlRows[0]).every(cell => spanOf(cell) === 1 && !verticalOf(cell)), "UNSUPPORTED_FEATURE", "Use distinct unmerged first-row column headings.");
    const physicalColumn = (row, logical) => {
      let at = Number(attr(xmlRows[row].getElementsByTagNameNS('*', 'gridBefore')[0], 'val', '0'));
      for (const [index, cell] of xmlCells(xmlRows[row]).entries()) {
        const width = spanOf(cell);
        if (at <= logical && logical < at + width) return { index, cell, width };
        at += width;
      }
      return null;
    };
    const normalize = value => workflow.caseSensitive ? value : value.toLocaleLowerCase();
    const column = name => {
      const matches = header.flatMap((value, index) => normalize(value) === normalize(name) ? [index] : []);
      requireThat(matches.length === 1, "TARGET_CONFLICT", `Header ${JSON.stringify(name)} must match exactly once.`);
      return matches[0];
    };
    const keyColumn = column(workflow.keyColumn), edits = [], seen = new Set();
    const addresses = (await doc.tables.getCells({ target })).cells;
    for (const update of workflow.updates) {
      const matches = rows.flatMap((row, index) => {
        const key = index > 0 ? physicalColumn(index, keyColumn) : null;
        return key && normalize(cellText(row.cells[key.index])) === normalize(update.key) ? [index] : [];
      });
      requireThat(matches.length === 1, "TARGET_CONFLICT", `Row key ${JSON.stringify(update.key)} must match exactly once.`);
      const row = matches[0], logicalColumn = column(update.column), value = physicalColumn(row, logicalColumn), key = physicalColumn(row, keyColumn);
      requireThat(value && key && value.width === 1 && key.width === 1 && !verticalOf(value.cell) && !verticalOf(key.cell),
        "UNSUPPORTED_FEATURE", "The chosen key or value cell is merged. Choose unique unmerged cells; a merge elsewhere in the table is allowed.");
      const col = value.index, identity = `${row}:${col}`, xmlCell = value.cell;
      const content = children(xmlCell).filter(node => node.localName !== 'tcPr');
      requireThat(content.length === 1 && content[0].localName === 'p' &&
        ['drawing', 'pict', 'object', 'fldChar', 'fldSimple', 'hyperlink', 'ins', 'del', 'sdt',
          'commentRangeStart', 'commentRangeEnd', 'bookmarkStart', 'bookmarkEnd', 'footnoteReference', 'endnoteReference']
          .every(tag => xmlCell.getElementsByTagNameNS('*', tag).length === 0),
        "UNSUPPORTED_FEATURE", "Keyed target cell must contain one plain paragraph without protected content.");
      requireThat(!seen.has(identity), "INVALID_ARGUMENT", "Repeated keyed update resolves to the same cell.");
      requireThat(cellText(rows[row].cells[col]) === update.expectedText, "TARGET_CONFLICT", "Table cell text changed.");
      const blocks = rows[row].cells[col].content ?? [];
      requireThat(blocks.length === 1 && blocks[0].kind === "paragraph" && nodeId(blocks[0]), "UNSUPPORTED_FEATURE", "The chosen cell needs one addressable plain paragraph.");
      requireThat(update.expectedText.length > 0 || !merges, "UNSUPPORTED_FEATURE", "Empty keyed cells in merged tables require an explicit cell insertion operation.");
      const address = addresses.find(cell => cell.rowIndex === row && cell.columnIndex === logicalColumn);
      requireThat(address?.address && address.colspan === 1 && address.rowspan === 1, "UNSUPPORTED_FEATURE", "The chosen cell is not exposed as an unmerged SDK cell.");
      let cellLocator;
      if (update.expectedText.length) {
        const match = await doc.query.match({ select: { type: "text", pattern: update.expectedText, mode: "contains", caseSensitive: true }, within: address.address, require: "any", offset: 0, limit: 2 });
        requireThat(match.total === 1 && match.items[0]?.target, "TARGET_CONFLICT", "The value does not resolve once within the chosen cell.");
        cellLocator = queryLocator(match.items[0]);
      }
      seen.add(identity);
      edits.push({ row, column: col, expectedText: update.expectedText, text: update.text, locator: cellLocator });
    }
    return {
      // Use the same SDK range replacement agents previously had to discover
      // manually. It preserves paragraph/run formatting and unrelated merges.
      steps: edits.map(edit => edit.expectedText.length
        ? { action: "replace_text", arguments: { edits: [{ find: edit.expectedText, replace: edit.text }], caseSensitive: true, changeMode: "direct" }, targetLocator: edit.locator }
        : { action: "write_table_cells", arguments: { cells: [{ row: edit.row, column: edit.column, expectedText: "", text: edit.text }] }, targetLocator: locator }),
      verify: async () => {
        requireThat(same(bodyBefore.filter(node => nodeId(node) !== target.nodeId), (await bodyModel(doc)).filter(node => nodeId(node) !== target.nodeId)),
          "VALIDATION_FAILED", "Content outside the selected table changed.");
        const after = (await tables(doc)).find(entry => entry.address.nodeId === target.nodeId)?.node;
        requireThat(after && after.table.rows.length === rows.length, "VALIDATION_FAILED", "Keyed update changed table dimensions.");
        for (const [r, row] of rows.entries()) {
          requireThat(after.table.rows[r].cells.length === row.cells.length, "VALIDATION_FAILED", "Keyed update changed table dimensions.");
          for (const [c, cell] of row.cells.entries()) {
            const edit = edits.find(edit => edit.row === r && edit.column === c);
            requireThat(edit ? cellText(after.table.rows[r].cells[c]) === edit.text : same(cell, after.table.rows[r].cells[c]),
              "VALIDATION_FAILED", edit ? "Updated cell differs from the request." : "Unselected table cell changed.");
          }
        }
        return { passed: true, checked: ["unique-headers-and-row-keys", "expected-cell-values", "changed-cell-readback", "unselected-cells-preserved", "outside-table-preserved"], changedCells: edits.length, method: "guarded-cell-text", mergesPreserved: merges };
      },
    };
  }
  if (workflow.type === "replace_section") {
    const replacements = workflow.blocks ?? workflow.paragraphs.map(text => ({ kind: 'paragraph', text }));
    requireThat(!replacements.some(b => b.kind === 'retain'), 'UNSUPPORTED_FEATURE', 'Retained opaque objects require the guarded package section adapter.');
    const before = await bodyModel(doc), stack = [], matches = [];
    for (const [index, node] of before.entries()) {
      if (node.kind !== "heading") continue;
      const depth = level(node); stack.length = Math.max(0, depth - 1); stack[depth - 1] = textOf(node);
      if (JSON.stringify(stack) === JSON.stringify(workflow.headingPath)) matches.push(index);
    }
    requireThat(matches.length === 1, "TARGET_CONFLICT", "Section headingPath must resolve exactly once among top-level body headings.");
    const start = matches[0], heading = before[start], depth = level(heading);
    let end = start + 1;
    while (end < before.length) {
      const node = before[end];
      if (node.kind === "heading" && (level(node) <= depth || workflow.subsections === "preserve")) break;
      end++;
    }
    const selected = before.slice(start + 1, end);
    const xmlBody = packageBody(task);
    requireThat(xmlBody.length === before.length && xmlBody.every((node, index) =>
      node.localName === 'p' ? ['paragraph', 'heading', 'listItem'].includes(before[index].kind) : node.localName === 'tbl' ? before[index].kind === 'table' : node.localName === 'sdt' && before[index].kind === 'sdt'),
      "UNSUPPORTED_FEATURE", "SDK and package body boundaries differ; section replacement is unsafe.");
    const protectedTags = ['drawing', 'pict', 'object', 'fldChar', 'fldSimple', 'hyperlink', 'ins', 'del',
      'commentRangeStart', 'commentRangeEnd', 'bookmarkStart', 'bookmarkEnd', 'footnoteReference', 'endnoteReference', 'sdt', 'sectPr'];
    requireThat(xmlBody.slice(start + 1, end).every(node => protectedTags.every(tag => node.getElementsByTagNameNS('*', tag).length === 0)),
      "UNSUPPORTED_FEATURE", "Section range contains protected package content; use explicit edits.");
    requireThat(selected.length <= 100, "LIMIT_EXCEEDED", "Section replacement supports at most 100 existing blocks.");
    requireThat(selected.every(node => ["paragraph", "heading", "listItem", "table"].includes(node.kind)), "UNSUPPORTED_FEATURE", "Section replacement needs SDK-exposed paragraphs, headings, lists or tables.");
    const blockText = node => node.kind === 'table' ? node.table.rows.map(row => row.cells.map(cellText).join('\t')).join('\n') : textOf(node);
    requireThat(selected.map(blockText).join("\n") === workflow.expectedText, "TARGET_CONFLICT", "Selected section text differs from expectedText.");
    // Reject inline images/fields/review/anchors rather than deleting hidden content.
    const plain = nodes => { for (const node of nodes) {
      if (node.kind === 'table') { for (const row of node.table.rows) for (const cell of row.cells) plain(cell.content ?? []); continue; }
      const content = node[node.kind];
      requireThat((content?.inlines ?? []).every(inline => inline.kind === "run" && typeof inline.run?.text === "string"),
        "UNSUPPORTED_FEATURE", "Section contains non-text inline content; use explicit edits.");
    } }; plain(selected);
    const prefix = before.slice(0, start + 1), suffix = before.slice(end);
    const locators = selected.filter(node => node.kind !== 'table').map(blockLocator);
    const tableEntries = await tables(doc);
    const deletedTables = selected.filter(node => node.kind === 'table').map(node => {
      const index = tableEntries.findIndex(entry => entry.address.nodeId === nodeId(node));
      requireThat(index >= 0, 'UNSUPPORTED_FEATURE', 'Selected table has no SDK deletion address.');
      return { action: 'delete_table', arguments: { changeMode: 'direct' }, targetLocator: { kind: 'table', address: tableEntries[index].address, nodeId: nodeId(node), index } };
    });
    const anchor = selected[0] ?? suffix[0];
    const placement = { at: anchor ? 'before' : 'document_end' }, targetLocator = anchor ? blockLocator(anchor) : undefined;
    const insertedSteps = replacements.map(block => {
      if (block.kind === 'table') {
        requireThat(block.rows.every(row => row.length === block.rows[0].length), 'INVALID_ARGUMENT', 'Replacement tables must be rectangular.');
        return { action: 'create_table', arguments: { rows: block.rows.length, columns: block.rows[0].length, cellTexts: block.rows, placement, changeMode: 'direct' }, targetLocator };
      }
      if (block.kind === 'list') return { action: 'append_list', arguments: { items: block.items, kind: block.listKind, placement, changeMode: 'direct' }, targetLocator };
      return { action: 'insert_paragraphs', ...(block.kind === 'paragraph' && (!anchor || anchor.kind === 'heading') ? { paragraphStyleId: 'Normal' } : {}), arguments: { texts: [block.text], ...(block.kind === 'heading' ? { headingLevel: block.level } : {}), placement, changeMode: 'direct' }, targetLocator };
    });
    return {
      steps: [
        ...insertedSteps,
        ...(locators.length ? [{ action: "delete_blocks", arguments: { changeMode: "direct" }, targetLocator: locators[0], targetLocators: locators }] : []),
        ...deletedTables,
      ],
      verify: async () => {
        const after = await bodyModel(doc), inserted = after.slice(prefix.length, after.length - suffix.length);
        requireThat(same(prefix, after.slice(0, prefix.length)) && same(suffix, suffix.length ? after.slice(-suffix.length) : []),
          "VALIDATION_FAILED", "Content outside the selected section changed.");
        const expected = replacements.flatMap(b => b.kind === 'list' ? b.items.map(text => ({ kind: 'paragraph', text, list: b.listKind })) : [{ kind: b.kind, text: b.kind === 'table' ? b.rows.map(row => row.join('\t')).join('\n') : b.text }]);
        const listItems = expected.some(n => n.list) ? await list(doc, 'lists.list') : [];
        const listIds = new Map(listItems.map(item => [item.address.nodeId, item.kind]));
        requireThat(inserted.length === expected.length && inserted.every((node, index) => (node.kind === expected[index].kind || expected[index].list && node.kind === 'listItem') && blockText(node) === expected[index].text && (!expected[index].list || listIds.get(nodeId(node)) === expected[index].list)),
          "VALIDATION_FAILED", `Section replacement readback differs. Expected block kinds: ${expected.map(n => n.kind).join(',')}; observed: ${inserted.map(n => n.kind).join(',')}; text matches: ${inserted.map((n,i) => blockText(n) === expected[i]?.text).join(',')}; list membership: ${inserted.map(n => listIds.has(nodeId(n))).join(',')}.`);
        return { passed: true, checked: ["unique-heading-path", "expected-section-text", "replacement-readback", "outside-section-preserved"], removedBlocks: selected.length, insertedBlocks: inserted.length, subsections: workflow.subsections };
      },
    };
  }
  throw Error(`Unsupported semantic workflow: ${workflow.type}`);
}
