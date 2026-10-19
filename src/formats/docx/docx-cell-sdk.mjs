import { packageCells, descendants, elements } from './docx-package-targets.mjs';
import { textOf } from './docx-sdk.mjs';
import { queryLocator } from './docx-agent-query.mjs';
import { parseXml } from '../opc/ooxml.mjs';
import { openFilePackage } from '../opc/file-package.mjs';
import { requireThat } from '../../infrastructure/errors.mjs';

/** Read-only eligibility check. An absent SDK range chooses the declared package adapter. */
export async function sdkCellEdit(doc, file, maxExpandedBytes, locator, args) {
  const archive = openFilePackage(file, maxExpandedBytes, maxExpandedBytes);
  let tree; try { tree = parseXml(Buffer.from(archive.parts['word/document.xml']).toString('utf8')); } finally { archive.close(); }
  const cell = packageCells(tree).find(c => c.id === locator.objectId);
  requireThat(cell && cell.text === args.expectedText, 'TARGET_CONFLICT', 'Selected cell is absent or expectedText is stale.');
  if (args.find === '') return { available: false, reason: 'Empty-cell insertion needs a guarded paragraph insertion' };
  const eligible = cell.paragraphs.filter(p => (args.paragraphIndex === undefined || p.index === args.paragraphIndex) && p.text.includes(args.find));
  requireThat(eligible.length === 1 && eligible[0].text.split(args.find).length === 2, 'TARGET_CONFLICT', 'Find must match exactly one own cell paragraph.');
  requireThat(!eligible[0].protected && !eligible[0].object && !/[\r\n\t]/.test(args.replace), 'UNSUPPORTED_FEATURE', 'Selected cell paragraph contains protected content or requires structural insertion.');
  const model = await doc.get({ options: { includeResolved: true } }), tables = [];
  function walk(nodes) { for (const node of nodes ?? []) { if (node.kind === 'table') { tables.push(node); for (const row of node.table.rows) for (const cell of row.cells) walk(cell.content); } else if (node.kind === 'sdt') walk(node.sdt?.content); } }
  walk(model.body);
  const xmlTables = descendants(tree, 'tbl');
  if (tables.length !== xmlTables.length) return { available: false, reason: 'SDK/package table inventories differ' };
  const table = tables[cell.tableIndex], xml = xmlTables[cell.tableIndex];
  if (!table || elements(xml, 'tr').length !== table.table.rows.length || !elements(xml, 'tr').every((row, i) => elements(row, 'tc').length === table.table.rows[i].cells.length)) return { available: false, reason: 'SDK normalizes this cell grid' };
  const selected = table.table.rows[cell.row]?.cells[cell.cell];
  if (!selected || selected.content.filter(n => ['paragraph', 'heading', 'listItem'].includes(n.kind)).map(textOf).join('\n') !== args.expectedText) return { available: false, reason: 'SDK omits selected cell text' };
  const id = table.paragraphIds?.paraId ?? table.id;
  if (!id) return { available: false, reason: 'SDK table has no address' };
  try {
    const cells = await doc.tables.getCells({ target: { kind: 'block', nodeType: 'table', nodeId: id } });
    const native = cells.cells.find(c => c.rowIndex === cell.row && c.columnIndex === cell.column);
    if (!native?.address) return { available: false, reason: 'SDK exposes no owner-cell address' };
    const match = await doc.query.match({ select: { type: 'text', pattern: args.find, caseSensitive: true, mode: 'contains' }, within: native.address, require: 'any', offset: 0, limit: 2 });
    if (match.total !== 1 || !match.items[0]?.target) return { available: false, reason: 'SDK cannot isolate the selected paragraph range' };
    return { available: true, locator: queryLocator(match.items[0]) };
  } catch { return { available: false, reason: 'SDK does not expose this cell range' }; }
}
