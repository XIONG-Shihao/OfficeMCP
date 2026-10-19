import { requireThat } from '../../infrastructure/errors.mjs';
export const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const elements = (node, name) => Array.from(node?.childNodes ?? []).filter(n => n.nodeType === 1 && (!name || n.localName === name) && n.namespaceURI === W);
export const descendants = (node, name) => Array.from(node?.getElementsByTagNameNS(W, name) ?? []);
export const wordText = node => descendants(node, 't').map(t => t.textContent ?? '').join('');
const val = node => node?.getAttributeNS(W, 'val');
export const protectedNames = ['commentRangeStart', 'commentRangeEnd', 'commentReference', 'ins', 'del', 'fldChar', 'fldSimple', 'hyperlink', 'bookmarkStart', 'bookmarkEnd', 'footnoteReference', 'endnoteReference', 'sectPr', 'sdt', 'customXml', 'altChunk'];
export const opaqueNames = ['drawing', 'pict', 'object', 'oMath', 'oMathPara'];
export const hasProtected = node => protectedNames.some(name => descendants(node, name).length) || ['sdt', 'altChunk', 'customXml'].includes(node.localName);
export const hasOpaque = node => opaqueNames.some(name => Array.from(node.getElementsByTagNameNS('*', name)).length);
export const ownCellParagraphs = cell => elements(cell, 'p');
export const ownCellText = cell => ownCellParagraphs(cell).map(wordText).join('\n');

/** Physical cells are named explicitly. Grid coordinates also expose merge ownership. */
export function packageCells(document, partName = 'word/document.xml') {
  const tables = descendants(document, 'tbl'), result = [];
  const emptyContinuation = tc => ownCellText(tc) === '' && elements(tc).every(n => n.localName === 'tcPr' || n.localName === 'p') && !hasProtected(tc) && !hasOpaque(tc);
  for (const [tableIndex, table] of tables.entries()) {
    let vertical = new Map();
    const rows = elements(table, 'tr');
    for (const [row, tr] of rows.entries()) {
      let column = Number(val(elements(elements(tr, 'trPr')[0], 'gridBefore')[0]) ?? 0);
      requireThat(Number.isInteger(column) && column >= 0, 'INVALID_DOCUMENT', 'Invalid table grid offset.');
      const physical = elements(tr, 'tc'), groups = [];
      for (const [cell, tc] of physical.entries()) {
        const props = elements(tc, 'tcPr')[0], span = Number(val(elements(props, 'gridSpan')[0]) ?? 1), horizontal = elements(props, 'hMerge')[0], merge = elements(props, 'vMerge')[0];
        requireThat(Number.isInteger(span) && span > 0 && span <= 16384, 'INVALID_DOCUMENT', 'Invalid table grid span.');
        const entry = { cell, tc, column, span, merge };
        if (horizontal && val(horizontal) !== 'restart') {
          const group = groups.at(-1);
          requireThat(group?.horizontal && emptyContinuation(tc), 'UNSUPPORTED_FEATURE', 'Horizontal merge ownership is inconsistent or continuation content is nonempty.');
          group.entries.push(entry); group.width += span;
        } else groups.push({ horizontal: Boolean(horizontal), column, width: span, entries: [entry] });
        column += span;
      }
      const next = new Map();
      for (const group of groups) {
        const first = group.entries[0], id = `${partName}#table-${tableIndex + 1}/row-${row + 1}/cell-${first.cell + 1}`, merge = first.merge;
        requireThat(group.entries.every(e => Boolean(e.merge) === Boolean(merge) && (!merge || (val(e.merge) === 'restart') === (val(merge) === 'restart'))), 'UNSUPPORTED_FEATURE', 'Combined horizontal/vertical merge markers disagree.');
        let owner;
        if (merge && val(merge) !== 'restart') {
          owner = vertical.get(group.column);
          requireThat(owner && owner.columnSpan === group.width && Array.from({ length: group.width }, (_, i) => vertical.get(group.column + i)).every(item => item === owner) && group.entries.every(e => emptyContinuation(e.tc)), 'UNSUPPORTED_FEATURE', 'Vertical merge ownership is inconsistent or continuation content is nonempty.');
          owner.rowSpan++;
        } else owner = { id, columnSpan: group.width, rowSpan: 1, text: ownCellText(first.tc), nativeTarget: { kind: 'opc_object', objectKind: 'tableCell', objectId: id, partName, tableIndex, row, cell: first.cell } };
        if (merge) for (let i = 0; i < group.width; i++) next.set(group.column + i, owner);
        for (const entry of group.entries) {
          const itemId = `${partName}#table-${tableIndex + 1}/row-${row + 1}/cell-${entry.cell + 1}`;
          result.push({ id: itemId, kind: 'tableCell', partName, tableIndex, row, cell: entry.cell, column: entry.column, columnSpan: entry.span, ownerColumnSpan: owner.columnSpan, ownerId: owner.id, mergeContinuation: itemId !== owner.id, text: owner.text, paragraphs: ownCellParagraphs(entry.tc).map((p, i) => ({ index: i, text: wordText(p), protected: hasProtected(p), object: hasOpaque(p) })), nestedTables: elements(entry.tc, 'tbl').length, owner, nativeTarget: owner.nativeTarget });
        }
      }
      vertical = next;
    }
  }
  requireThat(result.length <= 100000, 'LIMIT_EXCEEDED', 'Package cell inventory exceeds 100000 physical cells.');
  return result.map(({ owner, ...item }) => ({ ...item, rowSpan: owner.rowSpan }));
}

export function headingLevel(p, styles) {
  if (p?.localName !== 'p') return undefined;
  const props = elements(p, 'pPr')[0], styleId = val(elements(props, 'pStyle')[0]);
  const style = styles && descendants(styles, 'style').find(s => s.getAttributeNS(W, 'styleId') === styleId);
  const outline = val(elements(props, 'outlineLvl')[0]) ?? val(elements(elements(style, 'pPr')[0], 'outlineLvl')[0]);
  if (outline !== undefined && Number(outline) < 9) return Number(outline) + 1;
  const named = /^Heading([1-9])$/i.exec(styleId ?? '');
  return named ? Number(named[1]) : undefined;
}
export function sectionRanges(document, styles) {
  const body = descendants(document, 'body')[0], blocks = elements(body).filter(n => n.localName !== 'sectPr'), stack = [], sections = [];
  for (const [index, block] of blocks.entries()) {
    const level = headingLevel(block, styles); if (!level) continue;
    stack.length = level - 1; stack[level - 1] = wordText(block);
    sections.push({ index, level, headingPath: [...stack], block });
  }
  return { body, blocks, sections };
}
export function sectionSelection(document, styles, headingPath, subsections) {
  const info = sectionRanges(document, styles), matches = info.sections.filter(s => JSON.stringify(s.headingPath) === JSON.stringify(headingPath));
  requireThat(matches.length === 1, 'TARGET_CONFLICT', 'Section headingPath must resolve exactly once.');
  const selected = matches[0]; let end = selected.index + 1;
  while (end < info.blocks.length) { const level = headingLevel(info.blocks[end], styles); if (level && (level <= selected.level || subsections === 'preserve')) break; end++; }
  return { ...info, start: selected.index, end, selected: info.blocks.slice(selected.index + 1, end) };
}
export function sectionBlockText(block) {
  if (block.localName === 'tbl') return elements(block, 'tr').map(tr => elements(tr, 'tc').map(tc => elements(tc).filter(n => n.localName !== 'tcPr').map(sectionBlockText).join('\n')).join('\t')).join('\n');
  return wordText(block);
}
