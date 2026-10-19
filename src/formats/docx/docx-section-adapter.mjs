import { requireThat } from '../../infrastructure/errors.mjs';
import { W, elements, descendants, hasProtected, hasOpaque, sectionSelection, sectionBlockText } from './docx-package-targets.mjs';

/** Retain complex blocks exactly; replace surrounding content without SDK normalization. */
export function replacePackageSection(document, styles, args, numbering) {
  const range = sectionSelection(document, styles, args.headingPath, args.subsections);
  requireThat(range.selected.length <= 100, 'LIMIT_EXCEEDED', 'Section replacement supports at most 100 source blocks.');
  const before = range.selected.map(sectionBlockText).join('\n');
  requireThat(before === args.expectedText, 'TARGET_CONFLICT', 'Section text differs from expectedText.');
  const blocks = args.blocks ?? args.paragraphs.map(text => ({ kind: 'paragraph', text }));
  const retained = blocks.filter(b => b.kind === 'retain').map(b => b.index);
  requireThat(new Set(retained).size === retained.length && retained.every(i => range.selected[i]), 'INVALID_ARGUMENT', 'Retain indices must name distinct existing section blocks.');
  requireThat(retained.every((index, i) => !i || index > retained[i - 1]), 'INVALID_ARGUMENT', 'Retained blocks must keep their original relative order.');
  for (const [index, block] of range.selected.entries()) requireThat(retained.includes(index) || !hasProtected(block) && !hasOpaque(block), 'UNSUPPORTED_FEATURE', 'Objects, comments, review ranges, references and section properties must be explicitly retained by block index.');
  const prefix = range.blocks.slice(0, range.start + 1).map(n => n.toString()), suffix = range.blocks.slice(range.end).map(n => n.toString());
  const element = name => document.createElementNS(W, 'w:' + name);
  const paragraph = text => { const p = element('p'), r = element('r'), t = element('t'); t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve'); t.textContent = text; r.appendChild(t); p.appendChild(r); return p; };
  let numberingChanged = false;
  const inserted = blocks.flatMap(block => {
    if (block.kind === 'retain') return range.selected[block.index].cloneNode(true);
    if (block.kind === 'paragraph' || block.kind === 'heading') {
      const p = paragraph(block.text);
      if (block.kind === 'heading') { const props = element('pPr'), outline = element('outlineLvl'); outline.setAttributeNS(W, 'w:val', String(block.level - 1)); props.appendChild(outline); p.insertBefore(props, p.firstChild); }
      return p;
    }
    if (block.kind === 'list') {
      requireThat(numbering?.documentElement.namespaceURI === W, 'UNSUPPORTED_FEATURE', 'An object-bearing section with new lists needs an existing numbering part. Create a list through the SDK in a separate phase first.');
      const nextId = (name, attr) => Math.max(0, ...descendants(numbering, name).map(n => Number(n.getAttributeNS(W, attr)))) + 1;
      const abstractId = nextId('abstractNum', 'abstractNumId'), numId = nextId('num', 'numId');
      requireThat(Number.isSafeInteger(abstractId) && Number.isSafeInteger(numId) && numId < 2147483647 && abstractId < 2147483647, 'INVALID_DOCUMENT', 'Invalid numbering IDs.');
      const n = name => numbering.createElementNS(W, 'w:' + name), prop = (name, value) => { const p = n(name); p.setAttributeNS(W, 'w:val', value); return p; };
      const abstract = n('abstractNum'); abstract.setAttributeNS(W, 'w:abstractNumId', String(abstractId)); abstract.appendChild(prop('multiLevelType', 'singleLevel'));
      const lvl = n('lvl'); lvl.setAttributeNS(W, 'w:ilvl', '0'); lvl.appendChild(prop('start', '1')); lvl.appendChild(prop('numFmt', block.listKind === 'ordered' ? 'decimal' : 'bullet')); lvl.appendChild(prop('lvlText', block.listKind === 'ordered' ? '%1.' : '•')); lvl.appendChild(prop('lvlJc', 'left')); abstract.appendChild(lvl);
      numbering.documentElement.insertBefore(abstract, descendants(numbering, 'num')[0] ?? null);
      const num = n('num'); num.setAttributeNS(W, 'w:numId', String(numId)); num.appendChild(prop('abstractNumId', String(abstractId))); numbering.documentElement.appendChild(num); numberingChanged = true;
      return block.items.map(text => { const p = paragraph(text), props = element('pPr'), numPr = element('numPr'); for (const [name, value] of [['ilvl', '0'], ['numId', String(numId)]]) { const node = element(name); node.setAttributeNS(W, 'w:val', value); numPr.appendChild(node); } props.appendChild(numPr); p.insertBefore(props, p.firstChild); return p; });
    }
    requireThat(block.kind === 'table' && block.rows.every(row => row.length === block.rows[0].length), 'INVALID_ARGUMENT', 'Replacement table must be rectangular.');
    const table = element('tbl'), grid = element('tblGrid'); table.appendChild(grid);
    for (let c = 0; c < block.rows[0].length; c++) { const column = element('gridCol'); column.setAttributeNS(W, 'w:w', '2400'); grid.appendChild(column); }
    for (const row of block.rows) { const tr = element('tr'); for (const text of row) { const tc = element('tc'); tc.appendChild(paragraph(text)); tr.appendChild(tc); } table.appendChild(tr); }
    return table;
  });
  const anchor = range.blocks[range.end] ?? elements(range.body, 'sectPr')[0] ?? null;
  for (const old of range.selected) range.body.removeChild(old);
  for (const node of inserted) range.body.insertBefore(node, anchor);
  const after = elements(range.body).filter(n => n.localName !== 'sectPr');
  requireThat(JSON.stringify(prefix) === JSON.stringify(after.slice(0, prefix.length).map(n => n.toString())) && JSON.stringify(suffix) === JSON.stringify(after.slice(prefix.length + inserted.length).map(n => n.toString())), 'PRESERVATION_FAILED', 'Section edit changed outside blocks.');
  return { before, after: inserted.map(sectionBlockText).join('\n'), retainedBlocks: retained.length, numberingChanged };
}

export function sectionInventory(parsed, section) {
  const document = parsed.get('word/document.xml'); if (!document) return [];
  const body = descendants(document, 'body')[0], allBlocks = elements(body).filter(n => n.localName !== 'sectPr');
  const range = section ? sectionSelection(document, parsed.get('word/styles.xml'), section.headingPath, section.subsections) : null;
  const blocks = range?.selected ?? allBlocks;
  // The stream gives exact indices required by retain, without pretending opaque blocks are paragraphs.
  return blocks.map((block, index) => ({ id: `word/document.xml#body-${allBlocks.indexOf(block) + 1}`, kind: 'sectionBlock', index, indexScope: section ? 'selected-section' : 'whole-body', ...(section ? { headingPath: section.headingPath } : {}), xmlKind: block.localName, text: sectionBlockText(block), protected: hasProtected(block), object: hasOpaque(block), contentSource: 'read-only-ooxml-package' }));
}
