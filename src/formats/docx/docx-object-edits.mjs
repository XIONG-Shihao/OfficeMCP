/** Explicit object-only OOXML changes; no body rewrite or repair fallback. */
import path from 'node:path';
import { parseXml } from '../opc/ooxml.mjs';
import { requireThat } from '../../infrastructure/errors.mjs';
import { changeChartData } from './docx-chart-data.mjs';
import { packageCells, descendants as wordNodes, ownCellParagraphs, ownCellText, hasProtected, hasOpaque, wordText } from './docx-package-targets.mjs';
import { replacePackageSection } from './docx-section-adapter.mjs';
const nodes = (root, name) => Array.from(root.getElementsByTagNameNS('*', name));
const children = (root, name) => Array.from(root.childNodes ?? []).filter(n => n.nodeType === 1 && n.localName === name);
const attribute = (root, name) => Array.from(root.attributes ?? []).find(a => a.localName === name)?.value;
const visible = root => nodes(root, 't').map(n => n.textContent ?? '').join('');

function replaceText(root, expected, find, replacement) {
  const namespace = root.localName === 'oMath' ? 'http://schemas.openxmlformats.org/officeDocument/2006/math' : ['txbxContent', 'p'].includes(root.localName) ? 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' : 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const fields = Array.from(root.getElementsByTagNameNS(namespace, 't')), text = fields.map(n => n.textContent ?? '').join('');
  requireThat(text === expected, 'TARGET_CONFLICT', 'Object text differs from expectedText.');
  requireThat(find && text.split(find).length === 2, 'TARGET_CONFLICT', 'Literal text must occur exactly once in the selected object.');
  requireThat(!/[\r\n\t]/.test(replacement), 'UNSUPPORTED_FEATURE', 'Literal object edits cannot insert paragraph breaks or tabs.');
  requireThat(!['commentRangeStart', 'commentRangeEnd', 'commentReference', 'ins', 'del', 'fldChar', 'fldSimple', 'hyperlink', 'bookmarkStart', 'bookmarkEnd'].some(name => nodes(root, name).length),
    'UNSUPPORTED_FEATURE', 'Object text contains protected review, field or reference content.');
  const start = text.indexOf(find), end = start + find.length; let at = 0, inserted = false;
  if (root.localName === 'txbxContent') {
    const paragraphs = Array.from(root.getElementsByTagNameNS(namespace, 'p'));
    let offset = 0;
    requireThat(paragraphs.some(p => {
      const length = Array.from(p.getElementsByTagNameNS(namespace, 't')).reduce((sum, n) => sum + n.textContent.length, 0);
      const contains = start >= offset && end <= offset + length; offset += length; return contains;
    }), 'UNSUPPORTED_FEATURE', 'Textbox edits must stay within one paragraph.');
  }
  for (const node of fields) {
    const before = node.textContent ?? '', next = at + before.length;
    if (next > start && at < end) {
      const prefix = before.slice(0, Math.max(0, start - at)), suffix = before.slice(Math.max(0, end - at));
      node.textContent = prefix + (inserted ? '' : replacement) + suffix; inserted = true;
      if (/^\s|\s$/.test(node.textContent)) node.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
    }
    at = next;
  }
  requireThat(visible(root) === text.replace(find, replacement), 'VALIDATION_FAILED', 'Object replacement readback differs.');
  return { before: text, after: visible(root) };
}
function commentParts(archive) {
  const names = Object.keys(archive.parts), protectedParts = new Set(names.filter(n => /^word\/comments[^/]*\.xml$/i.test(n))), queue = [...protectedParts];
  while (queue.length) {
    const source = queue.pop(), relName = path.posix.join(path.posix.dirname(source), '_rels', path.posix.basename(source) + '.rels');
    if (!archive.parts[relName]) continue;
    const rels = parseXml(Buffer.from(archive.parts[relName]).toString('utf8'));
    for (const rel of nodes(rels, 'Relationship')) {
      if (rel.getAttribute('TargetMode') === 'External') continue;
      const value = rel.getAttribute('Target');
      const target = value.startsWith('/') ? value.slice(1) : path.posix.normalize(path.posix.join(path.posix.dirname(source), value));
      if (!protectedParts.has(target)) { protectedParts.add(target); queue.push(target); }
    }
  }
  return protectedParts;
}
export function prepareObjectChanges(archive, steps) {
  const types = parseXml(Buffer.from(archive.parts['[Content_Types].xml'] ?? []).toString('utf8'));
  const T = 'http://schemas.openxmlformats.org/package/2006/content-types';
  requireThat(types.documentElement?.namespaceURI === T && types.documentElement.localName === 'Types', 'INVALID_DOCUMENT', 'OPC content types are missing or invalid.');
  const defaults = new Map(Array.from(types.getElementsByTagNameNS(T, 'Default')).map(n => [n.getAttribute('Extension').toLowerCase(), n.getAttribute('ContentType')]));
  const overrides = new Map(Array.from(types.getElementsByTagNameNS(T, 'Override')).map(n => [decodeURI(n.getAttribute('PartName')), n.getAttribute('ContentType')]));
  for (const name of Object.keys(archive.parts).filter(n => n !== '[Content_Types].xml' && !n.endsWith('/'))) requireThat(overrides.get('/' + name) || defaults.get(name.split('.').at(-1).toLowerCase()), 'INVALID_DOCUMENT', 'Package member has no declared content type: ' + name);
  requireThat(!Object.keys(archive.parts).some(n => n.startsWith('_xmlsignatures/')), 'UNSUPPORTED_FEATURE', 'Signed packages require a separate signing workflow.');
  if (archive.parts['word/settings.xml']) {
    const settings = parseXml(Buffer.from(archive.parts['word/settings.xml']).toString('utf8'));
    requireThat(!nodes(settings, 'documentProtection').some(n => ['1', 'true', 'on'].includes(attribute(n, 'enforcement'))), 'UNSUPPORTED_FEATURE', 'Protected documents require an explicit authorized unlocking workflow.');
  }
  const parts = new Map(), documents = new Map(), readback = [], protectedParts = commentParts(archive);
  const document = name => {
    requireThat(archive.parts[name] && !protectedParts.has(name), 'UNSUPPORTED_FEATURE', 'This object is absent or belongs to read-only comments.');
    if (!documents.has(name)) documents.set(name, parseXml(Buffer.from(archive.parts[name]).toString('utf8')));
    return documents.get(name);
  };
  for (const step of steps) {
    const locator = step.targetLocator, args = step.arguments;
    requireThat((step.targetLocators ?? [locator]).length === 1, 'INVALID_ARGUMENT', 'One object target is required per guarded action.');
    requireThat(locator?.kind === 'opc_object' && locator.partName && locator.objectId, 'INVALID_ARGUMENT', 'Use an object target returned by docx_read/query.');
    const tree = document(locator.partName); let result, selected;
    if (step.action === 'replace_section_content') {
      requireThat(locator.objectKind === 'section' && locator.partName === 'word/document.xml', 'INVALID_ARGUMENT', 'Section adapter uses the body story.');
      const styles = archive.parts['word/styles.xml'] ? document('word/styles.xml') : null;
      const numbering = archive.parts['word/numbering.xml'] ? document('word/numbering.xml') : null;
      result = replacePackageSection(tree, styles, args, numbering);
      if (result.numberingChanged) parts.set('word/numbering.xml', Buffer.from(numbering.toString()));
    } else if (step.action === 'replace_table_cell_text') {
      requireThat(locator.objectKind === 'tableCell' && locator.partName === 'word/document.xml', 'INVALID_ARGUMENT', 'Use a body table-cell package target.');
      const cells = packageCells(tree), cell = cells.find(c => c.nativeTarget.objectId === locator.objectId && c.id === locator.objectId);
      requireThat(cell, 'TARGET_CONFLICT', 'The selected table-cell owner is unavailable.');
      const table = wordNodes(tree, 'tbl')[cell.tableIndex], tc = Array.from(table.childNodes).filter(n => n.localName === 'tr')[cell.row];
      const physical = Array.from(tc.childNodes).filter(n => n.localName === 'tc')[cell.cell];
      requireThat(ownCellText(physical) === args.expectedText, 'TARGET_CONFLICT', 'Table cell text differs from expectedText.');
      const paragraphs = ownCellParagraphs(physical), matching = paragraphs.filter((p, index) => (args.paragraphIndex === undefined || args.paragraphIndex === index) && wordText(p).includes(args.find));
      requireThat(matching.length === 1, 'TARGET_CONFLICT', 'Find must identify exactly one own paragraph in this cell; use paragraphIndex to disambiguate.');
      selected = matching[0];
      requireThat(!hasProtected(selected) && !hasOpaque(selected), 'UNSUPPORTED_FEATURE', 'The selected paragraph contains protected or opaque content. Edit an ordinary paragraph or use its dedicated object action.');
      if (args.find === '') {
        requireThat(wordText(selected) === '' && args.replace && !/[\r\n\t]/.test(args.replace), 'INVALID_ARGUMENT', 'Empty find inserts plain text only into one empty paragraph.');
        const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
        const existing = Array.from(selected.childNodes).find(n => n.namespaceURI === W && n.localName === 'r');
        requireThat(!existing || Array.from(existing.childNodes).every(n => ['rPr', 't'].includes(n.localName)), 'UNSUPPORTED_FEATURE', 'Empty paragraph contains unsupported run content.');
        const run = existing ?? tree.createElementNS(W, 'w:r'), text = tree.createElementNS(W, 'w:t'); text.textContent = args.replace; text.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve'); run.appendChild(text); if (!existing) selected.appendChild(run);
      } else replaceText(selected, wordText(selected), args.find, args.replace);
      result = { before: args.expectedText, after: ownCellText(physical) };
    } else if (step.action === 'replace_textbox_text' || step.action === 'replace_equation_text') {
      const type = step.action === 'replace_textbox_text' ? 'textbox' : 'equation', tag = type === 'textbox' ? 'txbxContent' : 'oMath';
      requireThat(locator.objectKind === type && /^word\/(document|header\d+|footer\d+|footnotes|endnotes)\.xml$/.test(locator.partName), 'INVALID_ARGUMENT', 'Object kind or story is outside this adapter.');
      selected = nodes(tree, tag)[locator.index];
      requireThat(selected, 'TARGET_CONFLICT', 'Object index is unavailable.');
      requireThat(selected.namespaceURI === (type === 'textbox' ? 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' : 'http://schemas.openxmlformats.org/officeDocument/2006/math'), 'UNSUPPORTED_FEATURE', 'Object namespace is outside the declared adapter.'); result = replaceText(selected, args.expectedText, args.find, args.replace);
    } else if (step.action === 'set_chart_title' || step.action === 'set_chart_data') {
      requireThat(locator.objectKind === 'chart' && tree.documentElement.localName === 'chartSpace' && tree.documentElement.namespaceURI === 'http://schemas.openxmlformats.org/drawingml/2006/chart', 'INVALID_ARGUMENT', 'Use a chart object target.');
      if (step.action === 'set_chart_data') {
        const working = { ...archive, parts: new Proxy(archive.parts, { get: (target, name) => parts.get(name) ?? target[name] }) };
        result = changeChartData(working, locator.partName, tree, args);
        for (const [name, bytes] of result.parts) { requireThat(!protectedParts.has(name), 'UNSUPPORTED_FEATURE', 'Chart workbook belongs to read-only comments.'); parts.set(name, bytes); }
      } else {
      const chart = children(tree.documentElement, 'chart')[0]; selected = chart && children(chart, 'title')[0];
      requireThat(selected && nodes(selected, 't').length > 0 && !nodes(selected, 'strRef').length && !nodes(selected, 'f').length,
        'UNSUPPORTED_FEATURE', 'Only an existing direct rich-text chart title is supported; computed or missing titles need another adapter.');
      result = replaceText(selected, args.expectedText, args.expectedText, args.text);
      }
    } else throw new Error('Unknown guarded object action: ' + step.action);
    parts.set(locator.partName, Buffer.from(tree.toString(), 'utf8'));
    readback.push({ kind: locator.objectKind, change: 'modified', before: { text: result.before, truncated: false }, after: { text: result.after, truncated: false }, nativeTarget: locator, targetKind: 'object' });
  }
  return { parts, readback };
}
