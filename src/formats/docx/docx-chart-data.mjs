import path from 'node:path';
import { unzipSync, zipSync } from 'fflate';
import { parseXml } from '../opc/ooxml.mjs';
import { checkArchiveLimits } from '../opc/archive-limits.mjs';
import { requireThat } from '../../infrastructure/errors.mjs';

const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const direct = (node, name, ns = C) => Array.from(node?.childNodes ?? []).filter(n => n.nodeType === 1 && n.localName === name && n.namespaceURI === ns);
const descendants = (node, name, ns = C) => Array.from(node?.getElementsByTagNameNS(ns, name) ?? []);
const one = (list, message) => { requireThat(list.length === 1, 'UNSUPPORTED_FEATURE', message); return list[0]; };
const same = (a, b) => JSON.stringify(a) === JSON.stringify(b);
const relPart = name => path.posix.join(path.posix.dirname(name), '_rels', path.posix.basename(name) + '.rels');
const resolvePart = (from, to) => to.startsWith('/') ? to.slice(1) : path.posix.normalize(path.posix.join(path.posix.dirname(from), to));

function points(cache, numeric) {
  const pt = direct(cache, 'pt'), count = Number(one(direct(cache, 'ptCount'), 'Chart cache needs one point count.').getAttribute('val'));
  requireThat(count > 0 && count <= 1000 && pt.length === count && pt.every((p, i) => Number(p.getAttribute('idx')) === i), 'UNSUPPORTED_FEATURE', 'Sparse or irregular chart caches require another adapter.');
  const values = pt.map(p => one(direct(p, 'v'), 'Chart points need one literal value.').textContent);
  requireThat(!numeric || values.every(v => v.trim() && Number.isFinite(Number(v))), 'UNSUPPORTED_FEATURE', 'Chart numeric values must be finite literals.');
  return { pt, values: numeric ? values.map(Number) : values };
}
function binding(series, key, numeric) {
  const container = one(direct(series, key), 'Unsupported chart series axis.');
  const source = one(Array.from(container.childNodes).filter(n => n.nodeType === 1), 'One chart data source is required.');
  const allowed = numeric ? ['numLit', 'numRef'] : ['strLit', 'strRef'];
  requireThat(source.namespaceURI === C && allowed.includes(source.localName), 'UNSUPPORTED_FEATURE', 'Only string categories and numeric values are supported.');
  const referenced = source.localName.endsWith('Ref');
  const cache = referenced ? one(direct(source, numeric ? 'numCache' : 'strCache'), 'Referenced charts require complete cached data.') : source;
  const formula = referenced ? one(direct(source, 'f'), 'Chart reference is missing.').textContent : undefined;
  const data = points(cache, numeric);
  return { cache, formula, pt: data.pt, cachedValues: data.values };
}
function range(formula) {
  const match = /^(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_ .-]*))!\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$/.exec(formula);
  requireThat(match, 'UNSUPPORTED_FEATURE', 'Chart data must reference one local simple A1 range; formulas and external references are refused.');
  const column = s => [...s].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
  const name = n => { let s = ''; while (n) { s = String.fromCharCode(65 + (n - 1) % 26) + s; n = Math.floor((n - 1) / 26); } return s; };
  const a = column(match[3]), b = column(match[5] ?? match[3]), x = Number(match[4]), y = Number(match[6] ?? match[4]);
  requireThat((a === b || x === y) && b >= a && y >= x && Math.max(b - a, y - x) < 1000 && y <= 1048576 && b <= 16384, 'UNSUPPORTED_FEATURE', 'Chart range must be one bounded row or column.');
  return { sheet: (match[1] ?? match[2]).replaceAll("''", "'"), cells: Array.from({ length: Math.max(b - a, y - x) + 1 }, (_, i) => name(a === b ? a : a + i) + (x === y ? x : x + i)) };
}

/** Update a chart cache and its embedded literal worksheet together. */
export function changeChartData(archive, partName, chartTree, args) {
  requireThat(!descendants(chartTree, 'pivotSource').length, 'UNSUPPORTED_FEATURE', 'Pivot charts require a separate adapter.');
  const series = descendants(chartTree, 'ser');
  const names = series.map(s => {
    const tx = direct(s, 'tx')[0];
    return descendants(tx, 'v').map(n => n.textContent).join('');
  });
  const selected = series.filter((_, i) => names[i] === args.seriesName);
  requireThat(selected.length === 1, 'TARGET_CONFLICT', 'Chart seriesName must match exactly once.');
  const fields = [{ key: 'val', numeric: true, expected: args.expectedValues, values: args.values }];
  if (args.categories !== undefined) fields.push({ key: 'cat', numeric: false, expected: args.expectedCategories, values: args.categories });
  const updates = fields.map(field => ({ ...field, ...binding(selected[0], field.key, field.numeric) }));
  for (const update of updates) {
    requireThat(same(update.expected, update.cachedValues), 'TARGET_CONFLICT', 'Chart data differs from expected values.');
    requireThat(update.values.length === update.pt.length, 'UNSUPPORTED_FEATURE', 'This adapter updates existing points; adding or removing points is not supported.');
  }
  const external = descendants(chartTree, 'externalData');
  const referenced = updates.some(u => u.formula);
  requireThat(!external.length || referenced, 'UNSUPPORTED_FEATURE', 'Literal chart with an external workbook cannot be updated safely.');
  let workbookName, workbookParts;
  if (referenced) {
    const data = one(external, 'Referenced chart needs one embedded workbook.');
    const relBytes = archive.parts[relPart(partName)];
    requireThat(relBytes, 'INVALID_DOCUMENT', 'Chart workbook relationship is missing.');
    const rels = parseXml(Buffer.from(relBytes).toString('utf8'));
    const rel = one(descendants(rels, 'Relationship', 'http://schemas.openxmlformats.org/package/2006/relationships').filter(r => r.getAttribute('Id') === data.getAttributeNS(R, 'id')), 'Chart workbook relationship is ambiguous.');
    requireThat(rel.getAttribute('TargetMode') !== 'External' && /\/package$/.test(rel.getAttribute('Type')), 'UNSUPPORTED_FEATURE', 'External chart workbooks are refused.');
    workbookName = resolvePart(partName, rel.getAttribute('Target'));
    requireThat(/^word\/embeddings\/[^/]+\.xlsx$/.test(workbookName) && archive.parts[workbookName], 'UNSUPPORTED_FEATURE', 'Chart data needs an existing embedded XLSX workbook.');
    for (const name of Object.keys(archive.parts).filter(n => /^word\/charts\/_rels\/.*\.rels$/.test(n) && n !== relPart(partName))) {
      const owner = name.replace('/_rels/', '/').replace(/\.rels$/, '');
      const other = parseXml(Buffer.from(archive.parts[name]).toString('utf8'));
      requireThat(!descendants(other, 'Relationship', 'http://schemas.openxmlformats.org/package/2006/relationships').some(r => r.getAttribute('TargetMode') !== 'External' && resolvePart(owner, r.getAttribute('Target')) === workbookName), 'UNSUPPORTED_FEATURE', 'A workbook shared by multiple charts requires a coordinated update.');
    }
    const bytes = Buffer.from(archive.parts[workbookName]);
    checkArchiveLimits(bytes, Math.min(archive.expandedBytes + 1048576, 64 * 1024 * 1024));
    workbookParts = unzipSync(bytes);
    requireThat(!Object.keys(workbookParts).some(n => /vbaProject|externalLinks|_xmlsignatures|calcChain/i.test(n)), 'UNSUPPORTED_FEATURE', 'Macro, external-link, signed or calculation-chain workbooks are refused.');
    const book = parseXml(Buffer.from(workbookParts['xl/workbook.xml'] ?? []).toString('utf8'));
    const bookRels = parseXml(Buffer.from(workbookParts['xl/_rels/workbook.xml.rels'] ?? []).toString('utf8'));
    const shared = workbookParts['xl/sharedStrings.xml'] ? parseXml(Buffer.from(workbookParts['xl/sharedStrings.xml']).toString('utf8')) : null;
    const sheets = new Map();
    for (const update of updates.filter(u => u.formula)) {
      const location = range(update.formula);
      requireThat(location.cells.length === update.values.length, 'UNSUPPORTED_FEATURE', 'Chart range and cache point counts differ.');
      // An edited range may not feed another series, even through a differently spelled A1 formula.
      for (const other of series) if (other !== selected[0]) for (const formula of descendants(other, 'f')) {
        const otherRange = range(formula.textContent);
        requireThat(otherRange.sheet !== location.sheet || !otherRange.cells.some(c => location.cells.includes(c)), 'UNSUPPORTED_FEATURE', 'Edited data also feeds another series; use a coordinated chart update.');
      }
      const sheet = one(descendants(book, 'sheet', S).filter(s => s.getAttribute('name') === location.sheet), 'Chart worksheet is missing or ambiguous.');
      const link = one(descendants(bookRels, 'Relationship', 'http://schemas.openxmlformats.org/package/2006/relationships').filter(r => r.getAttribute('Id') === sheet.getAttributeNS(R, 'id')), 'Worksheet relationship is missing.');
      requireThat(link.getAttribute('TargetMode') !== 'External', 'UNSUPPORTED_FEATURE', 'External worksheets are refused.');
      const member = resolvePart('xl/workbook.xml', link.getAttribute('Target'));
      requireThat(/^xl\/worksheets\/[^/]+\.xml$/.test(member) && workbookParts[member], 'INVALID_DOCUMENT', 'Worksheet path is invalid.');
      if (!sheets.has(member)) sheets.set(member, parseXml(Buffer.from(workbookParts[member]).toString('utf8')));
      const tree = sheets.get(member);
      requireThat(!descendants(tree, 'f', S).length && !descendants(tree, 'mergeCell', S).length && !descendants(tree, 'sheetProtection', S).length && !descendants(tree, 'extLst', S).length, 'UNSUPPORTED_FEATURE', 'Formula, merged, protected or extension-bearing worksheets are refused.');
      for (const [i, address] of location.cells.entries()) {
        const cell = one(descendants(tree, 'c', S).filter(c => c.getAttribute('r') === address), 'Chart workbook cell is missing or ambiguous.');
        const type = cell.getAttribute('t'), v = direct(cell, 'v', S)[0];
        const old = type === 's' ? descendants(descendants(shared, 'si', S)[Number(v?.textContent)], 't', S).map(n => n.textContent).join('') : type === 'inlineStr' ? descendants(cell, 't', S).map(n => n.textContent).join('') : v?.textContent;
        requireThat(update.numeric ? ['n', ''].includes(type) && Number(old) === update.expected[i] : ['s', 'inlineStr', 'str'].includes(type) && old === update.expected[i], 'TARGET_CONFLICT', 'Embedded workbook and chart cache disagree.');
        for (const child of Array.from(cell.childNodes)) if (child.nodeType === 1 && ['v', 'is'].includes(child.localName)) cell.removeChild(child);
        if (update.numeric) {
          cell.setAttribute('t', 'n'); const node = tree.createElementNS(S, 'v'); node.textContent = String(update.values[i]); cell.appendChild(node);
        } else {
          cell.setAttribute('t', 'inlineStr'); const is = tree.createElementNS(S, 'is'), t = tree.createElementNS(S, 't'); t.textContent = update.values[i]; t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve'); is.appendChild(t); cell.appendChild(is);
        }
      }
    }
    for (const [name, tree] of sheets) workbookParts[name] = Buffer.from(tree.toString());
  }
  for (const update of updates) for (const [i, point] of update.pt.entries()) direct(point, 'v')[0].textContent = String(update.values[i]);
  const packed = workbookName ? Buffer.from(zipSync(Object.fromEntries(Object.entries(workbookParts).map(([name, bytes]) => [name, [bytes, { mtime: new Date('1980-01-01T00:00:00Z') }]])))) : undefined;
  return { parts: workbookName ? new Map([[workbookName, packed]]) : new Map(), before: JSON.stringify({ series: args.seriesName, values: args.expectedValues, ...(args.expectedCategories ? { categories: args.expectedCategories } : {}) }), after: JSON.stringify({ series: args.seriesName, values: args.values, ...(args.categories ? { categories: args.categories } : {}) }) };
}
