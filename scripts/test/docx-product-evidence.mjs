/** Independent saved-file assertions; never imports OfficeMCP validators. */
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { unzipSync, strFromU8 } from 'fflate';
import { DOMParser } from '@xmldom/xmldom';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
export const hash = bytes => createHash('sha256').update(bytes).digest('hex');
export function productIdentity(root) {
  const files = [];
  function walk(relative) {
    const absolute = path.join(root, relative), stat = fs.statSync(absolute);
    if (stat.isDirectory()) for (const name of fs.readdirSync(absolute).sort()) {
      if (name === '.npmignore' || name === '__pycache__' || name.endsWith('.pyc')) continue;
      walk(path.join(relative, name));
    }
    else files.push([relative.replaceAll(path.sep, '/'), hash(fs.readFileSync(absolute))]);
  }
  for (const entry of ['package.json', 'npm-shrinkwrap.json', 'src', 'contracts', '.generated/docx-core']) walk(entry);
  return hash(JSON.stringify(files));
}
export function sdkIdentity(root) {
  root = path.resolve(root);
  const require = createRequire(path.join(root, 'package.json')), entry = require.resolve('@superdoc/sdk');
  let directory = path.dirname(entry), metadata;
  while (directory !== path.dirname(directory)) {
    const file = path.join(directory, 'package.json');
    if (fs.existsSync(file)) { const p = JSON.parse(fs.readFileSync(file)); if (p.name === '@superdoc/sdk') { metadata = p; break; } }
    directory = path.dirname(directory);
  }
  assert(metadata, 'SDK package metadata unavailable');
  const target = (process.platform === 'win32' ? 'windows' : process.platform) + '-' + process.arch;
  const binary = createRequire(entry).resolve(`@superdoc/sdk-${target}/bin/${process.platform === 'win32' ? 'superdoc.exe' : 'superdoc'}`);
  return { version: metadata.version, platform: target, binaryHash: hash(fs.readFileSync(binary)) };
}
const nodes = (e, local, ns = W) => Array.from(e.getElementsByTagNameNS(ns, local));
const text = e => nodes(e, 't').map(n => n.textContent).join('');
function xml(parts, name) {
  assert(parts[name], `Missing ${name}`);
  const source = strFromU8(parts[name]);
  assert(!/<!DOCTYPE|<!ENTITY/i.test(source));
  const errors = [];
  const doc = new DOMParser({ onError: (_level, message) => { errors.push(message); } }).parseFromString(source, 'application/xml');
  assert.equal(errors.length, 0, errors.join('\n'));
  return doc;
}
const enabled = n => !n.hasAttributeNS(W, 'val') || ['1', 'true', 'on'].includes(n.getAttributeNS(W, 'val'));
function paragraph(p) {
  const chars = nodes(p, 'r').flatMap(r => Array.from(text(r), ch => ({
    ch, bold: nodes(r, 'b').some(enabled), italic: nodes(r, 'i').some(enabled),
  })));
  return { text: text(p), chars };
}
export function inspectPackage(bytes) {
  const parts = unzipSync(bytes), doc = xml(parts, 'word/document.xml');
  const body = nodes(doc, 'body')[0];
  const paragraphs = Array.from(body.childNodes).filter(n => n.namespaceURI === W && n.localName === 'p').map(paragraph);
  const tables = nodes(body, 'tbl').map(t => nodes(t, 'tr').map(row =>
    Array.from(row.childNodes).filter(n => n.namespaceURI === W && n.localName === 'tc').map(text)));
  const controls = nodes(body, 'sdt').map(sdt => ({
    tag: nodes(sdt, 'tag')[0]?.getAttributeNS(W, 'val'),
    text: text(sdt), type: nodes(sdt, 'text').length ? 'text' : 'other',
  }));
  const stories = Object.fromEntries(Object.keys(parts).filter(n => /^word\/(header|footer|footnotes|endnotes)\d*\.xml$/.test(n)).sort().flatMap(n => {
    const doc = xml(parts, n);
    // Packer emits reserved separator notes even when there are no user notes;
    // SuperDoc may remove that empty part. Exclude only those reserved objects.
    const paragraphs = /\/(footnotes|endnotes)\.xml$/.test(n)
      ? nodes(doc, n.includes('footnotes') ? 'footnote' : 'endnote')
        .filter(note => !['-1', '0'].includes(note.getAttributeNS(W, 'id')))
        .flatMap(note => nodes(note, 'p').map(paragraph))
      : nodes(doc, 'p').map(paragraph);
    return /\/(footnotes|endnotes)\.xml$/.test(n) && !paragraphs.length ? [] : [[n, paragraphs]];
  }));
  const binary = Object.entries(parts).filter(([n]) => !n.endsWith('/') && !/\.(xml|rels)$/.test(n)).map(([, b]) => hash(b)).sort();
  const relationshipNs = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
  const relations = parts['word/_rels/document.xml.rels'] ? nodes(xml(parts, 'word/_rels/document.xml.rels'), 'Relationship', relationshipNs) : [];
  const allParagraphs = nodes(body, 'p');
  const images = nodes(body, 'drawing').map(drawing => {
    const blip = nodes(drawing, 'blip', A)[0]; assert(blip, 'Missing drawing image reference');
    const relation = relations.find(r => r.getAttribute('Id') === blip.getAttributeNS(R, 'embed'));
    assert(relation && relation.getAttribute('TargetMode') !== 'External');
    const target = relation.getAttribute('Target');
    const member = target.startsWith('/') ? target.slice(1) : path.posix.normalize(path.posix.join('word', target));
    assert(parts[member], 'Missing referenced image bytes');
    let p = drawing.parentNode; while (p && !(p.namespaceURI === W && p.localName === 'p')) p = p.parentNode;
    const extent = nodes(drawing, 'extent', WP)[0], properties = nodes(drawing, 'docPr', WP)[0]; assert(extent && properties);
    return { paragraph: allParagraphs.indexOf(p), hash: hash(parts[member]),
      width: extent.getAttribute('cx'), height: extent.getAttribute('cy'),
      alt: properties.getAttribute('descr'), inline: nodes(drawing, 'inline', WP).length === 1 };
  });
  const comments = parts['word/comments.xml'] ? nodes(xml(parts, 'word/comments.xml'), 'comment').map(c => {
    const id = c.getAttributeNS(W, 'id'), anchors = nodes(doc, 'commentRangeStart').filter(n => n.getAttributeNS(W, 'id') === id);
    const ends = nodes(doc, 'commentRangeEnd').filter(n => n.getAttributeNS(W, 'id') === id);
    const ancestor = n => { while (n && !(n.namespaceURI === W && n.localName === 'p')) n = n.parentNode; return n; };
    const p = ancestor(anchors[0]), last = nodes(c, 'p').at(-1), paraId = last?.getAttributeNS(W14, 'paraId');
    const ext = parts['word/commentsExtended.xml'] ? nodes(xml(parts, 'word/commentsExtended.xml'), 'commentEx', W15).filter(n => n.getAttributeNS(W15, 'paraId') === paraId) : [];
    return { text: text(c), anchor: p && anchors.length === 1 && ends.length === 1 && ancestor(ends[0]) === p ? text(p) : null,
      resolved: ext.length === 1 && ['1', 'true', 'on'].includes(ext[0].getAttributeNS(W15, 'done')) };
  }) : [];
  return { paragraphs, tables, controls, stories, binary, images, comments };
}

export function verifyJob(job, original, output, { reportedText, requireReport = false } = {}) {
  const before = inspectPackage(original), after = inspectPackage(output), expected = structuredClone(before);
  if (job === 'formatted-clause') {
    const p = expected.paragraphs.find(p => p.text === 'Payment within 30 days.'); assert(p);
    p.text = 'Payment within 45 days.';
    const position = 'Payment within '.length;
    p.chars[position].ch = '4'; p.chars[position + 1].ch = '5';
    assert(p.chars[position].bold && p.chars[position + 1].bold);
    if (requireReport) assert.match(reportedText ?? '', /500/, 'Agent must report footer qualification');
  } else if (job === 'table-by-key') {
    const table = expected.tables.find(t => t[0][0] === 'Service' && t[0][1] === 'Price USD'); assert(table);
    assert.equal(table[1][0], 'Support'); assert.equal(table[1][1], '125'); table[1][1] = '150';
  } else if (job === 'template-control') {
    const control = expected.controls.find(c => c.tag === 'client'); assert(control); control.text = 'Meridian Labs';
    const p = expected.paragraphs.find(p => p.text === 'Original client'); assert(p);
    p.text = 'Meridian Labs'; p.chars = Array.from(p.text, ch => ({ ch, bold: false, italic: false }));
  } else assert(['save-recovery', 'comment-read'].includes(job));
  assert.deepEqual(after, expected, 'Only the intended semantic change may differ in the independently inspected package');
  return { passed: true, scope: 'Exact top-level paragraph text/bold/italic, table cells, text-control identity/value, header/footer/note paragraphs, all binary-member hashes, body-image placement/size/alt/reference bytes, comment anchors/resolution. Not full style/layout or arbitrary OOXML fidelity.' };
}
