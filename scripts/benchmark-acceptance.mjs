// Independent, namespace-aware checks for this synthetic corpus. These do not
// call OfficeMCP's validators or infer success from an engine receipt.
import { createHash } from "node:crypto";
import path from "node:path";
import { DOMParser } from "@xmldom/xmldom";
import { unzipSync, strFromU8 } from "fflate";

const ns = {
  w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  w14: "http://schemas.microsoft.com/office/word/2010/wordml",
  w15: "http://schemas.microsoft.com/office/word/2012/wordml",
  wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  pic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  rels: "http://schemas.openxmlformats.org/package/2006/relationships",
};
const nodes = (element, namespace, local) => Array.from(element.getElementsByTagNameNS(namespace, local));
const children = (element, namespace, local) => Array.from(element.childNodes ?? []).filter((node) => node.namespaceURI === namespace && node.localName === local);
const textOf = (element) => nodes(element, ns.w, "t").map((node) => node.textContent).join("");
const hash = (bytes) => createHash("sha256").update(bytes).digest("hex");
const enabled = (value) => ["1", "true", "on"].includes(value);

function xml(parts, name) {
  if (!parts[name]) throw new Error(`Missing ${name}.`);
  const text = strFromU8(parts[name]);
  if (/<!DOCTYPE|<!ENTITY/i.test(text)) throw new Error("Unexpected XML declaration in benchmark output.");
  let failed = false;
  const document = new DOMParser({ errorHandler: { warning() { failed = true; }, error() { failed = true; }, fatalError() { failed = true; } } }).parseFromString(text, "application/xml");
  if (failed || !document.documentElement) throw new Error(`Invalid XML in ${name}.`);
  return document;
}

function tableData(document) {
  return nodes(document, ns.w, "tbl").map((table) => ({
    node: table,
    rows: children(table, ns.w, "tr").map((row) => children(row, ns.w, "tc").map(textOf)),
  }));
}

function neighboringParagraphText(node, direction) {
  for (let next = node[direction]; next; next = next[direction]) {
    if (next.nodeType !== 1) continue;
    if (next.namespaceURI !== ns.w || next.localName !== "p") return "";
    const text = textOf(next);
    if (text.trim()) return text;
  }
  return "";
}

function tableEvidence(before, after) {
  const original = tableData(before), output = tableData(after);
  const targets = original.flatMap(({ rows }, table) => rows.flatMap((cells, row) =>
    cells[0] === "Service" && cells[1] === "125" ? [{ table, row, column: 1 }] : []));
  const target = targets.length === 1 ? targets[0] : null;
  const sameStructure = original.length === output.length && original.every(({ rows }, i) =>
    rows.length === output[i].rows.length && rows.every((cells, j) => cells.length === output[i].rows[j].length));
  const unchangedExceptTarget = sameStructure && target && original.every(({ rows }, i) =>
    rows.every((cells, j) => cells.every((cell, k) =>
      i === target.table && j === target.row && k === target.column ? true : cell === output[i].rows[j][k])));
  const selected = target && output[target.table];
  return {
    tablePreserved: sameStructure && JSON.stringify(original.map(({ rows }) => rows)) === JSON.stringify(output.map(({ rows }) => rows)),
    tableUpdated: Boolean(unchangedExceptTarget && selected.rows[target.row][target.column] === "150"),
    unrelatedTableCellsPreserved: Boolean(unchangedExceptTarget),
    captionAdded: Boolean(selected && ["previousSibling", "nextSibling"].some((direction) => neighboringParagraphText(selected.node, direction).includes("Service fees"))),
  };
}

function paragraphAncestor(node) {
  for (let parent = node?.parentNode; parent; parent = parent.parentNode)
    if (parent.namespaceURI === ns.w && parent.localName === "p") return parent;
  return null;
}

function commentEvidence(parts, document) {
  if (!parts["word/comments.xml"]) return { existingComment: false, commentAnchoredToReview: false, existingCommentResolved: false };
  const comments = nodes(xml(parts, "word/comments.xml"), ns.w, "comment").filter((node) => textOf(node) === "Approved by legal");
  const comment = comments.length === 1 ? comments[0] : null;
  const id = comment?.getAttributeNS(ns.w, "id");
  const anchors = id == null ? [] : nodes(document, ns.w, "commentRangeStart").filter((node) => node.getAttributeNS(ns.w, "id") === id);
  const ends = id == null ? [] : nodes(document, ns.w, "commentRangeEnd").filter((node) => node.getAttributeNS(ns.w, "id") === id);
  const paragraph = anchors.length === 1 && paragraphAncestor(anchors[0]);
  const paragraphs = comment ? nodes(comment, ns.w, "p") : [];
  const paraId = paragraphs.at(-1)?.getAttributeNS(ns.w14, "paraId");
  const extensions = parts["word/commentsExtended.xml"] ? nodes(xml(parts, "word/commentsExtended.xml"), ns.w15, "commentEx") : [];
  const resolution = paraId ? extensions.filter((node) => node.getAttributeNS(ns.w15, "paraId") === paraId) : [];
  return {
    existingComment: Boolean(comment),
    commentAnchoredToReview: Boolean(paragraph && ends.length === 1 && paragraphAncestor(ends[0]) === paragraph && textOf(paragraph) === "Review clause remains effective."),
    existingCommentResolved: resolution.length === 1 && enabled(resolution[0].getAttributeNS(ns.w15, "done")),
  };
}

function imageData(parts, document) {
  const relations = parts["word/_rels/document.xml.rels"] ? nodes(xml(parts, "word/_rels/document.xml.rels"), ns.rels, "Relationship") : [];
  return nodes(document, ns.w, "drawing").flatMap((drawing) => {
    const blips = nodes(drawing, ns.a, "blip"), properties = nodes(drawing, ns.wp, "docPr"), extents = nodes(drawing, ns.wp, "extent");
    if (blips.length !== 1 || properties.length !== 1 || extents.length !== 1) return [];
    const id = blips[0].getAttributeNS(ns.r, "embed");
    const matching = relations.filter((node) => node.getAttribute("Id") === id && node.getAttribute("TargetMode") !== "External" && node.getAttribute("Type") === ns.r + "/image");
    if (matching.length !== 1) return [];
    const target = matching[0].getAttribute("Target");
    const member = target.startsWith("/") ? target.slice(1) : path.posix.normalize(path.posix.join("word", target));
    if (!parts[member]) return [];
    const shapeExtents = nodes(drawing, ns.pic, "spPr").flatMap((node) => nodes(node, ns.a, "ext"));
    return [{
      hash: hash(parts[member]),
      alt: properties[0].getAttribute("descr") || properties[0].getAttribute("title"),
      width: Number(extents[0].getAttribute("cx")), height: Number(extents[0].getAttribute("cy")),
      shapeExtents: shapeExtents.map((node) => ({ width: Number(node.getAttribute("cx")), height: Number(node.getAttribute("cy")) })),
    }];
  });
}

function imageEvidence(beforeParts, afterParts, before, after) {
  const original = imageData(beforeParts, before), output = imageData(afterParts, after);
  const unchangedImage = original.length === 1 && output.length === 1 && original[0].hash === output[0].hash;
  const mediaHashes = (parts) => Object.entries(parts).filter(([name]) => name.startsWith("word/media/") && !name.endsWith("/")).map(([, bytes]) => hash(bytes)).sort();
  const mediaPreserved = JSON.stringify(mediaHashes(beforeParts)) === JSON.stringify(mediaHashes(afterParts));
  const dimensions = output.length === 1 && [output[0], ...output[0].shapeExtents].every(({ width, height }) => width === 24 * 12700 && height === 24 * 12700);
  return {
    imageStillPresent: Boolean(unchangedImage),
    imageAltUpdated: Boolean(unchangedImage && output[0].alt === "Updated company mark"),
    imageDimensionsUpdated: Boolean(unchangedImage && dimensions),
    imageBytesPreserved: Boolean(unchangedImage && mediaPreserved),
  };
}

function boldDeadline(document) {
  const paragraph = nodes(document, ns.w, "p").find((node) => textOf(node) === "Payment within 45 days.");
  if (!paragraph) return false;
  let offset = 0;
  const start = "Payment within ".length, end = start + 2;
  const covering = nodes(paragraph, ns.w, "r").map((run) => {
    const next = offset + textOf(run).length;
    const overlaps = next > start && offset < end;
    offset = next;
    const bold = nodes(run, ns.w, "b");
    return { overlaps, bold: bold.some((node) => !node.hasAttributeNS(ns.w, "val") || enabled(node.getAttributeNS(ns.w, "val"))) };
  }).filter(({ overlaps }) => overlaps);
  return covering.length > 0 && covering.every(({ bold }) => bold);
}

export function benchmarkVerification({ originalBytes, outputBytes, reportedText = "" }) {
  if (!outputBytes) return { outputExists: false };
  try {
    const beforeParts = unzipSync(originalBytes), afterParts = unzipSync(outputBytes);
    const before = xml(beforeParts, "word/document.xml"), after = xml(afterParts, "word/document.xml");
    const body = textOf(after);
    const footer = Object.keys(afterParts).filter((name) => /^word\/footer\d+\.xml$/.test(name)).map((name) => textOf(xml(afterParts, name))).join("\n");
    return {
      outputExists: true,
      deadlineUpdated: body.includes("Payment within 45 days."),
      otherTextPreserved: body.includes("Deliverables are due on 30 September."),
      ...tableEvidence(before, after),
      footerPreserved: footer.includes("Liability is capped at USD 500."),
      boldNumberPreserved: boldDeadline(after),
      footerReported: reportedText.includes("500"),
      placeholderFilled: body.includes("Proposal for Meridian Labs.") && !body.includes("[insert client name]"),
      proposalContextPreserved: body.includes("Review clause remains effective."),
      footerUpdated: footer.includes("Liability is capped at USD 750."),
      bodyPreserved: textOf(before) === body,
      ...commentEvidence(afterParts, after),
      commentReported: reportedText.includes("Approved by legal") && /open|resolved/i.test(reportedText),
      commentsPreserved: JSON.stringify(commentEvidence(beforeParts, before)) === JSON.stringify(commentEvidence(afterParts, after)),
      reviewTextPreserved: body.includes("Review clause remains effective."),
      ...imageEvidence(beforeParts, afterParts, before, after),
      sentinelReported: reportedText.includes("Sentinel clause 287"),
      largeBodyPreserved: textOf(before) === body,
    };
  } catch (error) { return { outputExists: true, invalidOutput: true, verificationError: String(error.message) }; }
}

export function benchmarkSemanticPass(result, scenario) {
  return result.exit === 0 && !result.timedOut && result.originalUnchanged === true &&
    result.toolCalls > 0 && scenario.expectedSignals.every((key) => result.verification[key] === true);
}
