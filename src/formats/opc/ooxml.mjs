import { unzipSync, strFromU8 } from "fflate";
import { DOMParser } from "@xmldom/xmldom";
import { SaxesParser } from "saxes";
import path from "node:path";
import { OfficeError, requireThat } from "../../infrastructure/errors.mjs";
import { checkArchiveLimits } from "./archive-limits.mjs";
import {nativeUnpack} from './native-unpack.mjs';
import {openFilePackage} from './file-package.mjs';

// A context lives only as long as its decoded package. Reusing it within one
// operation avoids ZIP expansion and XML parsing on every inspection pass.
// Edits replace member buffers; cached XML/relationships invalidate on replacement.
const packageContexts = new WeakMap();
class DecodedPackage {
  constructor(bytes, limit, {nativeInflate=false}={}) {
    this.bytes = bytes;
    this.parts = nativeInflate ? nativeUnpack(bytes,limit) : unpack(bytes, limit);
    this.xml = new Map();
    this.relationships = new Map();
    packageContexts.set(this.parts, this);
  }
}
export function decodePackage(source, limit, options) {
  if (!(source instanceof DecodedPackage)) return new DecodedPackage(source, limit, options);
  // Restoration may replace an engine's small OLE wrapper with larger original
  // bytes. Recheck the current members without expanding their ZIP a second time.
  const expandedBytes = source.filePackage?.expandedBytes ?? Object.values(source.parts).reduce((total, part) => total + part.length, 0);
  requireThat(expandedBytes <= limit, "LIMIT_EXCEEDED", "ZIP expansion exceeds limit.");
  return source;
}
export function decodeFilePackage(file,limit,maxBytes) {
  const context=Object.create(DecodedPackage.prototype);
  context.filePackage=openFilePackage(file,limit,maxBytes);
  context.parts=context.filePackage.parts;context.xml=new Map();context.relationships=new Map();
  context.close=()=>{context.filePackage.close();context.xml.clear();context.relationships.clear();};
  packageContexts.set(context.parts,context);
  return context;
}

export function unpack(bytes, maxExpandedBytes) {
  const { entries } = checkArchiveLimits(bytes, maxExpandedBytes);
  let parts;
  try {
    parts = unzipSync(bytes);
  } catch {
    throw new OfficeError("INVALID_DOCUMENT", "Could not decode ZIP members.");
  }
  requireThat(
    Object.keys(parts).length === entries &&
      Object.values(parts).reduce((n, b) => n + b.length, 0) <=
        maxExpandedBytes,
    "INVALID_DOCUMENT",
    "ZIP inventory mismatch.",
  );
  requireThat(
    parts["[Content_Types].xml"],
    "UNSUPPORTED_FORMAT",
    "OOXML content types are missing.",
  );
  return parts;
}
export function parseXml(xml) {
  requireThat(
    !/<!DOCTYPE|<!ENTITY/i.test(xml),
    "INVALID_DOCUMENT",
    "DTD and entity declarations are disabled.",
  );
  let bad = false;
  const doc = new DOMParser({
    errorHandler: {
      warning() {
        bad = true;
      },
      error() {
        bad = true;
      },
      fatalError() {
        bad = true;
      },
    },
  }).parseFromString(xml, "application/xml");
  requireThat(
    !bad && doc.documentElement,
    "INVALID_DOCUMENT",
    "Malformed XML is unsupported.",
  );
  return doc;
}
/** Syntax/namespace checks without allocating a document tree. Metadata reads
 * still use parseXml; a well-formedness-only pass retains no XML nodes. */
export function validateXml(bytes) {
  const fail=()=>{throw new OfficeError("INVALID_DOCUMENT", "Malformed XML is unsupported.");};
  const parser=new SaxesParser({xmlns:true});
  parser.on('error',fail);
  parser.on('doctype',()=>{throw new OfficeError("INVALID_DOCUMENT", "DTD and entity declarations are disabled.");});
  const decoder=new TextDecoder('utf-8',{fatal:true});
  let carry='';
  const feed=text=>{
    requireThat(!/<!DOCTYPE|<!ENTITY/i.test(carry+text),"INVALID_DOCUMENT","DTD and entity declarations are disabled.");
    parser.write(text);carry=text.slice(-16);
  };
  try {
    for(let offset=0;offset<bytes.length;offset+=64*1024)feed(decoder.decode(bytes.subarray(offset,offset+64*1024),{stream:true}));
    feed(decoder.decode());parser.close();
  }catch(error){if(error instanceof OfficeError)throw error;fail();}
}
export const NS = {
  x: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
  p: "http://schemas.openxmlformats.org/presentationml/2006/main",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  rel: "http://schemas.openxmlformats.org/package/2006/relationships",
};
export const children = (node, ns, name) =>
  Array.from(node?.childNodes ?? []).filter(
    (n) =>
      n.nodeType === 1 &&
      n.namespaceURI === ns &&
      (!name || n.localName === name),
  );
export const child = (node, ns, name) => children(node, ns, name)[0] ?? null;
export const descendants = (node, ns, name) =>
  node ? Array.from(node.getElementsByTagNameNS(ns, name)) : [];
export const attrs = (node) =>
  Object.fromEntries(
    Array.from(node?.attributes ?? []).map((a) => [a.name, a.value]),
  );
export function xmlPart(parts, name) {
  requireThat(
    parts[name],
    "INVALID_DOCUMENT",
    "A required package part is missing.",
  );
  const cache = packageContexts.get(parts)?.xml;
  const cached = cache?.get(name);
  if (cached?.bytes === parts[name]) return cached.document;
  const document = parseXml(strFromU8(parts[name]));
  cache?.set(name, { bytes: parts[name], document });
  return document;
}
export function relationships(parts, source) {
  const relName = path.posix.join(
    path.posix.dirname(source),
    "_rels",
    path.posix.basename(source) + ".rels",
  );
  if (!parts[relName]) return [];
  const cache = packageContexts.get(parts)?.relationships;
  const cached = cache?.get(source);
  if (cached?.bytes === parts[relName])
    return cached.items.map((item) => ({
      ...item,
      available: !item.external && Object.hasOwn(parts, item.part),
    }));
  const ids = new Set();
  const items = children(
    xmlPart(parts, relName).documentElement,
    NS.rel,
    "Relationship",
  ).map((n) => {
    const id = n.getAttribute("Id"),
      type = n.getAttribute("Type"),
      target = n.getAttribute("Target"),
      external = n.getAttribute("TargetMode") === "External";
    requireThat(
      id && type && target && !ids.has(id),
      "INVALID_DOCUMENT",
      "Missing or duplicate relationship identity.",
    );
    ids.add(id);
    let part = null;
    if (!external) {
      let decoded;
      try {
        decoded = decodeURIComponent(target);
      } catch {
        throw new OfficeError(
          "INVALID_DOCUMENT",
          "Malformed relationship target.",
        );
      }
      requireThat(
        !decoded.includes("\\") &&
          !/[\x00-\x1f?#]/.test(decoded) &&
          !/^\/\/|^[a-z][a-z0-9+.-]*:/i.test(decoded),
        "INVALID_DOCUMENT",
        "Unsafe internal relationship target.",
      );
      part = path.posix.normalize(
        decoded.startsWith("/")
          ? decoded.slice(1)
          : path.posix.join(path.posix.dirname(source), decoded),
      );
      requireThat(
        part !== ".." && !part.startsWith("../"),
        "INVALID_DOCUMENT",
        "Relationship escapes the package.",
      );
    }
    return {
      id,
      type,
      target,
      external,
      part,
      available: !external && Object.hasOwn(parts, part),
    };
  });
  cache?.set(source, { bytes: parts[relName], items: items.map((item) => ({ ...item })) });
  return items;
}
export function relation(rels, id, suffix) {
  const r = rels.find((r) => r.id === id);
  requireThat(
    r &&
      !r.external &&
      r.available &&
      (!suffix || r.type === NS.r + "/" + suffix),
    "INVALID_DOCUMENT",
    "Missing, external or incompatible required relationship.",
  );
  return r.part;
}
export function standardRoot(parts, name, ns, local) {
  const root = xmlPart(parts, name).documentElement;
  requireThat(
    root.namespaceURI === ns && root.localName === local,
    "UNSUPPORTED_FEATURE",
    "Only Transitional OOXML namespaces are supported in this preview.",
  );
  return root;
}
export function packageInventory(parts) {
  return Object.keys(parts)
    .sort()
    .map((name) => ({
      partName: name,
      size: parts[name].length,
      mediaType: mediaType(name),
    }));
}
export function mediaType(name) {
  const ext = path.posix.extname(name).toLowerCase();
  return (
    {
      ".xml": "application/xml",
      ".rels": "application/xml",
      ".png": "image/png",
      ".jpg": "image/jpeg",
      ".jpeg": "image/jpeg",
      ".svg": "image/svg+xml",
    }[ext] ?? "application/octet-stream"
  );
}
