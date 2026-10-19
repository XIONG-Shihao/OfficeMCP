import { writeFileSync } from "node:fs";
import { digest } from "../../infrastructure/files.mjs";
import { requireThat } from "../../infrastructure/errors.mjs";
import {
  decodePackage,
  xmlPart,
  NS,
  children,
  child,
  descendants,
  attrs,
  relationships,
  relation,
  standardRoot,
  packageInventory,
  mediaType,
} from "./ooxml.mjs";
const X = NS.x,
  P = NS.p,
  A = NS.a;
const omittedX = [
  "calculation",
  "display-formatting",
  "chart-semantics",
  "table-semantics",
  "comments",
  "conditional-formatting",
  "extensions",
];
const omittedP = [
  "rendered-layout",
  "inherited-master-text",
  "chart-semantics",
  "animations",
  "SmartArt",
  "accessibility-semantics",
  "extensions",
];
function hasAncestor(node, ns, name) {
  for (let p = node.parentNode; p; p = p.parentNode)
    if (p.namespaceURI === ns && p.localName === name) return true;
  return false;
}
function stringText(node) {
  return descendants(node, X, "t")
    .filter((t) => !hasAncestor(t, X, "rPh"))
    .map((t) => t.textContent)
    .join("");
}
export function workbook(parts) {
  const root = standardRoot(parts, "xl/workbook.xml", X, "workbook"),
    rels = relationships(parts, "xl/workbook.xml");
  const sheets = children(child(root, X, "sheets"), X, "sheet").map(
    (s, index) => {
      const id = s.getAttribute("sheetId"),
        rid = s.getAttributeNS(NS.r, "id"),
        r = rels.find((r) => r.id === rid);
      requireThat(
        id && r && r.available && !r.external,
        "INVALID_DOCUMENT",
        "Sheet identity or relationship is invalid.",
      );
      return {
        sheetId: id,
        name: s.getAttribute("name"),
        state: s.getAttribute("state") || "visible",
        order: index + 1,
        partName: r.part,
        readable: r.type === NS.r + "/worksheet",
        relationshipType: r.type,
      };
    },
  );
  requireThat(
    new Set(sheets.map((s) => s.sheetId)).size === sheets.length,
    "INVALID_DOCUMENT",
    "Duplicate sheet IDs.",
  );
  return {
    root,
    rels,
    sheets,
    dateSystem: child(root, X, "workbookPr")
      ?.getAttribute("date1904")
      ?.match(/^(1|true)$/)
      ? "1904"
      : "1900",
  };
}
export function cellPosition(ref) {
  const m = /^\$?([A-Z]{1,3})\$?([1-9][0-9]*)$/.exec(ref);
  requireThat(
    m,
    "INVALID_ARGUMENT",
    "Use an A1 cell address without a sheet prefix.",
  );
  let col = 0;
  for (const c of m[1]) col = col * 26 + c.charCodeAt(0) - 64;
  const row = Number(m[2]);
  requireThat(
    col <= 16384 && row <= 1048576,
    "INVALID_ARGUMENT",
    "Cell address exceeds worksheet limits.",
  );
  return { row, col };
}
function address(row, col) {
  let letters = "";
  while (col) {
    col--;
    letters = String.fromCharCode(65 + (col % 26)) + letters;
    col = Math.floor(col / 26);
  }
  return letters + row;
}
function rangeOf(range) {
  const refs = range.split(":");
  requireThat(refs.length <= 2, "INVALID_ARGUMENT", "Use a single A1 range.");
  const start = cellPosition(refs[0]),
    end = cellPosition(refs[1] ?? refs[0]);
  requireThat(
    end.row >= start.row && end.col >= start.col,
    "INVALID_ARGUMENT",
    "Range must be ordered top-left to bottom-right.",
  );
  requireThat(
    (end.row - start.row + 1) * (end.col - start.col + 1) <= 10000,
    "LIMIT_EXCEEDED",
    "Read windows are limited to 10,000 cells.",
  );
  return { start, end };
}
function optionalRelPart(parts, rels, type) {
  const r = rels.find((r) => r.type === NS.r + "/" + type);
  return r ? relation(rels, r.id, type) : null;
}
function readSheet(parts, book, args) {
  const sheet = book.sheets.find((s) => s.sheetId === args.sheetId);
  requireThat(sheet, "TARGET_CONFLICT", "Sheet ID is unavailable.");
  requireThat(
    sheet.readable,
    "UNSUPPORTED_FEATURE",
    "This sheet type is inventoried but cannot be read as worksheet cells.",
  );
  const root = standardRoot(parts, sheet.partName, X, "worksheet");
  const { start, end } = rangeOf(args.range);
  const stringsPart = optionalRelPart(parts, book.rels, "sharedStrings");
  const shared = stringsPart
    ? children(xmlPart(parts, stringsPart).documentElement, X, "si").map(
        stringText,
      )
    : [];
  const stylesPart = optionalRelPart(parts, book.rels, "styles");
  const styles = stylesPart ? xmlPart(parts, stylesPart).documentElement : null;
  const formats = new Map(
    children(child(styles, X, "numFmts"), X, "numFmt").map((n) => [
      n.getAttribute("numFmtId"),
      n.getAttribute("formatCode"),
    ]),
  );
  const xfs = children(child(styles, X, "cellXfs"), X, "xf");
  const hiddenCols = children(child(root, X, "cols"), X, "col")
    .filter((c) => /^(1|true)$/.test(c.getAttribute("hidden")))
    .map((c) => [Number(c.getAttribute("min")), Number(c.getAttribute("max"))]);
  const cells = new Map();
  let inferredRow = 0;
  const hiddenRows = new Set();
  for (const row of children(child(root, X, "sheetData"), X, "row")) {
    const rn = row.hasAttribute("r")
      ? Number(row.getAttribute("r"))
      : inferredRow + 1;
    requireThat(
      Number.isSafeInteger(rn) && rn >= 1 && rn <= 1048576,
      "INVALID_DOCUMENT",
      "Invalid row coordinate.",
    );
    inferredRow = rn;
    if (/^(1|true)$/.test(row.getAttribute("hidden"))) hiddenRows.add(rn);
    let inferredCol = 0;
    for (const c of children(row, X, "c")) {
      const pos = c.hasAttribute("r")
        ? cellPosition(c.getAttribute("r"))
        : { row: rn, col: inferredCol + 1 };
      requireThat(
        pos.row === rn && pos.col <= 16384,
        "INVALID_DOCUMENT",
        "Cell coordinate disagrees with its row.",
      );
      inferredCol = pos.col;
      const ref = address(pos.row, pos.col);
      requireThat(
        !cells.has(ref),
        "INVALID_DOCUMENT",
        "Duplicate cell address.",
      );
      cells.set(ref, c);
    }
  }
  const items = [];
  for (let row = start.row; row <= end.row; row++)
    for (let col = start.col; col <= end.col; col++) {
      const ref = address(row, col),
        c = cells.get(ref);
      const base = {
        address: ref,
        rowHidden: hiddenRows.has(row),
        columnHidden: hiddenCols.some(([a, b]) => col >= a && col <= b),
      };
      if (!c) {
        items.push({
          ...base,
          type: "blank",
          value: null,
          rawValue: null,
          formula: null,
          cachedValue: null,
          numberFormat: null,
          storage: "absent",
        });
        continue;
      }
      const raw = child(c, X, "v")?.textContent ?? null,
        t = c.getAttribute("t") || "n",
        f = child(c, X, "f");
      let value = raw,
        type = t;
      const missingCache = raw === null || (raw === "" && t !== "str");
      if (f && missingCache) {
        value = null;
        type =
          {
            n: "number",
            b: "boolean",
            str: "string",
            e: "error",
            d: "iso-date",
          }[t] ?? t;
      } else if (t === "s") {
        requireThat(
          raw !== null && /^\d+$/.test(raw) && Number(raw) < shared.length,
          "INVALID_DOCUMENT",
          "Shared string index is invalid.",
        );
        value = shared[Number(raw)];
        type = "string";
      } else if (t === "inlineStr") {
        value = stringText(child(c, X, "is"));
        type = "string";
      } else if (t === "b") {
        requireThat(
          raw === "0" || raw === "1",
          "INVALID_DOCUMENT",
          "Invalid boolean cell.",
        );
        value = raw === "1";
        type = "boolean";
      } else if (t === "n") {
        type = raw === null || raw === "" ? "blank" : "number";
        value = type === "blank" ? null : raw;
      } else if (t === "str") type = "string";
      else if (t === "e") type = "error";
      else if (t === "d") type = "iso-date";
      else type = "unsupported:" + t;
      const styleIndex = c.hasAttribute("s") ? Number(c.getAttribute("s")) : 0;
      requireThat(
        Number.isInteger(styleIndex) &&
          styleIndex >= 0 &&
          (!c.hasAttribute("s") || styleIndex < xfs.length),
        "INVALID_DOCUMENT",
        "Cell style index is unavailable.",
      );
      const formatId = xfs[styleIndex]?.getAttribute("numFmtId") || "0";
      const formula = f
        ? {
            expression: f.textContent || null,
            attributes: attrs(f),
            interpretation: f.textContent
              ? "source-expression"
              : "shared-or-special-formula-unresolved",
          }
        : null;
      items.push({
        ...base,
        type,
        value: f ? null : value,
        rawValue: raw,
        formula,
        cachedValue: f
          ? {
              value,
              raw,
              type,
              freshness: missingCache ? "missing" : "unverified",
            }
          : null,
        numberFormat: {
          id: formatId,
          code: formats.get(formatId) ?? null,
          source: formats.has(formatId) ? "custom" : "builtin-id",
        },
        displayedText: null,
        storage: "present",
      });
    }
  return {
    items: args.query
      ? items.filter((c) =>
          [c.value, c.rawValue, c.formula?.expression, c.cachedValue?.value]
            .filter((v) => v !== null && v !== undefined)
            .some((v) => String(v).includes(args.query)),
        )
      : items,
    domain: "worksheet-cells",
    sheet,
    range: args.range,
    dateSystem: book.dateSystem,
    calculation: { performed: false, engine: null, freshness: "unverified" },
    dimensionHint: child(root, X, "dimension")?.getAttribute("ref") ?? null,
    mergedRanges: children(child(root, X, "mergeCells"), X, "mergeCell").map(
      (n) => n.getAttribute("ref"),
    ),
    omissions: omittedX,
  };
}
export function presentation(parts) {
  const root = standardRoot(parts, "ppt/presentation.xml", P, "presentation"),
    rels = relationships(parts, "ppt/presentation.xml");
  const slides = children(child(root, P, "sldIdLst"), P, "sldId").map(
    (n, i) => {
      const partName = relation(rels, n.getAttributeNS(NS.r, "id"), "slide");
      const slide = standardRoot(parts, partName, P, "sld");
      return {
        slide: i + 1,
        slideId: n.getAttribute("id"),
        partName,
        hidden: /^(0|false)$/.test(slide.getAttribute("show")),
        relationships: relationships(parts, partName),
      };
    },
  );
  return { slides, size: attrs(child(root, P, "sldSz")) };
}
function textParagraphs(node) {
  return descendants(node, A, "p").map((p) =>
    children(p, A)
      .map((n) =>
        n.localName === "br"
          ? "\n"
          : descendants(n, A, "t")
              .map((t) => t.textContent)
              .join(""),
      )
      .join(""),
  );
}
function slideShapes(parts, partName) {
  const root = xmlPart(parts, partName).documentElement,
    rels = relationships(parts, partName),
    tree = child(child(root, P, "cSld"), P, "spTree");
  requireThat(tree, "INVALID_DOCUMENT", "Slide shape tree is missing.");
  const items = [];
  function visit(parent, groupPath = []) {
    for (const node of Array.from(parent.childNodes).filter(
      (n) => n.nodeType === 1,
    )) {
      if (node.namespaceURI !== P) {
        items.push({
          shapeId: null,
          type: node.nodeName,
          groupPath,
          text: null,
          interpretation: "unsupported-object",
          namespace: node.namespaceURI,
        });
        continue;
      }
      if (["nvGrpSpPr", "grpSpPr", "extLst"].includes(node.localName)) continue;
      const nv = descendants(node, P, "cNvPr")[0];
      const shapeId = nv?.getAttribute("id") ?? null;
      const paras = textParagraphs(node);
      const type = node.localName;
      const xfrm =
        descendants(node, A, "xfrm")[0] ?? descendants(node, P, "xfrm")[0];
      const references = rels.filter((r) =>
        Array.from(node.getElementsByTagName("*")).some((n) =>
          Array.from(n.attributes ?? []).some(
            (a) => a.namespaceURI === NS.r && a.value === r.id,
          ),
        ),
      );
      const table = type === "grpSp" ? null : descendants(node, A, "tbl")[0];
      const item = {
        shapeId,
        name: nv?.getAttribute("name") ?? null,
        type,
        groupPath,
        text: type === "grpSp" ? "" : paras.join("\n"),
        paragraphs: type === "grpSp" ? [] : paras,
        geometry: {
          units: "EMU",
          coordinateSpace: groupPath.length ? "parent-group" : "slide",
          transform: attrs(xfrm),
          offset: attrs(child(xfrm, A, "off")),
          extent: attrs(child(xfrm, A, "ext")),
          childOffset: attrs(child(xfrm, A, "chOff")),
          childExtent: attrs(child(xfrm, A, "chExt")),
        },
        placeholder: attrs(descendants(node, P, "ph")[0]),
        references,
      };
      if (table)
        item.table = children(table, A, "tr").map((r, row) =>
          children(r, A, "tc").map((c, col) => ({
            row,
            col,
            text: textParagraphs(c).join("\n"),
            attributes: attrs(c),
          })),
        );
      if (!["sp", "pic", "graphicFrame", "cxnSp", "grpSp"].includes(type))
        item.interpretation = "unsupported-object";
      items.push(item);
      if (type === "grpSp") visit(node, [...groupPath, shapeId]);
    }
  }
  visit(tree);
  return items;
}
export function readPackage(task, bytes) {
  const { parts } = decodePackage(bytes, task.maxExpandedBytes);
  const { kind, args } = task;
  if (kind === "office_package_parts")
    return {
      items: packageInventory(parts),
      domain: "package-parts",
      omissions: ["semantic-interpretation"],
    };
  if (kind === "office_package_extract") {
    requireThat(
      Object.hasOwn(parts, args.partName),
      "NOT_FOUND",
      "Requested package part is unavailable.",
    );
    const part = parts[args.partName];
    requireThat(
      part.length <= 20 * 1024 * 1024,
      "LIMIT_EXCEEDED",
      "Extracted part exceeds 20 MiB.",
    );
    writeFileSync(task.output, part, { flag: "wx", mode: 0o600 });
    return {
      partName: args.partName,
      mediaType: mediaType(args.partName),
      size: part.length,
      hash: digest(part),
      evidence: "exact-package-bytes",
    };
  }
  if (kind.startsWith("xlsx")) {
    const book = workbook(parts);
    if (kind === "xlsx_inspect") {
      const names = children(
        child(book.root, X, "definedNames"),
        X,
        "definedName",
      ).map((n) => ({
        name: n.getAttribute("name"),
        expression: n.textContent,
        attributes: attrs(n),
      }));
      return {
        items: args.collection === "definedNames" ? names : book.sheets,
        domain: args.collection === "definedNames" ? "defined-names" : "sheets",
        sheetCount: book.sheets.length,
        dateSystem: book.dateSystem,
        calculationProperties: attrs(child(book.root, X, "calcPr")),
        omissions: omittedX,
      };
    }
    return readSheet(parts, book, args);
  }
  const deck = presentation(parts);
  if (kind === "pptx_inspect")
    return {
      items: deck.slides,
      domain: "slides",
      slideCount: deck.slides.length,
      size: deck.size,
      omissions: omittedP,
    };
  const slide = deck.slides[args.slide - 1];
  requireThat(slide, "TARGET_CONFLICT", "Slide number is unavailable.");
  let partName = slide.partName;
  if (args.domain === "notes") {
    const rel = slide.relationships.find(
      (r) => r.type === NS.r + "/notesSlide",
    );
    if (!rel)
      return {
        items: [],
        domain: "notes",
        slide: args.slide,
        notesState: "absent",
        omissions: omittedP,
      };
    partName = relation(slide.relationships, rel.id, "notesSlide");
  }
  const items = slideShapes(parts, partName);
  return {
    items: args.query
      ? items.filter((i) => i.text?.includes(args.query))
      : items,
    domain: args.domain,
    slide: args.slide,
    hidden: slide.hidden,
    partName,
    readingOrder: "source-shape-tree; not visual reading order",
    omissions: omittedP,
  };
}
