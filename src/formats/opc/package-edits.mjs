import { readFileSync, writeFileSync } from "node:fs";
import { zipSync, strFromU8, strToU8 } from "fflate";
import { XMLSerializer } from "@xmldom/xmldom";
import {
  unpack,
  parseXml,
  NS,
  child,
  children,
  descendants,
  attrs,
} from "./ooxml.mjs";
import { workbook, presentation, cellPosition } from "./package-readers.mjs";
import { digest } from "../../infrastructure/files.mjs";
import { requireThat } from "../../infrastructure/errors.mjs";
const serializer = new XMLSerializer();
const allowed = (ok, message) =>
  requireThat(ok, "UNSUPPORTED_FEATURE", message);
const guard = (ok) =>
  requireThat(
    ok,
    "TARGET_CONFLICT",
    "Target no longer matches the exact expected content.",
  );
// Associate parsed nodes with source spans: only explicitly changed nodes are serialized.
function editable(parts, part) {
  const source = strFromU8(parts[part]),
    doc = parseXml(source);
  const nodes = [
    doc.documentElement,
    ...Array.from(doc.documentElement.getElementsByTagName("*")),
  ];
  const spans = new Map(),
    stack = [];
  let index = 0;
  const tokens =
    /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<\/?[^<>"']*(?:(?:"[^"]*"|'[^']*')[^<>"']*)*>/g;
  for (const m of source.matchAll(tokens)) {
    if (/^<[!?]/.test(m[0])) continue;
    if (m[0].startsWith("</")) {
      const node = stack.pop();
      guard(node);
      spans.get(node).end = m.index + m[0].length;
    } else {
      const node = nodes[index++];
      guard(node && m[0].match(/^<([^\s/>]+)/)[1] === node.nodeName);
      spans.set(node, { start: m.index, end: m.index + m[0].length });
      if (!/\/\s*>$/.test(m[0])) stack.push(node);
    }
  }
  guard(index === nodes.length && stack.length === 0);
  const patches = [];
  return {
    doc,
    change(node, replacement = serializer.serializeToString(node)) {
      patches.push({ ...spans.get(node), replacement });
    },
    save() {
      let output = source,
        boundary = source.length;
      for (const p of patches.sort((a, b) => b.start - a.start)) {
        guard(p.end <= boundary);
        output = output.slice(0, p.start) + p.replacement + output.slice(p.end);
        boundary = p.start;
      }
      parseXml(output);
      parts[part] = strToU8(output);
    },
  };
}
function element(doc, parent, local, text) {
  const n = doc.createElementNS(
    NS.x,
    parent.prefix ? parent.prefix + ":" + local : local,
  );
  if (text !== undefined) n.appendChild(doc.createTextNode(text));
  return n;
}
function formulaSyntax(formula) {
  // Deliberately small grammar. Every token is consumed; strings, names, links and UDFs are absent.
  const tokens =
    formula.match(
      /[A-Z]+[1-9][0-9]*|[A-Z]+|(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?|[()+\-*/^,:]/g,
    ) ?? [];
  allowed(
    tokens.join("") === formula.replace(/\s/g, ""),
    "Formula contains unsupported syntax. Omit leading =.",
  );
  let i = 0;
  function primary() {
    if (tokens[i] === "+" || tokens[i] === "-") {
      i++;
      return primary();
    }
    const t = tokens[i++];
    if (t === "(") {
      expr();
      allowed(tokens[i++] === ")", "Unbalanced formula.");
      return;
    }
    if (/^[A-Z]+[1-9][0-9]*$/.test(t ?? "")) {
      cellPosition(t);
      if (tokens[i] === ":") {
        i++;
        cellPosition(tokens[i++]);
      }
      return;
    }
    if (/^(?:\d|\.)/.test(t ?? "")) return;
    allowed(
      ["SUM", "MIN", "MAX", "AVERAGE", "COUNT", "ABS", "ROUND"].includes(t),
      "Unsupported formula function.",
    );
    allowed(tokens[i++] === "(", "Function requires arguments.");
    expr();
    while (tokens[i] === ",") {
      i++;
      expr();
    }
    allowed(tokens[i++] === ")", "Unbalanced function.");
  }
  function expr() {
    primary();
    while (["+", "-", "*", "/", "^"].includes(tokens[i])) {
      i++;
      primary();
    }
  }
  expr();
  allowed(i === tokens.length, "Unexpected formula token.");
}
function apply(parts, tool, a) {
  // Parse all XML to reject active declarations even in opaque retained parts.
  for (const [name, b] of Object.entries(parts))
    if (/\.(xml|rels)$/.test(name)) parseXml(strFromU8(b));
  allowed(
    !Object.keys(parts).some((n) =>
      /(vbaProject|_xmlsignatures|activeX|embeddings)/i.test(n),
    ),
    "Signed or active-content packages cannot be edited.",
  );
  if (tool === "pptx_replace_text") {
    const slide = presentation(parts).slides[a.slide - 1];
    guard(slide);
    const e = editable(parts, slide.partName);
    const shapes = descendants(e.doc, NS.p, "sp").filter(
      (s) => descendants(s, NS.p, "cNvPr")[0]?.getAttribute("id") === a.shapeId,
    );
    guard(shapes.length === 1);
    allowed(
      descendants(shapes[0], NS.a, "fld").length === 0,
      "Field-bearing shapes cannot be edited.",
    );
    const targets = descendants(shapes[0], NS.a, "t").filter(
      (t) => t.textContent === a.expectedText,
    );
    guard(targets.length === 1);
    allowed(
      !/[\x00-\x08\x0b\x0c\x0e-\x1f]/.test(a.replacement),
      "Invalid XML text.",
    );
    targets[0].textContent = a.replacement;
    e.change(targets[0]);
    e.save();
    return [slide.partName + "#" + a.shapeId];
  }
  const book = workbook(parts);
  allowed(
    !Object.keys(parts).some((n) =>
      /^xl\/(charts|pivot|externalLinks|tables|connections|calcChain|queryTables)/.test(
        n,
      ),
    ),
    "Workbook dependencies require a broader preservation profile.",
  );
  allowed(
    !descendants(book.root, NS.x, "workbookProtection").length,
    "Protected workbook.",
  );
  const sheet = book.sheets.find((s) => s.sheetId === a.sheetId);
  guard(sheet?.readable);
  const edited = new Map(
    book.sheets.map((s) => {
      allowed(s.readable, "Only worksheet workbooks can be edited.");
      return [s.partName, editable(parts, s.partName)];
    }),
  );
  for (const e of edited.values())
    allowed(
      ![
        "sheetProtection",
        "extLst",
        "conditionalFormatting",
        "dataValidations",
        "drawing",
      ].some((n) => descendants(e.doc, NS.x, n).length),
      "Worksheet features outside this edit profile.",
    );
  const e = edited.get(sheet.partName),
    cs = descendants(e.doc, NS.x, "c"),
    touched = new Set();
  const targets = tool === "xlsx_set_formula" ? [a] : a.cells;
  for (const target of targets) {
    cellPosition(target.address);
    guard(!touched.has(target.address));
    touched.add(target.address);
    const found = cs.filter((c) => c.getAttribute("r") === target.address);
    guard(found.length === 1);
    const c = found[0],
      f = child(c, NS.x, "f"),
      raw = child(c, NS.x, "v")?.textContent ?? null;
    guard(raw === target.expectedRawValue);
    allowed(
      Object.keys(attrs(c)).every(
        (k) =>
          ["r", "s", "t"].includes(k) ||
          k === "xmlns" ||
          k.startsWith("xmlns:"),
      ) &&
        children(c, NS.x).every((n) => ["f", "v", "is"].includes(n.localName)),
      "Cell metadata outside profile.",
    );
    allowed(
      !descendants(c, NS.x, "r").length,
      "Rich text cells cannot be replaced.",
    );
    if (tool === "xlsx_set_formula") {
      guard((f?.textContent ?? null) === a.expectedFormula);
      allowed(
        !f || f.attributes.length === 0,
        "Shared/array/data-table formula targets unsupported.",
      );
      formulaSyntax(a.formula);
    } else allowed(!f, "Use formula operation for formula cells.");
    for (const n of Array.from(c.childNodes)) c.removeChild(n);
    c.removeAttribute("t");
    if (tool === "xlsx_set_formula")
      c.appendChild(element(e.doc, c, "f", a.formula));
    else if (target.type === "string") {
      allowed(
        !/[\x00-\x08\x0b\x0c\x0e-\x1f]/.test(target.value),
        "Invalid XML text.",
      );
      c.setAttribute("t", "inlineStr");
      const is = element(e.doc, c, "is");
      const t = element(e.doc, c, "t", target.value);
      t.setAttribute("xml:space", "preserve");
      is.appendChild(t);
      c.appendChild(is);
    } else if (target.type === "blank") guard(target.value === "");
    else {
      allowed(
        target.type === "boolean"
          ? /^[01]$/.test(target.value)
          : /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(
              target.value,
            ),
        "Invalid lexical cell value.",
      );
      if (target.type === "boolean") c.setAttribute("t", "b");
      c.appendChild(element(e.doc, c, "v", target.value));
    }
    e.change(c);
  }
  for (const [part, edit] of edited) {
    for (const c of descendants(edit.doc, NS.x, "c"))
      if (
        child(c, NS.x, "f") &&
        !(part === sheet.partName && touched.has(c.getAttribute("r")))
      ) {
        const v = child(c, NS.x, "v");
        if (v) {
          c.removeChild(v);
          edit.change(c);
        }
      }
    edit.save();
  }
  const w = editable(parts, "xl/workbook.xml"),
    calc = child(w.doc.documentElement, NS.x, "calcPr");
  allowed(
    calc,
    "Workbook needs existing calculation properties for this profile.",
  );
  calc.setAttribute("fullCalcOnLoad", "1");
  calc.setAttribute("forceFullCalc", "1");
  calc.setAttribute("calcMode", "auto");
  w.change(calc);
  w.save();
  return targets.map((t) => sheet.partName + "#" + t.address);
}
export function editPackage(task, bytes) {
  const parts = unpack(bytes, task.maxExpandedBytes);
  if (task.kind.endsWith("_validate")) {
    const expected = unpack(readFileSync(task.original), task.maxExpandedBytes);
    for (const op of task.operations) apply(expected, op.tool, op);
    requireThat(
      JSON.stringify(Object.keys(parts).sort()) ===
        JSON.stringify(Object.keys(expected).sort()),
      "VALIDATION_FAILED",
      "Package inventory changed.",
    );
    for (const n of Object.keys(expected))
      requireThat(
        digest(parts[n]) === digest(expected[n]),
        "VALIDATION_FAILED",
        "Candidate differs from declared edits or preserved parts.",
      );
    return {
      passed: true,
      revision: digest(bytes),
      checks: [
        "replayed-exact-edit-sequence",
        "all-package-member-bytes",
        "well-formed-XML",
      ],
      notVerified: [
        "Office-layout",
        "formula-calculation",
        "embedded-object-behavior",
      ],
    };
  }
  const affectedTargets = apply(parts, task.kind, task.args),
    output = Buffer.from(
      zipSync(parts, { level: 6, mtime: new Date(1980, 0, 1) }),
    );
  writeFileSync(task.output, output, { flag: "wx", mode: 0o600 });
  return { revision: digest(output), affectedTargets };
}
