import test from "./support/test.mjs";
import assert from "node:assert/strict";
import { strToU8, strFromU8, zipSync } from "fflate";
import { decodePackage, xmlPart, relationships, NS } from "../src/formats/opc/ooxml.mjs";
import { readPackage } from "../src/formats/opc/package-readers.mjs";
import { readDocument } from "../src/formats/common/document-read.mjs";
import { preserveWorkbookEmbeddings } from "../src/formats/opc/preserve-embeddings.mjs";
import { packageEvidence } from "../src/formats/opc/engine-evidence.mjs";
import { digest } from "../src/infrastructure/files.mjs";
import { xlsxBytes, pptxBytes } from "./four-format-fixtures.mjs";
const limit = 20 * 1024 * 1024;

test("common worksheet and slide reads reuse decoded members and XML while preserving read windows", () => {
  for (const [format, bytes, detail] of [
    ["xlsx", xlsxBytes(), { sheet: "Budget", range: "A1:E4" }],
    ["pptx", pptxBytes(), { slide: 1 }],
  ]) {
    const task = { kind: "office_read", format, maxExpandedBytes: limit,
      args: { scope: "content", offset: 1, limit: 2, ...detail } };
    const context = decodePackage(bytes, limit);
    const result = readDocument(task, context);
    assert.deepEqual(result, readDocument(task, bytes));
    assert.equal(result.items.length, 2);
    assert.equal(result.offset, 1);
    assert.equal(result.nextOffset, 3);
    const xmls = new Map(context.xml);
    assert(xmls.has(format === "xlsx" ? "xl/workbook.xml" : "ppt/presentation.xml"));
    assert.deepEqual(readDocument(task, context), result);
    for (const [name, cached] of xmls) assert.equal(context.xml.get(name).document, cached.document);
    assert.equal(decodePackage(context, limit), context);
    assert.throws(() => decodePackage(context, 1), { code: "LIMIT_EXCEEDED" });
  }
});

test("package context invalidates replaced XML and relationships and rechecks target availability", () => {
  const context = decodePackage(xlsxBytes(), limit), { parts } = context;
  const original = xmlPart(parts, "xl/workbook.xml");
  parts["xl/workbook.xml"] = strToU8(strFromU8(parts["xl/workbook.xml"]).replace('name="Budget"', 'name="Forecast"'));
  assert.notEqual(xmlPart(parts, "xl/workbook.xml"), original);
  const result = readPackage({ kind: "xlsx_inspect", maxExpandedBytes: limit, args: { collection: "sheets" } }, context);
  assert.equal(result.items[0].name, "Forecast");
  const first = relationships(parts, "xl/workbook.xml");
  assert(first.find((r) => r.id === "budget").available);
  delete parts["xl/worksheets/budget.xml"];
  assert.equal(relationships(parts, "xl/workbook.xml").find((r) => r.id === "budget").available, false);
  parts["xl/_rels/workbook.xml.rels"] = strToU8(strFromU8(parts["xl/_rels/workbook.xml.rels"]).replace('Id="budget"', 'Id="renamed"'));
  assert(relationships(parts, "xl/workbook.xml").some((r) => r.id === "renamed"));
  assert(!relationships(parts, "xl/workbook.xml").some((r) => r.id === "budget"));
  parts["xl/workbook.xml"] = strToU8('<!DOCTYPE workbook><workbook/>');
  assert.throws(() => xmlPart(parts, "xl/workbook.xml"), { code: "INVALID_DOCUMENT" });
});

test("embedding restoration and engine evidence share package contexts without losing original binary bytes", () => {
  const parts = decodePackage(xlsxBytes(), limit).parts;
  parts["[Content_Types].xml"] = strToU8(strFromU8(parts["[Content_Types].xml"]).replace('</Types>', '<Default Extension="bin" ContentType="application/octet-stream"/></Types>'));
  parts["xl/worksheets/budget.xml"] = strToU8(strFromU8(parts["xl/worksheets/budget.xml"]).replace('</worksheet>', `<oleObjects><oleObject xmlns:r="${NS.r}" r:id="ole" shapeId="9" progId="Package"/></oleObjects></worksheet>`));
  parts["xl/worksheets/_rels/budget.xml.rels"] = strToU8(strFromU8(parts["xl/worksheets/_rels/budget.xml.rels"]).replace('</Relationships>', `<Relationship Id="ole" Type="${NS.r}/oleObject" Target="../embeddings/embedded.bin"/></Relationships>`));
  parts["xl/embeddings/embedded.bin"] = Buffer.from("original embedded data");
  const before = decodePackage(Buffer.from(zipSync(parts)), limit);
  parts["xl/embeddings/embedded.bin"] = Buffer.from("engine rewritten container");
  const after = decodePackage(Buffer.from(zipSync(parts)), limit);
  assert.throws(() => packageEvidence(before, after, limit), { code: "PRESERVATION_FAILED" });
  const restored = preserveWorkbookEmbeddings(before, after, limit);
  assert.equal(restored.packageContext, after);
  assert.equal(restored.restoredEmbeddings, 1);
  assert.deepEqual(after.parts["xl/embeddings/embedded.bin"], before.parts["xl/embeddings/embedded.bin"]);
  const evidence = packageEvidence(before, restored.packageContext, limit);
  assert.equal(evidence.outputRevision, digest(restored.bytes));
  assert.equal(evidence.binaryMembersPreserved, 1);
  assert.deepEqual(evidence, packageEvidence(before.bytes, restored.bytes, limit));
});
