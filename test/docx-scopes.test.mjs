import test from "./support/test.mjs";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { Document, Packer, Paragraph, Header, Table, TableRow, TableCell, HeadingLevel } from "docx";
import { createServiceFixture } from "./service-fixture.mjs";
import { queryAgentDocument, resolveInlineLocator } from "../src/formats/docx/docx-agent-query.mjs";

const table = (value) => new Table({ rows: [new TableRow({ children: [new TableCell({ children: [new Paragraph("Shared header")] }), new TableCell({ children: [new Paragraph(value)] })] })] });

test("structural selectors constrain body containers and reject unsupported header inventories", async () => {
  const s = createServiceFixture({ toolset: "full" });
  try {
    const file = path.join(s.docs, "scoped.docx");
    const container = new Table({ rows: [new TableRow({ children: [new TableCell({ children: [
      new Paragraph("Container header"),
      new Paragraph({ text: "Shared heading", heading: HeadingLevel.HEADING_1 }), table("Nested value"),
    ] })] })] });
    fs.writeFileSync(file, await Packer.toBuffer(new Document({ styles: { paragraphStyles: [{ id: "Heading1", name: "heading 1", paragraph: { outlineLevel: 0 } }] }, sections: [{
      headers: { default: new Header({ children: [new Paragraph({ text: "Shared heading", heading: HeadingLevel.HEADING_1 }), table("Header value")] }) },
      children: [new Paragraph({ text: "Shared heading", heading: HeadingLevel.HEADING_1 }), table("Body value"), container],
    }] })));
    let view = (await s.call("docx_open", { path: file })).view;
    const section = (await s.call("docx_read", { view, representation: "domain", domain: "sections" })).data.items[0];
    const story = { type: "header", sectionId: section.itemId, variant: "default" };
    for (const select of [{ type: "table", header: "Shared header" }, { type: "headingPath", path: ["Shared heading"] }]) {
      const rejected = await s.service.call("docx_query", { view, story, select, require: "exactlyOne" });
      assert.equal(rejected.payload.error.code, "UNSUPPORTED_FEATURE");
      const inline = await s.service.call("docx_apply", { view, operationId: `unsupported-${select.type}`, select, story, action: "write_table_cells", arguments: { cells: [{ row: 0, column: 0, expectedText: "Shared header", text: "Wrong story" }] } });
      assert.equal(inline.payload.error.code, "UNSUPPORTED_FEATURE");
    }
    // The SDK's supported scoped text route still sees the header, not the body.
    const headerText = await s.call("docx_query", { view, story, select: { type: "text", pattern: "Header value" }, require: "exactlyOne" });
    assert.equal(headerText.data.items[0].text, "Header value");
    const outer = await s.call("docx_query", { view, select: { type: "table", header: "Container header" }, require: "exactlyOne" });
    const withinTarget = outer.data.items[0].target;
    const nested = await s.call("docx_query", { view, withinTarget, select: { type: "table", header: "Shared header" }, require: "exactlyOne" });
    assert.match(nested.data.items[0].text, /Nested value/);
    assert.doesNotMatch(nested.data.items[0].text, /Body value|Header value/);
    const heading = await s.call("docx_query", { view, withinTarget, select: { type: "headingPath", path: ["Shared heading"] }, require: "exactlyOne" });
    assert.equal(heading.data.items.length, 1);
    const nestedText = await s.call("docx_query", { view, withinTarget: nested.data.items[0].target, select: { type: "text", pattern: "Nested value" }, require: "exactlyOne" });
    view = (await s.call("docx_apply", { view, operationId: "scoped-write", action: "replace_text", selectorTarget: nestedText.data.items[0].target, arguments: { edits: [{ find: "Nested value", replace: "Updated nested value" }], caseSensitive: true } })).view;
    const body = await s.call("docx_query", { view, select: { type: "table", contains: "Body value" }, require: "exactlyOne" });
    assert.match(body.data.items[0].text, /Body value/);
    const headerAfter = await s.call("docx_query", { view, story, select: { type: "text", pattern: "Header value" }, require: "exactlyOne" });
    assert.equal(headerAfter.data.items[0].text, "Header value");
    await s.call("docx_save", { view, operationId: "scoped-save", destination: { rootId: "output", relativePath: "scoped-output.docx" } });
  } finally { s.close(); }
});

test("unsupported entity scopes fail before consulting a document-wide inventory", async () => {
  const select = { type: "entity", domain: "comments", author: "Editor" };
  const noInventory = { comments: { list: () => assert.fail("scoped query widened to global inventory") } };
  for (const scope of [{ story: { type: "body" } }, { story: { type: "header", sectionId: "section" } }, { withinLocator: { address: { nodeId: "table" } } }]) {
    await assert.rejects(queryAgentDocument(noInventory, { select, ...scope }), { code: "UNSUPPORTED_FEATURE" });
    await assert.rejects(resolveInlineLocator(noInventory, { action: "reply_to_comment", select, ...scope }), { code: "UNSUPPORTED_FEATURE" });
  }
});
