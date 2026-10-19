import test from "./support/test.mjs";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { createServiceFixture } from "./service-fixture.mjs";

const cases = [
  ["delete_blocks", {}, "paragraph"],
  ["append_list", { items: ["One", "Two"], kind: "bullet" }],
  ["normalize_body_font_size", { fontSize: 11 }],
  ["set_font_family", { fontFamily: "Arial" }, "paragraph"],
  ["apply_letter_spacing", { letterSpacing: 1 }, "paragraph"],
  ["insert_table_row", { rowIndex: 0, position: "after", cellTexts: ["C", "D"] }, "table"],
  ["insert_table_column", { columnIndex: 0, position: "right", headerText: "New" }, "table"],
  ["delete_table_row", { rowIndex: 1 }, "table"],
  ["delete_table_column", { columnIndex: 1 }, "table"],
  ["convert_list", { kind: "bullet", fromText: "Body", toText: "Body" }],
  ["apply_style", { headingLevel: 2 }, "paragraph"],
  ["delete_table", {}, "table"],
  ["add_hyperlink", { text: "Body", url: "https://example.com" }],
];

function setup() {
  return createServiceFixture({
    prefix: "docx-qualification-",
    toolset: "docx-agent",
    limits: { workerTimeoutMs: 30000 },
  });
}

for (const [action, argumentsValue, targetKind] of cases)
  test(`qualified action evidence: ${action}`, async () => {
    const fixture = setup();
    try {
      const created = await fixture.call("docx_create", {
        operationId: `create-${action}`,
        document: {
          format: "docx",
          footer: "Preserve footer",
          blocks: [
            { kind: "paragraph", text: "Body" },
            { kind: "table", rows: [["A", "B"], ["1", "2"]] },
          ],
        },
      });
      const view = created.view;
      let selectorTarget;
      if (targetKind === "paragraph")
        selectorTarget = (
          await fixture.call("docx_query", {
            view,
            select: { type: "text", pattern: "Body" },
            require: "exactlyOne",
            offset: 0,
            limit: 2,
          })
        ).data.items[0].target;
      if (targetKind === "table")
        selectorTarget = (
          await fixture.call("docx_read", {
            view,
            representation: "domain",
            domain: "tables",
            offset: 0,
            limit: 1,
          })
        ).data.items[0].target;

      const before = await fixture.call("docx_read", {
        view,
        representation: "text",
        offset: 0,
        limit: 20,
      });
      const preview = await fixture.call("docx_apply", {
        view,
        mode: "preview",
        action,
        arguments: argumentsValue,
        ...(selectorTarget ? { selectorTarget } : {}),
      });
      assert.equal(preview.data.preview, true);
      assert.equal(preview.data.applied, false);
      const afterPreview = await fixture.call("docx_read", {
        view,
        representation: "text",
        offset: 0,
        limit: 20,
      });
      assert.deepEqual(afterPreview.data.items, before.data.items);

      const request = {
        view,
        operationId: `apply-${action}`,
        action,
        arguments: argumentsValue,
        ...(selectorTarget ? { selectorTarget } : {}),
      };
      const applied = await fixture.call("docx_apply", request);
      assert.equal(applied.data.applied, true);
      assert.equal(applied.data.verificationPassed, true);
      const replay = await fixture.call("docx_apply", request);
      assert.deepEqual(replay.view, applied.view);
      assert.equal(replay.data.action, applied.data.action);

      const validation = await fixture.call("docx_validate", {
        view: applied.view,
      });
      assert.equal(validation.data.passed, true);
      assert.equal(validation.data.packageSummary.unrelatedBinaryChanges, 0);
      const comparison = await fixture.call("docx_compare", {
        view: applied.view,
      });
      assert.equal(comparison.data.hasChanges, true);
    } finally {
      fixture.service.close();
      fs.rmSync(fixture.root, { recursive: true, force: true });
    }
  });
