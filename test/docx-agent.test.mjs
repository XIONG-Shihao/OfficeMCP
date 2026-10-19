import test from "./support/test.mjs";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { manifest } from "../src/contracts/registry.mjs";
import { createServiceFixture } from "./service-fixture.mjs";
import { Client } from "@modelcontextprotocol/client";
import { StdioClientTransport } from "@modelcontextprotocol/client/stdio";
import { unzipSync, zipSync, strFromU8, strToU8 } from "fflate";

function setup({ libreOffice } = {}) {
  const fixture = createServiceFixture({
    prefix: "docx-agent-",
    toolset: "docx-agent",
    engines: libreOffice ? { libreOffice } : {},
    limits: { workerTimeoutMs: 30000 },
    contractVersion: "docx/6",
  });
  const { root, docs } = fixture;
  const imagePath = path.join(docs, "pixel.png");
  fs.writeFileSync(
    imagePath,
    Buffer.from(
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9ZQmcAAAAASUVORK5CYII=",
      "base64",
    ),
  );
  return { ...fixture, imagePath };
}

test("DOCX agent facade reads canonical items and applies a verified high-level action", async () => {
  const s = setup();
  try {
    assert.equal(manifest("docx-agent").length, 14);
    assert(manifest("docx-agent").every((tool) => tool.name.startsWith("docx_")));
    let view = (
      await s.call("docx_create", {
        operationId: "create-agent",
        document: {
          format: "docx",
          footer: "Confidential",
          blocks: [
            { kind: "heading", level: 1, text: "Terms" },
            { kind: "paragraph", text: "Payment within 30 days." },
            {
              kind: "table",
              rows: [
                ["Item", "Price"],
                ["Support", "100"],
              ],
            },
            {
              kind: "table",
              rows: [["Owner", "Finance"]],
            },
          ],
        },
      })
    ).view;

    const inspection = await s.call("docx_inspect", { view });
    assert(inspection.data.counts.paragraphs >= 2);
    assert(inspection.data.detectedDomains.includes("linear"));
    assert(inspection.data.readingPlan.some((step) => step.tool === "docx_read"));

    const linear = await s.call("docx_read", {
      view,
      representation: "linear",
      offset: 0,
      limit: 30,
    });
    assert(linear.data.items.some((item) => item.text.includes("30 days")));
    assert(linear.data.items.every((item) => item.itemId && item.story));

    const firstParagraphPage = await s.call("docx_query", {
      view,
      select: { type: "node", nodeType: "paragraph" },
      require: "all",
      offset: 0,
      limit: 1,
    });
    assert(firstParagraphPage.data.total > 1);
    assert.equal(firstParagraphPage.data.nextOffset, 1);
    const secondParagraphPage = await s.call("docx_query", {
      view,
      select: { type: "node", nodeType: "paragraph" },
      require: "all",
      offset: firstParagraphPage.data.nextOffset,
      limit: 1,
    });
    assert.equal(secondParagraphPage.data.items.length, 1);
    assert.notEqual(
      firstParagraphPage.data.items[0].itemId,
      secondParagraphPage.data.items[0].itemId,
    );

    const query = await s.call("docx_query", {
      view,
      select: { type: "text", pattern: "30 days" },
      require: "exactlyOne",
      offset: 0,
      limit: 10,
    });
    assert.equal(query.data.matchedTotal, 1);
    assert(query.data.items[0].target.id.startsWith("target_"));

    const edited = await s.call("docx_apply", {
      view,
      operationId: "agent-replace",
      action: "rewrite_block",
      selectorTarget: query.data.items[0].target,
      arguments: {
        text: "Payment within 45 days.",
        changeMode: "direct",
      },
    });
    assert.equal(edited.data.applied, true);
    assert.equal(edited.data.action, "rewrite_block");
    assert.equal(edited.data.verificationPassed, true);
    assert.equal(edited.data.engineEvidence, undefined);
    assert.equal(edited.data.receipt, undefined);
    view = edited.view;

    const after = await s.call("docx_read", {
      view,
      representation: "text",
      offset: 0,
      limit: 30,
    });
    assert(after.data.items.some((item) => item.text.includes("45 days")));

    const validation = await s.call("docx_validate", { view });
    assert.equal(validation.data.passed, true);
    assert(validation.data.checks.includes("output-XML-well-formed"));
    assert(validation.data.checks.includes("all-input-binary-members-preserved-by-content"));
    assert.equal(validation.data.evidence, undefined);
    assert.equal(validation.data.packageSummary.unrelatedBinaryChanges, 0);

    const table = await s.call("docx_read", {
      view,
      representation: "domain",
      domain: "tables",
      offset: 0,
      limit: 1,
    });
    assert.equal(table.data.total, 2);
    assert.equal(table.data.nextOffset, 1);
    assert.equal(table.data.items[0].rows[1][1].text, "100");
    const secondTable = await s.call("docx_read", {
      view,
      representation: "domain",
      domain: "tables",
      offset: table.data.nextOffset,
      limit: 1,
    });
    assert.equal(secondTable.data.items.length, 1);
    assert.notEqual(table.data.items[0].itemId, secondTable.data.items[0].itemId);
    assert(table.data.items[0].target);
    const duplicateCells = await s.service.call("docx_apply", {
      view,
      operationId: "agent-table-duplicate",
      action: "write_table_cells",
      selectorTarget: table.data.items[0].target,
      arguments: {
        cells: [
          { row: 1, column: 1, expectedText: "100", text: "125" },
          { row: 1, column: 1, expectedText: "100", text: "150" },
        ],
      },
    });
    assert.equal(duplicateCells.payload.error.code, "INVALID_ARGUMENT");
    view = (
      await s.call("docx_apply", {
        view,
        operationId: "agent-table",
        action: "write_table_cells",
        selectorTarget: table.data.items[0].target,
        arguments: {
          cells: [{ row: 1, column: 1, expectedText: "100", text: "125" }],
        },
      })
    ).view;
    let secondTableTarget = (
      await s.call("docx_read", {
        view,
        representation: "domain",
        domain: "tables",
        offset: 1,
        limit: 1,
      })
    ).data.items[0].target;
    view = (
      await s.call("docx_apply", {
        view,
        operationId: "agent-table-merge",
        action: "merge_table_cells",
        selectorTarget: secondTableTarget,
        arguments: {
          start: { row: 0, column: 0 },
          end: { row: 0, column: 1 },
        },
      })
    ).view;
    secondTableTarget = (
      await s.call("docx_read", {
        view,
        representation: "domain",
        domain: "tables",
        offset: 1,
        limit: 1,
      })
    ).data.items[0].target;
    view = (
      await s.call("docx_apply", {
        view,
        operationId: "agent-table-unmerge",
        action: "unmerge_table_cells",
        selectorTarget: secondTableTarget,
        arguments: { row: 0, column: 0 },
      })
    ).view;

    view = (
      await s.call("docx_apply", {
        view,
        operationId: "agent-image",
        action: "insert_image",
        arguments: {
          path: s.imagePath,
          alt: "Status indicator",
          width: 24,
          height: 24,
          placement: "end",
        },
      })
    ).view;
    let images = await s.call("docx_read", {
      view,
      representation: "domain",
      domain: "images",
      offset: 0,
      limit: 10,
    });
    assert.equal(images.data.total, 1);
    assert(images.data.items[0].target);
    assert.equal(images.data.items[0].altText, "Status indicator");
    assert.match(images.data.items[0].packagePart, /^word\/media\//);
    view = (
      await s.call("docx_apply", {
        view,
        operationId: "agent-image-alt",
        action: "set_image_alt_text",
        selectorTarget: images.data.items[0].target,
        arguments: { description: "Updated status indicator" },
      })
    ).view;
    images = await s.call("docx_read", {
      view,
      representation: "domain",
      domain: "images",
      offset: 0,
      limit: 10,
    });
    assert.equal(images.data.items[0].altText, "Updated status indicator");
    view = (
      await s.call("docx_apply", {
        view,
        operationId: "agent-image-resize",
        action: "resize_image",
        selectorTarget: images.data.items[0].target,
        arguments: { width: 30, height: 18, unit: "pt" },
      })
    ).view;
    images = await s.call("docx_read", {
      view,
      representation: "domain",
      domain: "images",
      offset: 0,
      limit: 10,
    });
    view = (
      await s.call("docx_apply", {
        view,
        operationId: "agent-image-rotate",
        action: "rotate_image",
        selectorTarget: images.data.items[0].target,
        arguments: { angle: 15 },
      })
    ).view;
    images = await s.call("docx_read", {
      view,
      representation: "domain",
      domain: "images",
      offset: 0,
      limit: 10,
    });
    view = (
      await s.call("docx_apply", {
        view,
        operationId: "agent-image-replace",
        action: "replace_image",
        selectorTarget: images.data.items[0].target,
        arguments: { path: s.imagePath, resetSize: false },
      })
    ).view;
    images = await s.call("docx_read", {
      view,
      representation: "domain",
      domain: "images",
      offset: 0,
      limit: 10,
    });
    const image = await s.service.call("docx_image_get", {
      view,
      target: images.data.items[0].target,
    });
    assert(!image.isError, JSON.stringify(image.payload));
    assert.equal(image.payload.data.mediaType, "image/png");
    assert(image.content[0].data.length > 50);

    const existingComments = await s.call("docx_read", { view, representation: "domain", domain: "comments", limit: 20 });
    assert.equal(existingComments.data.total, 0);

    const finalValidation = await s.call("docx_validate", { view });
    assert(
      finalValidation.data.checks.includes(
        "only-targeted-media-binaries-and-relationships-may-change",
      ),
    );

    const comparison = await s.call("docx_compare", { view });
    assert.equal(comparison.data.hasChanges, true);
    assert(comparison.data.changedComponents.includes("body"));

    const recovered = await s.call("docx_operation_get", {
      operationId: "agent-replace",
    });
    assert.equal(recovered.operationId, "agent-replace");
    assert.equal(recovered.data.verificationPassed, true);
    assert.equal(recovered.data.receipt, undefined);
    assert.equal(recovered.data.engineEvidence, undefined);

    const saved = await s.call("docx_save", {
      view,
      operationId: "agent-save",
      destination: { rootId: "output", relativePath: "agent.docx" },
    });
    assert(fs.existsSync(saved.data.path));
    const recoveredSave = await s.call("docx_operation_get", {
      operationId: "agent-save",
    });
    assert.equal(recoveredSave.data.published, true);
    assert.equal(recoveredSave.data.path, saved.data.path);
    assert.equal(recoveredSave.data.assetId, saved.data.assetId);
    await s.call("docx_close", { view, disposition: "retain" });
  } finally {
    s.service.close();
    fs.rmSync(s.root, { recursive: true, force: true });
  }
});

test("DOCX agent inserts a heading through the public action facade", async () => {
  const s = setup();
  try {
    let view = (
      await s.call("docx_create", {
        operationId: "heading-create",
        document: {
          format: "docx",
          blocks: [{ kind: "paragraph", text: "Body" }],
        },
      })
    ).view;
    const applied = await s.call("docx_apply", {
      view,
      operationId: "heading-insert",
      action: "insert_heading",
      arguments: {
        text: "Inserted heading",
        level: 2,
        placement: { at: "document_start" },
      },
    });
    assert.equal(applied.data.verificationPassed, true);
    view = applied.view;
    const headings = await s.call("docx_read", {
      view,
      representation: "linear",
      offset: 0,
      limit: 20,
    });
    assert(
      headings.data.items.some(
        (item) => item.kind === "heading" && item.text === "Inserted heading",
      ),
    );
  } finally {
    s.service.close();
    fs.rmSync(s.root, { recursive: true, force: true });
  }
});

test("DOCX agent facade returns preset-specific recovery and explicit PDF export requirements", async () => {
  const s = setup();
  try {
    const view = (
      await s.call("docx_create", {
        operationId: "create-export-requirements",
        document: {
          format: "docx",
          blocks: [{ kind: "paragraph", text: "Export me" }],
        },
      })
    ).view;
    const stale = await s.service.call("docx_apply", {
      view,
      operationId: "bad-target",
      action: "rewrite_block",
      selectorTarget: { id: "missing" },
      arguments: { text: "Changed" },
    });
    assert.equal(stale.payload.error.code, "NOT_FOUND");
    assert(!stale.payload.error.recovery.includes("office_read"));

    const exported = await s.service.call("docx_export", {
      view,
      operationId: "export-without-converter",
      format: "pdf",
      destination: { rootId: "output", relativePath: "exported.pdf" },
    });
    assert.equal(exported.payload.error.code, "ENGINE_UNAVAILABLE");
    assert.match(exported.payload.error.recovery, /LibreOffice/);
  } finally {
    s.service.close();
    fs.rmSync(s.root, { recursive: true, force: true });
  }
});

test("DOCX agent facade normalizes tracked changes for review", async () => {
  const s = setup();
  try {
    let view = (
      await s.call("docx_create", {
        operationId: "create-tracked",
        document: {
          format: "docx",
          blocks: [{ kind: "paragraph", text: "The old review clause." }],
        },
      })
    ).view;
    const target = (
      await s.call("docx_query", {
        view,
        select: { type: "text", pattern: "The old review clause." },
        require: "exactlyOne",
        offset: 0,
        limit: 10,
      })
    ).data.items[0].target;
    view = (
      await s.call("docx_apply", {
        view,
        operationId: "tracked-rewrite",
        action: "rewrite_block",
        selectorTarget: target,
        arguments: {
          text: "The revised review clause.",
          changeMode: "tracked",
        },
        author: "Reviewer",
      })
    ).view;
    const review = await s.call("docx_read", {
      view,
      representation: "domain",
      domain: "tracked_changes",
      offset: 0,
      limit: 20,
    });
    assert(review.data.total > 0);
    assert(
      review.data.items.every(
        (item) => item.changeType && item.target && item.itemId,
      ),
    );
  } finally {
    s.service.close();
    fs.rmSync(s.root, { recursive: true, force: true });
  }
});

for (const decision of ["accept", "reject"])
  test(`DOCX agent can ${decision} one explicitly selected tracked change`, async () => {
    const s = setup();
    try {
      let view = (
        await s.call("docx_create", {
          operationId: `${decision}-change-create`,
          document: {
            format: "docx",
            blocks: [{ kind: "paragraph", text: "Original clause" }],
          },
        })
      ).view;
      const target = (
        await s.call("docx_query", {
          view,
          select: { type: "text", pattern: "Original clause" },
          require: "exactlyOne",
          offset: 0,
          limit: 2,
        })
      ).data.items[0].target;
      view = (
        await s.call("docx_apply", {
          view,
          operationId: `${decision}-tracked-edit`,
          action: "rewrite_block",
          selectorTarget: target,
          arguments: { text: "Revised clause", changeMode: "tracked" },
          author: "Reviewer",
        })
      ).view;
      const changes = await s.call("docx_query", {
        view,
        select: {
          type: "entity",
          domain: "tracked_changes",
          author: "Reviewer",
        },
        require: "first",
        offset: 0,
        limit: 1,
      });
      view = (
        await s.call("docx_apply", {
          view,
          operationId: `${decision}-one-change`,
          action: `${decision}_tracked_change`,
          selectorTarget: changes.data.items[0].target,
          arguments: {},
        })
      ).view;
      const validation = await s.call("docx_validate", { view });
      assert.equal(validation.data.passed, true);
      const text = JSON.stringify(
        (
          await s.call("docx_read", {
            view,
            representation: "text",
            offset: 0,
            limit: 20,
          })
        ).data.items,
      );
      assert(text.includes(decision === "accept" ? "Revised clause" : "Original clause"));
    } finally {
      s.service.close();
      fs.rmSync(s.root, { recursive: true, force: true });
    }
  });

test("DOCX agent descriptors, complete-read guidance and previewable plans are machine-readable", async () => {
  const s = setup();
  try {
    let view = (
      await s.call("docx_create", {
        operationId: "plan-create",
        document: {
          format: "docx",
          footer: "Plan footer",
          blocks: [
            { kind: "heading", level: 1, text: "Plan" },
            { kind: "paragraph", text: "Alpha clause" },
            { kind: "paragraph", text: "Beta clause" },
          ],
        },
      })
    ).view;

    const inspection = await s.call("docx_inspect", {
      view,
      actions: ["rewrite_block", "merge_table_cells", "set_checkbox_control"],
    });
    const rewrite = inspection.data.actionDescriptors.rewrite_block;
    assert(inspection.data.availableActions.includes("rewrite_block"));
    assert(!inspection.data.availableActions.includes("undo_changes"));
    assert.equal(inspection.data.qualifiedActions, undefined);
    assert.equal(inspection.data.experimentalActions, undefined);
    assert.equal(inspection.data.knownLimitedActions, undefined);
    assert.equal(inspection.data.actionSelectionPolicy, undefined);
    assert(Array.isArray(inspection.data.storyInventory));
    assert(Array.isArray(inspection.data.objectInventory));
    assert.equal(typeof inspection.data.accessibilityReport.passed, "boolean");
    assert(Array.isArray(inspection.data.accessibilityReport.notChecked));
    assert.equal(typeof inspection.data.templateReport.missingDataCount, "number");
    assert.equal(rewrite.requiredTargetKind, "block");
    assert.equal(rewrite.atomicity, "superdoc-action");
    assert.equal(rewrite.qualification, undefined);
    assert.equal(rewrite.qualificationEvidence, undefined);
    assert.equal(rewrite.argumentSchema.properties.text.type, "string");
    assert.equal(
      inspection.data.actionDescriptors.set_checkbox_control.argumentSchema
        .properties.checked.type,
      "boolean",
    );

    const invalid = await s.service.call("docx_apply", {
      view,
      operationId: "invalid-plan",
      action: "rewrite_block",
      arguments: {},
      select: { type: "text", pattern: "Alpha clause" },
    });
    assert.equal(invalid.payload.error.code, "INVALID_ARGUMENT");
    assert.match(invalid.payload.error.message, /rewrite_block.*text/);

    const formerlyExperimental = await s.call("docx_apply", {
      view,
      mode: "preview",
      action: "delete_blocks",
      arguments: {},
      select: { type: "text", pattern: "Beta clause" },
    });
    assert.equal(formerlyExperimental.data.preview, true);
    assert.deepEqual(formerlyExperimental.warnings, []);

    const inventory = await s.call("docx_read", {
      view,
      representation: "inventory",
      offset: 0,
      limit: 100,
    });
    assert.equal(inventory.coverage.completeDocument, false);
    assert(inventory.data.remainingReads.length > 0);
    assert(
      Object.values(inventory.coverage.matrix).includes("detected_not_read"),
    );

    const complete = await s.call("docx_read", {
      view,
      representation: "complete",
      offset: 0,
      limit: 1,
    });
    assert(complete.data.readSnapshotId);
    assert.equal(complete.coverage.contentComplete, false);
    assert.equal(
      complete.data.remainingReads[0].readSnapshotId,
      complete.data.readSnapshotId,
    );
    const completedSnapshot = await s.call("docx_read", {
      view,
      representation: "complete",
      readSnapshotId: complete.data.readSnapshotId,
      offset: 0,
      limit: 100,
    });
    assert.equal(completedSnapshot.coverage.contentComplete, true);
    assert.equal(completedSnapshot.coverage.matrix.body, "complete");
    assert.equal(completedSnapshot.coverage.matrix.headers_footers, "complete");
    assert.equal(completedSnapshot.coverage.dimensions.storyInventory, "complete");
    assert(
      completedSnapshot.data.items.some((item) => item.text.includes("Plan footer")),
    );

    const structured = await s.call("docx_read", {
      view,
      representation: "structured",
      offset: 0,
      limit: 20,
    });
    assert(structured.data.items.every((item) => Array.isArray(item.path)));
    assert(
      structured.data.items.every(
        (item) => item.provenance?.engine && !item.nativeTarget,
      ),
    );

    const steps = [
      {
        action: "rewrite_block",
        arguments: { text: "Alpha revised" },
        select: { type: "text", pattern: "Alpha clause" },
      },
      {
        action: "rewrite_block",
        arguments: { text: "Beta revised" },
        select: { type: "text", pattern: "Beta clause" },
      },
    ];
    const preview = await s.call("docx_apply", {
      view,
      mode: "preview",
      steps,
    });
    assert.equal(preview.data.preview, true);
    assert.equal(preview.data.applied, false);
    assert.equal(preview.data.steps.length, 2);
    const unchanged = await s.call("docx_read", {
      view,
      representation: "text",
      offset: 0,
      limit: 20,
    });
    assert(JSON.stringify(unchanged.data.items).includes("Alpha clause"));

    const applied = await s.call("docx_apply", {
      view,
      mode: "apply",
      operationId: "apply-plan",
      steps,
    });
    view = applied.view;
    assert.equal(applied.data.action, "mutation_plan");
    assert.equal(applied.data.steps.length, 2);
    assert.equal(applied.data.verificationPassed, true);
    const changed = await s.call("docx_read", {
      view,
      representation: "text",
      offset: 0,
      limit: 20,
    });
    const changedText = JSON.stringify(changed.data.items);
    assert(changedText.includes("Alpha revised"));
    assert(changedText.includes("Beta revised"));

    view = (
      await s.call("docx_apply", {
        view,
        operationId: "insert-agent-toc",
        action: "insert_toc",
        arguments: { title: "Contents" },
      })
    ).view;
    const toc = await s.call("docx_read", {
      view,
      representation: "domain",
      domain: "toc",
      offset: 0,
      limit: 10,
    });
    assert(toc.data.total > 0);
    view = (
      await s.call("docx_apply", {
        view,
        operationId: "refresh-agent-toc",
        action: "update_toc",
        arguments: {},
      })
    ).view;
    const tocValidation = await s.call("docx_validate", { view });
    assert.equal(tocValidation.data.passed, true);
  } finally {
    s.service.close();
    fs.rmSync(s.root, { recursive: true, force: true });
  }
});

test("DOCX structural selectors, composite workflows and postconditions remain deterministic", async () => {
  const s = setup();
  try {
    let view = (
      await s.call("docx_create", {
        operationId: "workflow-create",
        document: {
          format: "docx",
          blocks: [
            { kind: "heading", level: 1, text: "Commercial" },
            { kind: "heading", level: 2, text: "Pricing" },
            { kind: "paragraph", text: "Client: {{CLIENT}}" },
            { kind: "table", rows: [["Service", "Price"], ["Support", "100"]] },
          ],
        },
      })
    ).view;

    const heading = await s.call("docx_query", {
      view,
      select: { type: "headingPath", path: ["Commercial", "Pricing"] },
      require: "exactlyOne",
      offset: 0,
      limit: 10,
    });
    assert.equal(heading.data.items[0].text, "Pricing");
    assert.equal(heading.data.coverage, undefined);
    assert.equal(heading.data.engineSession, undefined);

    const table = await s.call("docx_query", {
      view,
      select: { type: "table", header: "Service", contains: "Support" },
      require: "exactlyOne",
      offset: 0,
      limit: 10,
    });
    assert.equal(table.data.items[0].rows[1][1].text, "100");

    const wrongCount = await s.service.call("docx_apply", {
      view,
      operationId: "wrong-template-count",
      workflow: {
        type: "fill_template",
        bindings: [
          { placeholder: "{{CLIENT}}", value: "Meridian Labs", expectedCount: 2 },
        ],
      },
    });
    assert.equal(wrongCount.payload.error.code, "TARGET_CONFLICT");

    const filled = await s.call("docx_apply", {
      view,
      operationId: "fill-template",
      workflow: {
        type: "fill_template",
        bindings: [
          { placeholder: "{{CLIENT}}", value: "Meridian Labs", expectedCount: 1 },
        ],
      },
    });
    assert.equal(filled.data.workflow, "fill_template");
    assert.equal(filled.data.postconditions.passed, true);
    assert.equal(filled.data.semanticDelta.changedOccurrences, 1);
    view = filled.view;

    const updatedTableTarget = (
      await s.call("docx_query", {
        view,
        select: { type: "table", header: "Service", contains: "Support" },
        require: "exactlyOne",
        offset: 0,
        limit: 10,
      })
    ).data.items[0].target;
    const updated = await s.call("docx_apply", {
      view,
      operationId: "update-table-workflow",
      selectorTarget: updatedTableTarget,
      workflow: {
        type: "update_table",
        cells: [{ row: 1, column: 1, expectedText: "100", text: "125" }],
        caption: { position: "below", label: "Table", text: "Service fees" },
      },
    });
    assert.equal(updated.data.workflow, "update_table");
    view = updated.view;
    const sectionTarget = (
      await s.call("docx_read", {
        view,
        representation: "domain",
        domain: "sections",
        offset: 0,
        limit: 1,
      })
    ).data.items[0].target;
    view = (
      await s.call("docx_apply", {
        view,
        operationId: "set-section-layout",
        action: "set_section_layout",
        selectorTarget: sectionTarget,
        arguments: {
          orientation: "landscape",
          marginsPt: { top: 54, right: 54, bottom: 54, left: 54 },
        },
      })
    ).view;
    const validation = await s.call("docx_validate", { view });
    assert.equal(validation.data.passed, true);

    const unavailableExport = await s.service.call("docx_export", {
      view,
      operationId: "export-pdf",
      format: "pdf",
      destination: { rootId: "output", relativePath: "result.pdf" },
    });
    assert.equal(unavailableExport.payload.error.code, "ENGINE_UNAVAILABLE");
  } finally {
    s.service.close();
    fs.rmSync(s.root, { recursive: true, force: true });
  }
});

test("DOCX actors isolate independent documents while serializing each document", async () => {
  const s = setup();
  try {
    const first = await s.call("docx_create", {
      operationId: "actor-first",
      document: {
        format: "docx",
        blocks: [{ kind: "paragraph", text: "First actor" }],
      },
    });
    const second = await s.call("docx_create", {
      operationId: "actor-second",
      document: {
        format: "docx",
        blocks: [{ kind: "paragraph", text: "Second actor" }],
      },
    });
    const [firstRead, secondRead] = await Promise.all([
      s.call("docx_read", {
        view: first.view,
        representation: "text",
        offset: 0,
        limit: 10,
      }),
      s.call("docx_read", {
        view: second.view,
        representation: "text",
        offset: 0,
        limit: 10,
      }),
    ]);
    assert.equal(firstRead.data.engineSession, undefined);
    assert.equal(secondRead.data.engineSession, undefined);
    assert(firstRead.data.items.some((item) => item.text.includes("First actor")));
    assert(secondRead.data.items.some((item) => item.text.includes("Second actor")));
    assert(s.service.docxPool.actors.size >= 2);
  } finally {
    s.service.close();
    fs.rmSync(s.root, { recursive: true, force: true });
  }
});

test("DOCX broad formatting actions use typed targets and survive package validation", async () => {
  const s = setup();
  try {
    let view = (
      await s.call("docx_create", {
        operationId: "format-create",
        document: {
          format: "docx",
          blocks: [
            { kind: "paragraph", text: "Formatted clause" },
            { kind: "table", rows: [["Heading", "Value"], ["A", "1"]] },
          ],
        },
      })
    ).view;

    const paragraph = await s.call("docx_query", {
      view,
      select: { type: "text", pattern: "Formatted clause" },
      require: "exactlyOne",
      offset: 0,
      limit: 2,
    });
    view = (
      await s.call("docx_apply", {
        view,
        operationId: "format-paragraph-advanced",
        action: "format_paragraph_advanced",
        selectorTarget: paragraph.data.items[0].target,
        arguments: {
          indentation: { left: 18, firstLine: 9 },
          tabStop: { position: 144, alignment: "right", leader: "dot" },
          shading: { fill: "FFF2CC" },
          borders: [
            { edge: "bottom", lineStyle: "single", lineWeightPt: 0.5, color: "808080" },
          ],
        },
      })
    ).view;

    const cells = await s.call("docx_query", {
      view,
      select: { type: "node", nodeType: "tableCell" },
      require: "all",
      offset: 0,
      limit: 20,
    });
    assert.equal(cells.data.total, 4);
    view = (
      await s.call("docx_apply", {
        view,
        operationId: "format-cell",
        action: "format_table_cell",
        selectorTarget: cells.data.items[0].target,
        arguments: {
          verticalAlign: "center",
          wrapText: true,
          padding: { topPt: 2, rightPt: 3, bottomPt: 2, leftPt: 3 },
          shading: "D9EAF7",
          borders: [
            { edge: "bottom", lineStyle: "single", lineWeightPt: 0.5, color: "4472C4" },
          ],
        },
      })
    ).view;

    const rows = await s.call("docx_query", {
      view,
      select: { type: "node", nodeType: "tableRow" },
      require: "all",
      offset: 0,
      limit: 20,
    });
    assert.equal(rows.data.total, 2);
    view = (
      await s.call("docx_apply", {
        view,
        operationId: "format-row",
        action: "format_table_row",
        selectorTarget: rows.data.items[0].target,
        arguments: { repeatHeader: true, allowBreakAcrossPages: false, heightPt: 18, heightRule: "atLeast" },
      })
    ).view;

    const validation = await s.call("docx_validate", { view });
    assert.equal(validation.data.passed, true);
    assert(validation.data.checks.includes("OPC-relationship-targets-resolve"));
    const saved = await s.call("docx_save", {
      view,
      operationId: "format-save",
      destination: { rootId: "output", relativePath: "formatted.docx" },
    });
    const documentXml = strFromU8(unzipSync(fs.readFileSync(saved.data.path))["word/document.xml"]);
    assert.match(documentXml, /w:tblHeader/);
    assert.match(documentXml, /w:tcMar/);
    assert.match(documentXml, /w:tabs/);
  } finally {
    s.service.close();
    fs.rmSync(s.root, { recursive: true, force: true });
  }
});

test("DOCX reference maintenance and inherited header materialization preserve unrelated content", async () => {
  const {
    Document,
    Packer,
    Paragraph,
    TextRun,
    Header,
    ExternalHyperlink,
    FootnoteReferenceRun,
    SimpleField,
  } = await import("docx");
  const s = setup();
  try {
    const file = path.join(s.docs, "references.docx");
    fs.writeFileSync(
      file,
      await Packer.toBuffer(
        new Document({
          footnotes: {
            1: { children: [new Paragraph("Original footnote")] },
          },
          sections: [
            {
              headers: {
                default: new Header({ children: [new Paragraph("Shared header")] }),
              },
              children: [new Paragraph("First section")],
            },
            {
              children: [
                new Paragraph({
                  children: [
                    new ExternalHyperlink({
                      link: "https://example.invalid/old",
                      children: [new TextRun("Reference link")],
                    }),
                    new TextRun(" "),
                    new FootnoteReferenceRun(1),
                  ],
                }),
                new Paragraph({ children: [new SimpleField("PAGE")] }),
              ],
            },
          ],
        }),
      ),
    );
    let view = (await s.call("docx_open", { path: file })).view;

    const sections = await s.call("docx_read", {
      view,
      representation: "domain",
      domain: "sections",
      offset: 0,
      limit: 20,
    });
    assert(sections.data.total >= 2);
    view = (
      await s.call("docx_apply", {
        view,
        operationId: "materialize-header",
        action: "materialize_header_footer",
        arguments: {
          sectionId: sections.data.items[1].itemId,
          kind: "header",
          variant: "default",
        },
      })
    ).view;

    view = (
      await s.call("docx_apply", {
        view,
        operationId: "replace-materialized-header",
        action: "replace_header_footer_text",
        arguments: {
          sectionId: sections.data.items[1].itemId,
          kind: "header",
          variant: "default",
          find: "Shared header",
          replace: "Section-specific header",
        },
      })
    ).view;

    let hyperlinks = await s.call("docx_read", {
      view,
      representation: "domain",
      domain: "hyperlinks",
      offset: 0,
      limit: 20,
    });
    assert(hyperlinks.data.items[0].target);
    view = (
      await s.call("docx_apply", {
        view,
        operationId: "patch-link",
        action: "patch_hyperlink",
        selectorTarget: hyperlinks.data.items[0].target,
        arguments: { href: "https://example.invalid/new", tooltip: "Updated destination" },
      })
    ).view;

    hyperlinks = await s.call("docx_read", {
      view,
      representation: "domain",
      domain: "hyperlinks",
      offset: 0,
      limit: 20,
    });
    assert.equal(hyperlinks.data.items[0].details.properties.href, "https://example.invalid/new");

    const footnotes = await s.call("docx_read", {
      view,
      representation: "domain",
      domain: "footnotes",
      offset: 0,
      limit: 20,
    });
    assert(footnotes.data.items[0].target);
    view = (
      await s.call("docx_apply", {
        view,
        operationId: "update-footnote",
        action: "update_footnote",
        selectorTarget: footnotes.data.items[0].target,
        arguments: { content: "Updated footnote" },
      })
    ).view;

    const validation = await s.call("docx_validate", { view });
    assert.equal(validation.data.passed, true);
    const headers = await s.call("docx_read", {
      view,
      representation: "domain",
      domain: "headers_footers",
      offset: 0,
      limit: 100,
    });
    const explicitHeaders = headers.data.items.filter(
      (item) => item.kind === "header" && item.details.resolution?.status === "explicit",
    );
    assert(explicitHeaders.length >= 2);
    assert(explicitHeaders.some((item) => item.text.includes("Shared header")));
    assert(
      explicitHeaders.some((item) => item.text.includes("Section-specific header")),
    );
  } finally {
    s.service.close();
    fs.rmSync(s.root, { recursive: true, force: true });
  }
});

test("DOCX typed checkbox, date and choice controls are independently addressable", async () => {
  const { Document, Packer, Paragraph } = await import("docx");
  const s = setup();
  try {
    const source = await Packer.toBuffer(
      new Document({
        sections: [
          {
            children: [
              new Paragraph("CHECKBOX_MARKER"),
              new Paragraph("DATE_MARKER"),
              new Paragraph("CHOICE_MARKER"),
              new Paragraph("TEXT_MARKER"),
            ],
          },
        ],
      }),
    );
    const parts = unzipSync(source);
    let xml = strFromU8(parts["word/document.xml"]);
    const replaceRun = (marker, value) => {
      const pattern = new RegExp(`<w:p(?:\\s[^>]*)?>(?:(?!<\\/w:p>)[\\s\\S])*?${marker}(?:(?!<\\/w:p>)[\\s\\S])*?<\\/w:p>`);
      assert(pattern.test(xml), `missing ${marker}`);
      xml = xml.replace(pattern, `<w:p>${value}</w:p>`);
    };
    replaceRun(
      "CHECKBOX_MARKER",
      '<w:sdt><w:sdtPr><w:id w:val="101"/><w:tag w:val="check"/><w14:checkbox><w14:checked w14:val="0"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox></w:sdtPr><w:sdtContent><w:r><w:t>☐</w:t></w:r></w:sdtContent></w:sdt>',
    );
    replaceRun(
      "DATE_MARKER",
      '<w:sdt><w:sdtPr><w:id w:val="102"/><w:tag w:val="date"/><w:date w:fullDate="2026-01-01T00:00:00Z"><w:dateFormat w:val="yyyy-MM-dd"/><w:lid w:val="en-US"/></w:date></w:sdtPr><w:sdtContent><w:r><w:t>2026-01-01</w:t></w:r></w:sdtContent></w:sdt>',
    );
    replaceRun(
      "CHOICE_MARKER",
      '<w:sdt><w:sdtPr><w:id w:val="103"/><w:tag w:val="choice"/><w:dropDownList><w:listItem w:displayText="Alpha" w:value="Alpha"/><w:listItem w:displayText="Beta" w:value="Beta"/></w:dropDownList></w:sdtPr><w:sdtContent><w:r><w:t>Alpha</w:t></w:r></w:sdtContent></w:sdt>',
    );
    replaceRun(
      "TEXT_MARKER",
      '<w:sdt><w:sdtPr><w:id w:val="104"/><w:tag w:val="text"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>Original</w:t></w:r></w:sdtContent></w:sdt>',
    );
    parts["word/document.xml"] = strToU8(xml);
    const file = path.join(s.docs, "controls.docx");
    fs.writeFileSync(file, zipSync(parts));
    let view = (await s.call("docx_open", { path: file })).view;

    let controls = await s.call("docx_read", {
      view,
      representation: "domain",
      domain: "content_controls",
      offset: 0,
      limit: 20,
    });
    assert.equal(controls.data.total, 4);
    assert(controls.data.items.every((item) => item.target));

    const byTag = (items, tag) =>
      items.find(
        (item) => item.details.tag === tag || item.details.properties?.tag === tag,
      );
    view = (
      await s.call("docx_apply", {
        view,
        operationId: "set-checkbox",
        action: "set_checkbox_control",
        selectorTarget: byTag(controls.data.items, "check").target,
        arguments: { checked: true },
      })
    ).view;
    controls = await s.call("docx_read", {
      view,
      representation: "domain",
      domain: "content_controls",
      offset: 0,
      limit: 20,
    });
    view = (
      await s.call("docx_apply", {
        view,
        operationId: "set-date",
        action: "set_date_control",
        selectorTarget: byTag(controls.data.items, "date").target,
        arguments: { value: "2026-09-15" },
      })
    ).view;
    controls = await s.call("docx_read", {
      view,
      representation: "domain",
      domain: "content_controls",
      offset: 0,
      limit: 20,
    });
    const choice = byTag(controls.data.items, "choice");
    view = (
      await s.call("docx_apply", {
        view,
        operationId: "set-choice",
        action: "set_choice_control",
        selectorTarget: choice.target,
        arguments: { value: "Beta" },
      })
    ).view;
    controls = await s.call("docx_read", {
      view,
      representation: "domain",
      domain: "content_controls",
      offset: 0,
      limit: 20,
    });
    view = (
      await s.call("docx_apply", {
        view,
        operationId: "set-control-text",
        action: "set_content_control_text",
        selectorTarget: byTag(controls.data.items, "text").target,
        arguments: { expectedText: "Original", text: "Updated" },
      })
    ).view;
    const validation = await s.call("docx_validate", { view });
    assert.equal(validation.data.passed, true);
    const finalControls = await s.call("docx_read", {
      view,
      representation: "domain",
      domain: "content_controls",
      offset: 0,
      limit: 20,
    });
    const controlValue = JSON.stringify(finalControls.data.items);
    assert(controlValue.includes("2026-09-15"), controlValue);
    assert(controlValue.includes("Beta"), controlValue);
    assert(controlValue.includes("true"), controlValue);
    assert(controlValue.includes("Updated"), controlValue);
  } finally {
    s.service.close();
    fs.rmSync(s.root, { recursive: true, force: true });
  }
});

test("DOCX repeating-section controls can clone an item without losing the template", async () => {
  const { Document, Packer, Paragraph } = await import("docx");
  const s = setup();
  try {
    const source = await Packer.toBuffer(
      new Document({
        sections: [{ children: [new Paragraph("REPEATING_MARKER")] }],
      }),
    );
    const parts = unzipSync(source);
    let xml = strFromU8(parts["word/document.xml"]);
    if (!xml.includes("xmlns:w15="))
      xml = xml.replace(
        /<w:document\s/,
        '<w:document xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" ',
      );
    const paragraph = /<w:p(?:\s[^>]*)?>(?:(?!<\/w:p>)[\s\S])*?REPEATING_MARKER(?:(?!<\/w:p>)[\s\S])*?<\/w:p>/;
    assert(paragraph.test(xml));
    xml = xml.replace(
      paragraph,
      '<w:sdt><w:sdtPr><w:id w:val="201"/><w:tag w:val="repeat"/><w15:repeatingSection w15:sectionTitle="Line item"/></w:sdtPr><w:sdtContent><w:sdt><w:sdtPr><w:id w:val="202"/><w15:repeatingSectionItem/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Template item</w:t></w:r></w:p></w:sdtContent></w:sdt></w:sdtContent></w:sdt>',
    );
    parts["word/document.xml"] = strToU8(xml);
    const file = path.join(s.docs, "repeating-control.docx");
    fs.writeFileSync(file, zipSync(parts));
    let view = (await s.call("docx_open", { path: file })).view;
    const inspection = await s.call("docx_inspect", { view });
    assert.equal(inspection.data.templateReport.repeatingSections.length, 1);
    const controls = await s.call("docx_read", {
      view,
      representation: "domain",
      domain: "content_controls",
      offset: 0,
      limit: 20,
    });
    const repeating = controls.data.items.find(
      (item) =>
        item.details.tag === "repeat" || item.details.properties?.tag === "repeat",
    );
    assert(repeating?.target);
    view = (
      await s.call("docx_apply", {
        view,
        operationId: "repeat-clone",
        action: "edit_repeating_section",
        selectorTarget: repeating.target,
        arguments: { action: "clone", index: 0 },
      })
    ).view;
    const text = await s.call("docx_read", {
      view,
      representation: "text",
      offset: 0,
      limit: 20,
    });
    assert(
      (JSON.stringify(text.data.items).match(/Template item/g)?.length ?? 0) >=
        2,
    );
    const validation = await s.call("docx_validate", { view });
    assert.equal(validation.data.passed, true);
  } finally {
    s.service.close();
    fs.rmSync(s.root, { recursive: true, force: true });
  }
});

test("DOCX endnotes are inventoried, readable and updateable through typed targets", async () => {
  const { Document, Packer, Paragraph } = await import("docx");
  const s = setup();
  try {
    const parts = unzipSync(
      await Packer.toBuffer(
        new Document({ sections: [{ children: [new Paragraph("Endnote anchor")] }] }),
      ),
    );
    let documentXml = strFromU8(parts["word/document.xml"]);
    documentXml = documentXml.replace(
      /(<w:t[^>]*>Endnote anchor<\/w:t>[\s\S]*?<\/w:r>)/,
      '$1<w:r><w:endnoteReference w:id="1"/></w:r>',
    );
    parts["word/document.xml"] = strToU8(documentXml);
    parts["word/endnotes.xml"] = strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:id="1"><w:p><w:r><w:t>Original endnote</w:t></w:r></w:p></w:endnote></w:endnotes>',
    );
    let relationships = strFromU8(parts["word/_rels/document.xml.rels"]);
    relationships = relationships.replace(
      "</Relationships>",
      '<Relationship Id="rIdOfficeMcpEndnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/></Relationships>',
    );
    parts["word/_rels/document.xml.rels"] = strToU8(relationships);
    let contentTypes = strFromU8(parts["[Content_Types].xml"]);
    contentTypes = contentTypes.replace(
      "</Types>",
      '<Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/></Types>',
    );
    parts["[Content_Types].xml"] = strToU8(contentTypes);
    const file = path.join(s.docs, "endnotes.docx");
    fs.writeFileSync(file, zipSync(parts));
    let view = (await s.call("docx_open", { path: file })).view;
    const inspection = await s.call("docx_inspect", {
      view,
      includeDomains: ["endnotes"],
    });
    assert(
      inspection.data.storyInventory.some(
        (story) => story.type === "endnote" && story.id === "1",
      ),
    );
    const endnotes = await s.call("docx_read", {
      view,
      representation: "domain",
      domain: "endnotes",
      offset: 0,
      limit: 20,
    });
    assert.equal(endnotes.data.total, 1);
    assert(endnotes.data.items[0].target);
    view = (
      await s.call("docx_apply", {
        view,
        operationId: "endnote-update",
        action: "update_endnote",
        selectorTarget: endnotes.data.items[0].target,
        arguments: { content: "Updated endnote" },
      })
    ).view;
    const updated = await s.call("docx_read", {
      view,
      representation: "domain",
      domain: "endnotes",
      offset: 0,
      limit: 20,
    });
    assert(JSON.stringify(updated.data.items).includes("Updated endnote"));
    const validation = await s.call("docx_validate", { view });
    assert.equal(validation.data.passed, true);
  } finally {
    s.service.close();
    fs.rmSync(s.root, { recursive: true, force: true });
  }
});

test("DOCX textbox text is returned as bounded read-only package inventory", async () => {
  const { Document, Packer, Paragraph } = await import("docx");
  const s = setup();
  try {
    const parts = unzipSync(
      await Packer.toBuffer(
        new Document({ sections: [{ children: [new Paragraph("Body text")] }] }),
      ),
    );
    let xml = strFromU8(parts["word/document.xml"]);
    if (!xml.includes("xmlns:v="))
      xml = xml.replace(
        "<w:document ",
        '<w:document xmlns:v="urn:schemas-microsoft-com:vml" ',
      );
    xml = xml.replace(
      "<w:body>",
      '<w:body><w:p><w:r><w:pict><v:shape id="OfficeMcpTextbox"><v:textbox><w:txbxContent><w:p><w:r><w:t>Inventory-only textbox content</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>',
    );
    parts["word/document.xml"] = strToU8(xml);
    const file = path.join(s.docs, "textbox.docx");
    fs.writeFileSync(file, zipSync(parts));
    const view = (await s.call("docx_open", { path: file })).view;
    const inspection = await s.call("docx_inspect", { view });
    const story = inspection.data.storyInventory.find(
      (item) => item.type === "textbox",
    );
    assert.equal(story.inventoryReadable, true);
    assert.equal(story.mutationTargetAvailable, false);
    assert.match(story.textPreview, /Inventory-only textbox content/);
    assert.equal(inspection.data.textboxInventory.length, 1);
    const complete = await s.call("docx_read", {
      view,
      representation: "complete",
      offset: 0,
      limit: 100,
    });
    const textbox = complete.data.items.find(
      (item) => item.kind === "textboxInventory",
    );
    assert.match(textbox.text, /Inventory-only textbox content/);
    assert.equal(textbox.mutationTargetAvailable, false);
    assert.match(textbox.provenance.engine, /read-only OOXML/);
    assert.equal(complete.coverage.matrix.textboxes, "complete");
  } finally {
    s.service.close();
    fs.rmSync(s.root, { recursive: true, force: true });
  }
});

test("DOCX exact text replacement preserves run formatting and structured tables satisfy the wire schema", async () => {
  const {
    Document,
    Packer,
    Paragraph,
    TextRun,
    Table,
    TableRow,
    TableCell,
  } = await import("docx");
  const s = setup();
  try {
    const file = path.join(s.docs, "formatted-run.docx");
    fs.writeFileSync(
      file,
      await Packer.toBuffer(
        new Document({
          sections: [
            {
              children: [
                new Paragraph({
                  children: [
                    new TextRun("Payment within "),
                    new TextRun({ text: "30", bold: true }),
                    new TextRun(" days."),
                  ],
                }),
                new Table({
                  rows: [
                    new TableRow({
                      children: [
                        new TableCell({ children: [new Paragraph("Value")] }),
                      ],
                    }),
                  ],
                }),
              ],
            },
          ],
        }),
      ),
    );
    let view = (await s.call("docx_open", { path: file })).view;
    const structured = await s.call("docx_read", {
      view,
      representation: "structured",
      offset: 0,
      limit: 30,
    });
    assert(structured.data.items.some((item) => item.kind === "table" && item.rows.length === 1));
    const query = await s.call("docx_query", {
      view,
      select: { type: "text", pattern: "30", caseSensitive: true },
      require: "exactlyOne",
      offset: 0,
      limit: 2,
    });
    assert.equal(query.data.items[0].formattingSummary.bold, true);
    const paragraphTarget = structured.data.items.find(
      (item) => item.kind === "paragraph" && item.text.includes("Payment within"),
    ).target;
    const preview = await s.call("docx_apply", {
      view,
      mode: "preview",
      action: "replace_text",
      selectorTarget: paragraphTarget,
      arguments: {
        edits: [{ find: "30", replace: "45" }],
        caseSensitive: true,
      },
    });
    assert.equal(preview.data.preview, true);
    view = (
      await s.call("docx_apply", {
        view,
        operationId: "replace-formatted-run",
        action: "replace_text",
        selectorTarget: paragraphTarget,
        arguments: {
          edits: [{ find: "30", replace: "45" }],
          caseSensitive: true,
        },
      })
    ).view;
    const saved = await s.call("docx_save", {
      view,
      operationId: "save-formatted-run",
      destination: { rootId: "output", relativePath: "formatted-run-output.docx" },
    });
    const xml = strFromU8(unzipSync(fs.readFileSync(saved.data.path))["word/document.xml"]);
    assert.match(xml, /<w:b(?:\s[^>]*)?\/>[\s\S]*?<w:t[^>]*>45<\/w:t>/);
    assert(!xml.includes(">30<"));
  } finally {
    s.service.close();
    fs.rmSync(s.root, { recursive: true, force: true });
  }
});

test("DOCX qualified core actions insert, format, delete and create structures through guarded views", async () => {
  const s = setup();
  try {
    let view = (
      await s.call("docx_create", {
        operationId: "qualified-core-create",
        document: {
          format: "docx",
          blocks: [
            { kind: "paragraph", text: "Anchor paragraph" },
            { kind: "paragraph", text: "Comment target" },
          ],
        },
      })
    ).view;
    let anchor = await s.call("docx_query", {
      view,
      select: { type: "text", pattern: "Anchor paragraph" },
      require: "exactlyOne",
      offset: 0,
      limit: 2,
    });
    view = (
      await s.call("docx_apply", {
        view,
        operationId: "qualified-insert",
        action: "insert_paragraphs",
        selectorTarget: anchor.data.items[0].target,
        arguments: {
          texts: ["Inserted paragraph", "Second inserted paragraph"],
          placement: { at: "after" },
        },
      })
    ).view;
    let inserted = await s.call("docx_query", {
      view,
      select: { type: "text", pattern: "Inserted paragraph" },
      require: "exactlyOne",
      offset: 0,
      limit: 2,
    });
    view = (
      await s.call("docx_apply", {
        view,
        operationId: "qualified-format",
        action: "format_text",
        selectorTarget: inserted.data.items[0].target,
        arguments: { targetText: "Inserted paragraph", bold: true },
      })
    ).view;
    inserted = await s.call("docx_query", {
      view,
      select: { type: "text", pattern: "Inserted paragraph" },
      require: "exactlyOne",
      offset: 0,
      limit: 2,
    });
    assert.equal(inserted.data.items[0].formattingSummary.bold, true);
    view = (
      await s.call("docx_apply", {
        view,
        operationId: "qualified-delete-text",
        action: "delete_text",
        selectorTarget: inserted.data.items[0].target,
        arguments: { finds: ["Inserted "] },
      })
    ).view;
    view = (
      await s.call("docx_apply", {
        view,
        operationId: "qualified-table",
        action: "create_table",
        arguments: {
          rows: 2,
          columns: 2,
          cellTexts: [["Name", "Value"], ["Alpha", "1"]],
        },
      })
    ).view;
    const table = await s.call("docx_read", {
      view,
      representation: "domain",
      domain: "tables",
      offset: 0,
      limit: 10,
    });
    assert.equal(table.data.total, 1);
    assert(JSON.stringify(table.data.items).includes("Alpha"));
    const existingComments = await s.call("docx_read", { view, representation: "domain", domain: "comments", limit: 20 });
    assert.equal(existingComments.data.total, 0);
    const validation = await s.call("docx_validate", { view });
    assert.equal(validation.data.passed, true);
  } finally {
    s.service.close();
    fs.rmSync(s.root, { recursive: true, force: true });
  }
});

test("DOCX captions and cross-references use typed targets and survive validation", async () => {
  const s = setup();
  try {
    let view = (
      await s.call("docx_create", {
        operationId: "caption-create",
        document: {
          format: "docx",
          blocks: [
            { kind: "paragraph", text: "Diagram placeholder" },
            { kind: "paragraph", text: "See the referenced diagram." },
          ],
        },
      })
    ).view;
    const diagram = await s.call("docx_query", {
      view,
      select: { type: "text", pattern: "Diagram placeholder" },
      require: "exactlyOne",
      offset: 0,
      limit: 2,
    });
    view = (
      await s.call("docx_apply", {
        view,
        operationId: "caption-insert",
        action: "insert_caption",
        selectorTarget: diagram.data.items[0].target,
        arguments: {
          position: "below",
          label: "Figure",
          text: "System overview",
        },
      })
    ).view;
    const captions = await s.call("docx_read", {
      view,
      representation: "domain",
      domain: "captions",
      offset: 0,
      limit: 20,
    });
    assert.equal(captions.data.total, 1);
    assert(captions.data.items[0].target);
    const insertion = await s.call("docx_query", {
      view,
      select: { type: "text", pattern: "referenced diagram" },
      require: "exactlyOne",
      offset: 0,
      limit: 2,
    });
    view = (
      await s.call("docx_apply", {
        view,
        operationId: "cross-reference-insert",
        action: "insert_cross_reference",
        selectorTarget: insertion.data.items[0].target,
        arguments: {
          reference: {
            kind: "caption",
            nodeId: captions.data.items[0].itemId,
          },
          display: "labelAndNumber",
        },
      })
    ).view;
    const references = await s.call("docx_read", {
      view,
      representation: "domain",
      domain: "cross_references",
      offset: 0,
      limit: 20,
    });
    assert(references.data.total > 0);
    assert(references.data.items[0].target);
    const validation = await s.call("docx_validate", { view });
    assert.equal(validation.data.passed, true);
  } finally {
    s.service.close();
    fs.rmSync(s.root, { recursive: true, force: true });
  }
});

test("DOCX qualified image deletion preserves its surrounding content", async () => {
  const s = setup();
  try {
    let view = (
      await s.call("docx_create", {
        operationId: "qualified-removals-create",
        document: {
          format: "docx",
          blocks: [{ kind: "paragraph", text: "Linked text and image anchor" }],
        },
      })
    ).view;
    const anchor = (
      await s.call("docx_query", {
        view,
        select: { type: "text", pattern: "image anchor" },
        require: "exactlyOne",
        offset: 0,
        limit: 2,
      })
    ).data.items[0].target;
    view = (
      await s.call("docx_apply", {
        view,
        operationId: "qualified-insert-image",
        action: "insert_image",
        selectorTarget: anchor,
        arguments: { path: s.imagePath, placement: "after", alt: "Temporary" },
      })
    ).view;
    const image = (
      await s.call("docx_read", {
        view,
        representation: "domain",
        domain: "images",
        offset: 0,
        limit: 10,
      })
    ).data.items[0];
    view = (
      await s.call("docx_apply", {
        view,
        operationId: "qualified-delete-image",
        action: "delete_image",
        selectorTarget: image.target,
        arguments: {},
      })
    ).view;
    const imagesAfter = await s.call("docx_read", {
      view,
      representation: "domain",
      domain: "images",
      offset: 0,
      limit: 10,
    });
    assert.equal(imagesAfter.data.total, 0);
    assert.equal((await s.call("docx_validate", { view })).data.passed, true);
  } finally {
    s.service.close();
    fs.rmSync(s.root, { recursive: true, force: true });
  }
});

for (const action of ["update_caption", "remove_caption"])
  test(`DOCX qualified caption action: ${action}`, async () => {
    const s = setup();
    try {
      let view = (
        await s.call("docx_create", {
          operationId: `${action}-create`,
          document: {
            format: "docx",
            blocks: [{ kind: "paragraph", text: "Caption anchor" }],
          },
        })
      ).view;
      const anchor = (
        await s.call("docx_query", {
          view,
          select: { type: "text", pattern: "Caption anchor" },
          require: "exactlyOne",
          offset: 0,
          limit: 2,
        })
      ).data.items[0].target;
      view = (
        await s.call("docx_apply", {
          view,
          operationId: `${action}-insert`,
          action: "insert_caption",
          selectorTarget: anchor,
          arguments: { position: "below", label: "Figure", text: "Before" },
        })
      ).view;
      const caption = (
        await s.call("docx_read", {
          view,
          representation: "domain",
          domain: "captions",
          offset: 0,
          limit: 10,
        })
      ).data.items[0];
      view = (
        await s.call("docx_apply", {
          view,
          operationId: `${action}-apply`,
          action,
          selectorTarget: caption.target,
          arguments: action === "update_caption" ? { text: "After" } : {},
        })
      ).view;
      const captionsAfter = await s.call("docx_read", {
        view,
        representation: "domain",
        domain: "captions",
        offset: 0,
        limit: 10,
      });
      if (action === "update_caption")
        assert(JSON.stringify(captionsAfter.data.items).includes("After"));
      else assert.equal(captionsAfter.data.total, 0);
      assert.equal((await s.call("docx_validate", { view })).data.passed, true);
    } finally {
      s.service.close();
      fs.rmSync(s.root, { recursive: true, force: true });
    }
  });

test(
  "DOCX agent exports a semantically validated candidate to PDF through LibreOffice",
  { skip: !process.env.OFFICEMCP_LIBREOFFICE },
  async () => {
    const s = setup({ libreOffice: process.env.OFFICEMCP_LIBREOFFICE });
    try {
      const created = await s.call("docx_create", {
        operationId: "export-semantic-create",
        document: {
          format: "docx",
          blocks: [{ kind: "paragraph", text: "Original export text" }],
        },
      });
      const target = await s.call("docx_query", {
        view: created.view,
        select: { type: "text", pattern: "Original export text" },
        require: "exactlyOne",
      });
      const changed = await s.call("docx_apply", {
        view: created.view,
        operationId: "export-semantic-change",
        action: "rewrite_block",
        selectorTarget: target.data.items[0].target,
        arguments: { text: "Changed export text" },
      });
      const compared = await s.call("docx_compare", { view: changed.view });
      assert.equal(compared.data.hasChanges, true);
      assert(compared.data.changedComponents.includes("body"));
      assert.equal(compared.data.visualEvidence, undefined);
      const validated = await s.call("docx_validate", { view: changed.view });
      assert.equal(validated.data.passed, true);
      assert.equal(validated.data.visualEvidence, undefined);
      const exported = await s.call("docx_export", {
        view: changed.view,
        operationId: "export-semantic-pdf",
        format: "pdf",
        destination: { rootId: "output", relativePath: "exported.pdf" },
      });
      assert.equal(exported.data.published, true);
      assert.equal(exported.data.format, "pdf");
      assert.equal(fs.readFileSync(exported.data.path).subarray(0, 5).toString(), "%PDF-");
      assert.match(exported.data.engine, /LibreOffice/);
      const recoveredExport = await s.call("docx_operation_get", {
        operationId: "export-semantic-pdf",
      });
      assert.equal(recoveredExport.data.path, exported.data.path);
    } finally {
      s.close();
    }
  },
);

test("DOCX PDF export is durable, recoverable and never overwrites", async () => {
  const converterRoot = fs.mkdtempSync("/tmp/docx-export-converter-");
  const converter = path.join(converterRoot, "fake-libreoffice.cjs");
  const { PDFDocument } = await import("pdf-lib");
  const pdf = await PDFDocument.create();
  pdf.addPage([612, 792]);
  const pdfBase64 = Buffer.from(await pdf.save()).toString("base64");
  fs.writeFileSync(
    converter,
    `#!/usr/bin/env node
const fs = require("node:fs");
const path = require("node:path");
const args = process.argv.slice(2);
const outdir = args[args.indexOf("--outdir") + 1];
const input = args.at(-1);
fs.writeFileSync(path.join(outdir, path.basename(input, path.extname(input)) + ".pdf"), Buffer.from(${JSON.stringify(pdfBase64)}, "base64"));
`,
    { mode: 0o700 },
  );
  const s = setup({ libreOffice: converter });
  try {
    const created = await s.call("docx_create", {
      operationId: "export-create",
      document: {
        format: "docx",
        blocks: [{ kind: "paragraph", text: "Export me" }],
      },
    });
    const request = {
      view: created.view,
      operationId: "export-operation",
      format: "pdf",
      destination: { rootId: "output", relativePath: "exported.pdf" },
    };
    const exported = await s.call("docx_export", request);
    assert.equal(exported.data.published, true);
    assert.equal(exported.data.format, "pdf");
    assert.equal(
      fs.readFileSync(exported.data.path).subarray(0, 5).toString(),
      "%PDF-",
    );
    const replay = await s.call("docx_export", request);
    assert.deepEqual(replay.data, exported.data);
    const recovered = await s.call("docx_operation_get", {
      operationId: "export-operation",
    });
    assert.equal(recovered.data.path, exported.data.path);
    const overwrite = await s.service.call("docx_export", {
      ...request,
      operationId: "export-overwrite",
    });
    assert.equal(overwrite.payload.error.code, "DESTINATION_CONFLICT");
  } finally {
    s.service.close();
    fs.rmSync(s.root, { recursive: true, force: true });
    fs.rmSync(converterRoot, { recursive: true, force: true });
  }
});

test("DOCX agent facade is exposed over modern MCP with accurate risk annotations", async () => {
  const s = setup();
  let client;
  try {
    const created = await s.call("docx_create", {
      operationId: "wire-create",
      document: {
        format: "docx",
        blocks: [{ kind: "paragraph", text: "Wire contract" }],
      },
    });
    const saved = await s.call("docx_save", {
      view: created.view,
      operationId: "wire-save",
      destination: { rootId: "output", relativePath: "wire.docx" },
    });
    s.service.close();

    client = new Client(
      { name: "docx-agent-test", version: "1" },
      {
        capabilities: {},
        versionNegotiation: { mode: { pin: "2026-07-28" } },
      },
    );
    const transport = new StdioClientTransport({
      command: process.execPath,
      args: [
        new URL("../src/cli.mjs", import.meta.url).pathname,
        "serve",
        "--config",
        path.join(s.root, "config.json"),
      ],
      stderr: "pipe",
    });
    transport.stderr?.resume();
    await client.connect(transport);
    const listing = await client.listTools();
    assert.equal(listing.tools.length, 14);
    assert.equal(
      listing.tools.find((tool) => tool.name === "docx_read").annotations
        .readOnlyHint,
      true,
    );
    assert.equal(
      listing.tools.find((tool) => tool.name === "docx_apply").annotations
        .readOnlyHint,
      false,
    );
    assert.equal(
      listing.tools.find((tool) => tool.name === "docx_export").annotations
        .readOnlyHint,
      false,
    );
    const opened = await client.callTool({
      name: "docx_open",
      arguments: { path: saved.data.path },
    });
    assert.equal(opened.isError, false);
    assert.equal(opened.structuredContent.contractVersion, "docx/6");
    const inspected = await client.callTool({
      name: "docx_inspect",
      arguments: { view: opened.structuredContent.view },
    });
    assert.equal(inspected.isError, false);
    assert(inspected.structuredContent.data.readingPlan.length > 0);
  } finally {
    await client?.close();
    s.service.close();
    fs.rmSync(s.root, { recursive: true, force: true });
  }
});
