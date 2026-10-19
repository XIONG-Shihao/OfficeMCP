import test from "./support/test.mjs";
import assert from "node:assert/strict";
import * as fs from "node:fs";
import path from "node:path";
import { createServiceFixture } from "./service-fixture.mjs";
import { PDFDocument, PDFName, PDFString } from "pdf-lib";

function setup(engines = {}) {
  return createServiceFixture({
    prefix: "office-daily-",
    engines,
    limits: { workerTimeoutMs: 60000 },
    payloadOnly: false,
  });
}
for (const document of [
  {
    format: "docx",
    header: "Confidential",
    footer: "Terms apply",
    blocks: [
      { kind: "heading", text: "Quarterly report" },
      { kind: "paragraph", text: "Payment within 30 days." },
      {
        kind: "table",
        rows: [
          ["Item", "Amount"],
          ["Services", "125"],
        ],
      },
    ],
  },
  {
    format: "xlsx",
    sheets: [
      {
        name: "Budget",
        rows: [
          ["Item", "Amount"],
          ["Services", 125],
          ["Total", { formula: "SUM(B2:B2)" }],
        ],
        header: true,
        table: false,
      },
    ],
  },
  {
    format: "pptx",
    title: "Quarterly review",
    slides: [
      {
        title: "Results",
        blocks: [
          { kind: "paragraph", text: "Revenue increased." },
          {
            kind: "table",
            rows: [
              ["Metric", "Value"],
              ["Revenue", "125"],
            ],
          },
        ],
        notes: "Explain the assumptions.",
      },
    ],
  },
  {
    format: "pdf",
    pages: [
      {
        title: "Quarterly report",
        lines: ["Revenue: 125", "Payment within 30 days."],
      },
    ],
  },
])
  test(`Create, inspect and save ${document.format} with the recommended agent workflow`, async () => {
    const s = setup();
    try {
      const created = (
        await s.call("office_create", { operationId: "create", document })
      ).payload;
      assert.equal(created.view.revision, created.document.sourceRevision);
      const same = (
        await s.call("office_create", { operationId: "create", document })
      ).payload;
      assert.deepEqual(same.view, created.view);
      const read = (
        await s.call("office_read", {
          documentId: created.view.documentId,
          candidateId: created.view.candidateId,
          ...(document.format === "xlsx"
            ? { sheet: "Budget", range: "A1:B3" }
            : document.format === "pptx"
              ? { slide: 1 }
              : document.format === "pdf"
                ? { page: 1 }
                : {}),
        })
      ).payload;
      assert(read.data.items.length > 0);
      if (document.format === "xlsx") {
        const number = read.data.items.find((c) => c.address === "B2");
        assert.equal(number.value, "125");
        assert.equal(number.expectedValue, 125);
        assert.equal(number.exactValue, "125");
      }
      assert.equal(read.coverage.completeDocument, false);
      if (document.format === "docx") {
        assert(read.data.items.some((x) => x.kind === "table"));
        assert(read.data.items.some((x) => x.text === "Terms apply"));
        const p = read.data.items.find(
          (x) => x.text === "Payment within 30 days.",
        );
        const edited = (
          await s.call("office_edit", {
            view: read.view,
            operationId: "replace",
            action: {
              type: "replace_text",
              target: p.target,
              expectedText: "30 days",
              replacement: "45 days",
            },
          })
        ).payload;
        const checked = (
          await s.call("office_read", {
            documentId: edited.view.documentId,
            candidateId: edited.view.candidateId,
          })
        ).payload;
        assert(
          checked.data.items.some((x) => x.text === "Payment within 45 days."),
        );
        created.view = edited.view;
        const formatted = (
          await s.call("office_edit", {
            view: created.view,
            operationId: "format",
            action: {
              type: "format_text",
              target: p.target,
              expectedText: "45 days",
              format: { bold: true },
            },
          })
        ).payload;
        created.view = formatted.view;
        // Pinned SDK 2.2.1 rejects this tracked range in the table-bearing fixture.
        // Assert the documented refusal and checkpoint preservation, never downgrade to direct.
        const tracked = await s.service.call("office_edit", {
          view: created.view,
          operationId: "tracked",
          action: {
            type: "replace_text",
            target: p.target,
            expectedText: "45 days",
            replacement: "60 days",
            tracked: true,
            author: "Reviewer",
          },
        });
        assert.equal(tracked.payload.error.code, "UNSUPPORTED_FEATURE");
        assert.match(tracked.payload.error.message, /SuperDoc 2.2.1 rejected/);
        assert(
          (
            await s.call("office_read", { view: created.view })
          ).payload.data.items.some(
            (i) => i.text === "Payment within 45 days.",
          ),
        );
        const table = checked.data.items.find((x) => x.kind === "table");
        created.view = (
          await s.call("office_edit", {
            view: created.view,
            operationId: "table",
            action: {
              type: "table_cell",
              target: { ...table.target, row: 1, column: 1 },
              expectedText: "125",
              text: "150",
            },
          })
        ).payload.view;
        created.view = (
          await s.call("office_edit", {
            view: created.view,
            operationId: "row",
            action: {
              type: "table_row",
              target: { ...table.target, row: 1 },
              action: "insert_after",
              values: ["Hardware", "25"],
            },
          })
        ).payload.view;
        const end = (await s.call("office_read", { view: created.view }))
          .payload;
        assert(
          end.data.items.some((x) => x.text === "Payment within 45 days."),
        );
        assert.equal(
          end.data.items.find((x) => x.kind === "table").rows.length,
          3,
        );
      }
      if (document.format === "pptx") {
        const shape = read.data.items.find(
          (x) => x.kind === "shape" && x.text === "Revenue increased.",
        );
        created.view = (
          await s.call("office_edit", {
            view: created.view,
            operationId: "text",
            action: {
              type: "replace_shape_text",
              slide: 1,
              shapeId: shape.shapeId,
              expectedText: shape.text,
              replacement: "Revenue increased by 20%.",
            },
          })
        ).payload.view;
        created.view = (
          await s.call("office_edit", {
            view: created.view,
            operationId: "position",
            action: {
              type: "move_shape",
              slide: 1,
              shapeId: shape.shapeId,
              x: 50,
              y: 100,
              width: 700,
              height: 70,
            },
          })
        ).payload.view;
        const note = read.data.items.find(
          (x) => x.kind === "note" && x.text === "Explain the assumptions.",
        );
        assert(note);
        created.view = (
          await s.call("office_edit", {
            view: created.view,
            operationId: "notes",
            action: {
              type: "set_notes",
              slide: 1,
              expectedText: note.text,
              text: "Explain the updated assumptions.",
            },
          })
        ).payload.view;
      }
      if (document.format === "pdf") {
        const edited = (
          await s.call("office_edit", {
            view: created.view,
            operationId: "pages",
            action: {
              type: "pages",
              pages: [
                { page: 1, rotation: 90 },
                { page: 1, rotation: 0 },
              ],
            },
          })
        ).payload;
        created.view = edited.view;
        const inspected = (await s.call("office_read", { view: created.view, scope: "overview" })).payload;
        assert.equal(inspected.data.items.find((item) => item.kind === "document").pageCount, 2);
      }
      const args = {
        view: created.view,
        operationId: "save",
        destination: {
          rootId: "output",
          relativePath: "result." + document.format,
        },
      };
      const saved = (await s.call("office_save", args)).payload;
      assert(fs.existsSync(saved.data.path));
      assert(saved.validation.checks.length);
      if (document.format === "pdf") {
        const reopened = await PDFDocument.load(
          fs.readFileSync(saved.data.path),
        );
        assert.deepEqual(
          reopened.getPages().map((p) => p.getRotation().angle),
          [90, 0],
        );
      }
      assert.equal(
        (await s.call("office_save", args)).payload.data.path,
        saved.data.path,
      );
    } finally {
      s.service.close();
    }
  });

test("Unicode PDF creation, form filling and bookmark-aware page order", async () => {
  const s = setup();
  try {
    const created = (
      await s.call("office_create", {
        operationId: "unicode",
        document: {
          format: "pdf",
          pages: [
            {
              title: "季度报告",
              lines: ["付款期限为三十天。", "Revenue: 125"],
            },
          ],
        },
      })
    ).payload;
    const read = (await s.call("office_read", { view: created.view, page: 1 }))
      .payload;
    assert(JSON.stringify(read.data.items).includes("季度报告"));
    const pdf = await PDFDocument.create();
    const page = pdf.addPage();
    pdf.addPage();
    const form = pdf.getForm();
    const field = form.createTextField("customer");
    field.setText("Old");
    field.addToPage(page, { x: 40, y: 600, width: 200, height: 30 });
    const check = form.createCheckBox("agree");
    check.addToPage(page, { x: 40, y: 550, width: 20, height: 20 });
    const outlines = pdf.context.obj({ Type: "Outlines", Count: 1 }),
      outlineRef = pdf.context.register(outlines),
      item = pdf.context.obj({
        Title: PDFString.of("Terms"),
        Parent: outlineRef,
        Dest: [page.ref, PDFName.of("Fit")],
      }),
      itemRef = pdf.context.register(item);
    outlines.set(PDFName.of("First"), itemRef);
    outlines.set(PDFName.of("Last"), itemRef);
    pdf.catalog.set(PDFName.of("Outlines"), outlineRef);
    const file = path.join(s.docs, "form.pdf");
    fs.writeFileSync(file, await pdf.save());
    const opened = (await s.call("office_open", { path: file })).payload
      .document;
    const initial = (await s.call("office_read", { documentId: opened.id }))
      .payload;
    assert(
      initial.data.items.some(
        (x) => x.name === "customer" && x.value === "Old",
      ),
    );
    const filled = (
      await s.call("office_edit", {
        view: initial.view,
        operationId: "fill",
        action: {
          type: "fill_form",
          fields: [
            { name: "customer", expectedValue: "Old", value: "张三" },
            { name: "agree", expectedValue: false, value: true },
          ],
        },
      })
    ).payload;
    const after = (await s.call("office_read", { view: filled.view })).payload;
    assert(
      after.data.items.some((x) => x.name === "customer" && x.value === "张三"),
    );
    assert(
      after.data.items.some((x) => x.name === "agree" && x.value === true),
    );

    const reordered = (
      await s.call("office_edit", {
        view: filled.view,
        operationId: "order",
        action: {
          type: "pages",
          pages: [
            { page: 2, rotation: 0 },
            { page: 1, rotation: 0 },
          ],
        },
      })
    ).payload;
    const ordered = (await s.call("office_read", { view: reordered.view }))
      .payload;
    assert.equal(ordered.data.items[0].outline[0].page, 1);
    assert(
      ordered.data.items.some(
        (x) => x.name === "customer" && x.value === "张三",
      ),
    );
    await s.call("office_cleanup", {});
    const phase = s.service.get("candidate", reordered.view.candidateId)
      .stages[0];
    const phaseBytes = fs.readFileSync(phase.path);
    fs.writeFileSync(phase.path, Buffer.from("tampered phase"));
    const refused = await s.service.call("office_save", {
      view: reordered.view,
      operationId: "tampered-save",
      destination: { rootId: "output", relativePath: "must-not-exist.pdf" },
    });
    assert(refused.isError);
    assert(!fs.existsSync(path.join(s.docs, "must-not-exist.pdf")));
    fs.writeFileSync(phase.path, phaseBytes);
    const saved = (
      await s.call("office_save", {
        view: reordered.view,
        operationId: "save-form",
        destination: { rootId: "output", relativePath: "filled.pdf" },
      })
    ).payload;
    assert(saved.validation.checks.includes("ordered-edit-phase-chain"));
  } finally {
    s.service.close();
  }
});

test(
  "Optional OCR exposes recognized text, confidence and coordinates",
  { skip: !process.env.OFFICEMCP_TESSERACT },
  async () => {
    const s = setup({ tesseract: process.env.OFFICEMCP_TESSERACT });
    try {
      const pdf = await PDFDocument.create();
      const page = pdf.addPage([600, 300]);
      page.drawText("Invoice total 125", { x: 40, y: 160, size: 32 });
      const file = path.join(s.docs, "ocr.pdf");
      fs.writeFileSync(file, await pdf.save());
      const d = (await s.call("office_open", { path: file })).payload.document;
      const ocr = (await s.call("office_read", {
        documentId: d.id, page: 1, ocr: true, limit: 100,
      })).payload;
      const word = ocr.data.items.find((x) => x.kind === "ocr-word" && x.text.includes("125"));
      assert(word);
      assert(word.confidence >= 0 && word.confidence <= 100);
      assert.equal(word.bounds.units, "points");
      assert(ocr.coverage.features.some((x) => x.feature.startsWith("OCR:")));
    } finally { s.service.close(); }
  },
);

test(
  "Aspose workbook formula, rows and worksheet lifecycle",
  { skip: !process.env.OFFICEMCP_ASPOSE_PYTHON },
  async () => {
    const s = setup({
      python: process.env.OFFICEMCP_ASPOSE_PYTHON,
      allowEvaluation: true,
    });
    try {
      let view = (
        await s.call("office_create", {
          operationId: "create",
          document: {
            format: "xlsx",
            sheets: [
              {
                name: "Budget",
                rows: [
                  ["Item", "Amount"],
                  ["Services", 125],
                  ["Hardware", 25],
                  ["Total", 0],
                ],
              },
            ],
          },
        })
      ).payload.view;
      view = (
        await s.call("office_edit", {
          view,
          operationId: "formula",
          action: {
            type: "set_formula",
            sheet: "Budget",
            address: "B4",
            expectedFormula: null,
            formula: "SUM(B2:B3)",
          },
        })
      ).payload.view;
      const read = (
        await s.call("office_read", { view, sheet: "Budget", range: "B4:B4" })
      ).payload;
      assert.equal(Number(read.data.items[0].cachedValue.value), 150);
      view = (
        await s.call("office_edit", {
          view,
          operationId: "row",
          action: {
            type: "rows_columns",
            sheet: "Budget",
            axis: "rows",
            action: "insert",
            index: 2,
            count: 1,
          },
        })
      ).payload.view;
      const shifted = (
        await s.call("office_read", { view, sheet: "Budget", range: "B5:B5" })
      ).payload;
      assert.equal(Number(shifted.data.items[0].cachedValue.value), 150);
      view = (
        await s.call("office_edit", {
          view,
          operationId: "sheet",
          action: { type: "worksheet", action: "add", sheet: "Summary" },
        })
      ).payload.view;
      view = (
        await s.call("office_edit", {
          view,
          operationId: "rename",
          action: {
            type: "worksheet",
            action: "rename",
            sheet: "Summary",
            name: "Overview",
          },
        })
      ).payload.view;
      const sheets = (await s.call("office_read", { view, scope: "overview" }))
        .payload;
      assert(sheets.data.items.some((x) => x.name === "Overview"));
      await s.call("office_save", {
        view,
        operationId: "save",
        destination: { rootId: "output", relativePath: "budget.xlsx" },
      });
    } finally {
      s.service.close();
    }
  },
);

test("Discover and extract PDF attachments and reuse exact images in new documents", async () => {
  const { imageBytes } = await import("./four-format-fixtures.mjs");
  const s = setup();
  try {
    const pdf = await PDFDocument.create();
    pdf.addPage();
    await pdf.attach(imageBytes, "logo.png", { mimeType: "image/png" });
    const file = path.join(s.docs, "attachment.pdf");
    fs.writeFileSync(file, await pdf.save());
    const opened = (await s.call("office_open", { path: file })).payload
      .document;
    const read = (await s.call("office_read", { documentId: opened.id }))
      .payload;
    const attachment = read.data.items.find((x) => x.kind === "attachment");
    assert.deepEqual(attachment.source, {
      kind: "pdf_attachment",
      name: "logo.png",
    });
    const extracted = await s.call("office_extract", {
      view: read.view,
      source: attachment.source,
    });
    assert.deepEqual(
      Buffer.from(extracted.content[0].data, "base64"),
      Buffer.from(imageBytes),
    );
    const imageFile = path.join(s.docs, "logo.png");
    fs.writeFileSync(imageFile, imageBytes);
    await s.call("office_create", {
      operationId: "local-image",
      document: {
        format: "pptx",
        slides: [
          {
            title: "Logo",
            blocks: [
              { kind: "image", path: imageFile, width: 100, height: 100 },
            ],
          },
        ],
      },
    });
    const created = (
      await s.call("office_create", {
        operationId: "image-doc",
        document: {
          format: "docx",
          blocks: [
            {
              kind: "image",
              assetId: extracted.payload.data.assetId,
              width: 100,
              height: 100,
              alt: "Sample logo",
            },
          ],
        },
      })
    ).payload;
    const document = (await s.call("office_read", { view: created.view }))
      .payload;
    const image = document.data.items.find((x) => x.kind === "image");
    assert.equal(image.source.kind, "sdk_image");
    const roundtrip = await s.service.call("office_extract", {
      view: created.view,
      source: { kind: "package_part", name: "word/media/image1.png" },
    });
    assert.equal(roundtrip.payload.error.code, "UNSUPPORTED_FEATURE");
  } finally {
    s.service.close();
  }
});
