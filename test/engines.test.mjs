import test from "./support/test.mjs";
import assert from "node:assert/strict";
import * as fs from "node:fs";
import path from "node:path";
import { digest } from "../src/infrastructure/files.mjs";
import { writeFixtures } from "./fixtures.mjs";
import { createServiceFixture } from "./service-fixture.mjs";
import { PDFDocument } from "pdf-lib";

function setup(engines = {}) {
  return createServiceFixture({
    prefix: "office-engines-",
    engines,
    limits: { workerTimeoutMs: 60000 },
    prepare: writeFixtures,
  });
}

test("SuperDoc reads, replaces across runs, formats and inserts; saved candidate remains guarded", async () => {
  const s = setup();
  try {
    // Fixture filename is discovered so the test stays independent of helpers.
    const source = fs.readdirSync(s.docs).find((n) => n.endsWith(".docx"));
    const d = (await s.call("office_open", { path: path.join(s.docs, source) }))
      .document;
    const before = digest(fs.readFileSync(path.join(s.docs, source)));
    const read = (
      await s.call("docx_structure", {
        documentId: d.id,
        expectedRevision: d.sourceRevision,
      })
    ).data;
    assert(read.nodes.length >= 2);
    const base = {
      documentId: d.id,
      expectedSourceRevision: d.sourceRevision,
      paragraphIndex: 0,
      preservationProfile: "docx.superdoc.v1",
    };
    const edited = (
      await s.call("docx_edit_paragraph", {
        ...base,
        operationId: "replace",
        expectedText: read.nodes[0].text,
        replacement: "Payment within 45 days",
      })
    ).data;
    const formatted = (
      await s.call("docx_format_text", {
        ...base,
        candidateId: edited.candidateId,
        expectedCandidateRevision: edited.candidateRevision,
        operationId: "format",
        expectedText: "Payment within 45 days",
        start: 0,
        end: 7,
        format: { bold: true, italic: true, underline: true, fontSize: 18, fontFamily: "Arial", color: "#123456" },
      })
    ).data;
    const crossRun = (await s.call("docx_edit_paragraph", {
      ...base, candidateId: edited.candidateId, expectedCandidateRevision: formatted.candidateRevision,
      operationId: "cross-run", expectedText: "Payment within 45 days", replacement: "Payment within 60 days",
    })).data;
    const inserted = (
      await s.call("docx_insert_paragraph", {
        ...base,
        candidateId: edited.candidateId,
        expectedCandidateRevision: crossRun.candidateRevision,
        operationId: "insert",
        expectedText: "Payment within 60 days",
        text: "Please retain this new paragraph.",
        placement: "after",
      })
    ).data;
    const final = (
      await s.call("docx_structure", {
        documentId: d.id,
        candidateId: edited.candidateId,
        expectedRevision: inserted.candidateRevision,
      })
    ).data;
    assert.equal(final.nodes[1].text, "Please retain this new paragraph.");
    assert.equal(final.nodes[2].text, read.nodes[1].text);
    const v = (
      await s.call("office_validate", {
        documentId: d.id,
        candidateId: edited.candidateId,
        expectedCandidateRevision: inserted.candidateRevision,
      })
    ).data;
    assert.equal(v.passed, true);
    assert(v.notVerified.includes("visual-layout"));
    const bad = await s.service.call("docx_edit_paragraph", {
      ...base,
      operationId: "stale-target",
      expectedText: "Wrong",
      replacement: "No",
    });
    assert.equal(bad.payload.error.code, "TARGET_CONFLICT");
    assert.equal(digest(fs.readFileSync(path.join(s.docs, source))), before);
  } finally {
    s.close();
  }
});

test("PDF page selection supports duplication and rotation with deterministic validation", async () => {
  const s = setup();
  try {
    const d = (
      await s.call("office_open", { path: path.join(s.docs, "invoice.pdf") })
    ).document;
    const c = (
      await s.call("pdf_edit_pages", {
        documentId: d.id,
        expectedSourceRevision: d.sourceRevision,
        operationId: "pages",
        preservationProfile: "pdf.pages.v1",
        pages: [
          { page: 1, rotation: 90 },
          { page: 1, rotation: 0 },
        ],
      })
    ).data;
    const v = (
      await s.call("office_validate", {
        documentId: d.id,
        candidateId: c.candidateId,
        expectedCandidateRevision: c.candidateRevision,
      })
    ).data;
    const saved = (
      await s.call("office_commit", {
        documentId: d.id,
        expectedSourceRevision: d.sourceRevision,
        candidateId: c.candidateId,
        expectedCandidateRevision: c.candidateRevision,
        validationId: v.validationId,
        operationId: "publish",
        mode: "save_as",
        overwrite: false,
        destination: { rootId: "output", relativePath: "selected.pdf" },
      })
    ).data;
    const pdf = await PDFDocument.load(fs.readFileSync(saved.path));
    assert.equal(pdf.getPageCount(), 2);
    assert.equal(pdf.getPage(0).getRotation().angle, 90);
  } finally {
    s.close();
  }
});

test("Aspose discovery and license gating do not silently create evaluation outputs", async () => {
  const s = setup(
    process.env.OFFICEMCP_ASPOSE_PYTHON
      ? { python: process.env.OFFICEMCP_ASPOSE_PYTHON }
      : {},
  );
  try {
    const status = (await s.call("office_engine_status", {})).data;
    assert.equal(status.superdoc.version, "2.2.1");
    const fixture = path.join(s.docs, "book.xlsx");
    fs.copyFileSync(
      new URL("./corpus/producer.xlsx", import.meta.url),
      fixture,
    );
    const d = (await s.call("office_open", { path: fixture })).document;
    const result = await s.service.call("xlsx_recalculate", {
      documentId: d.id,
      expectedSourceRevision: d.sourceRevision,
      operationId: "no-license",
      preservationProfile: "xlsx.aspose.v1",
    });
    assert.equal(
      result.payload.error.code,
      process.env.OFFICEMCP_ASPOSE_PYTHON
        ? "ENGINE_LICENSE_REQUIRED"
        : "ENGINE_UNAVAILABLE",
    );
    assert.equal(s.service.rows("candidate").length, 0);
  } finally {
    s.close();
  }
});

test(
  "Aspose edits blanks and formatted cells in a chart/table workbook, preserves literal formulas and validates",
  { skip: !process.env.OFFICEMCP_ASPOSE_PYTHON },
  async () => {
    const s = setup({
      python: process.env.OFFICEMCP_ASPOSE_PYTHON,
      allowEvaluation: true,
    });
    try {
      const file = path.join(s.docs, "producer.xlsx");
      fs.copyFileSync(new URL("./corpus/producer.xlsx", import.meta.url), file);
      const d = (await s.call("office_open", { path: file })).document;
      const sheets = (
        await s.call("xlsx_inspect", {
          documentId: d.id,
          expectedRevision: d.sourceRevision,
        })
      ).data.items;
      const sheet = sheets[0].name;
      const c = (
        await s.call("xlsx_edit_range", {
          documentId: d.id,
          expectedSourceRevision: d.sourceRevision,
          operationId: "styled-blank",
          sheet,
          cells: [
            {
              address: "H20",
              expectedValue: null,
              expectedFormula: null,
              value: "=1+1",
              format: { bold: true, wrapText: true },
            },
            {
              address: "I20",
              expectedValue: null,
              expectedFormula: null,
              value: 125.25,
              format: { numberFormat: "0.00" },
            },
          ],
          preservationProfile: "xlsx.aspose.v1",
        })
      ).data;
      assert.equal(c.evaluation, true);
      const read = (
        await s.call("xlsx_engine_read", {
          documentId: d.id,
          candidateId: c.candidateId,
          expectedRevision: c.candidateRevision,
          sheet,
          range: "H20:I20",
        })
      ).data;
      assert.equal(read.cells[0].value, "=1+1");
      assert.equal(read.cells[0].formula, null);
      assert.equal(read.cells[0].format.bold, true);
      assert.equal(read.cells[1].value, 125.25);
      const v = (
        await s.call("office_validate", {
          documentId: d.id,
          candidateId: c.candidateId,
          expectedCandidateRevision: c.candidateRevision,
        })
      ).data;
      assert.equal(v.passed, true);
      assert.equal(v.evaluation, true);
      const pptx = path.join(s.docs, "producer.pptx");
      fs.copyFileSync(new URL("./corpus/producer.pptx", import.meta.url), pptx);
      const p = (await s.call("office_open", { path: pptx })).document;
      const blocked = await s.service.call("pptx_set_notes", {
        documentId: p.id,
        expectedSourceRevision: p.sourceRevision,
        operationId: "unlicensed-notes",
        slide: 1,
        expectedText: "",
        text: "New notes",
        preservationProfile: "pptx.aspose.v1",
      });
      assert.equal(blocked.payload.error.code, "ENGINE_LICENSE_REQUIRED");
    } finally {
      s.close();
    }
  },
);
