import test from "./support/test.mjs";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { digest } from "../src/infrastructure/files.mjs";
import { createServiceFixture } from "./service-fixture.mjs";
function setup() {
  return createServiceFixture({
    prefix: "officemcp-review-",
    limits: { workerTimeoutMs: 30000 },
  });
}
test("SuperDoc detail domains read populated footnotes, hyperlinks, fields, images and section/header references", async () => {
  const {
    Document,
    Packer,
    Paragraph,
    TextRun,
    Header,
    Footer,
    ExternalHyperlink,
    FootnoteReferenceRun,
    SimpleField,
    ImageRun,
    Table,
    TableRow,
    TableCell,
  } = await import("docx");
  const { imageBytes } = await import("./four-format-fixtures.mjs");
  const s = setup();
  try {
    const bytes = await Packer.toBuffer(
      new Document({
        footnotes: {
          1: { children: [new Paragraph("Footnote qualification.")] },
        },
        sections: [
          {
            headers: {
              default: new Header({
                children: [new Paragraph("Confidential header")],
              }),
            },
            footers: {
              default: new Footer({
                children: [new Paragraph("Liability cap applies.")],
              }),
            },
            children: [
              new Paragraph({
                children: [
                  new TextRun("Payment within 30 days."),
                  new FootnoteReferenceRun(1),
                ],
              }),
              new Paragraph({
                children: [
                  new ExternalHyperlink({
                    link: "https://example.invalid/terms",
                    children: [new TextRun("Terms link")],
                  }),
                ],
              }),
              new Paragraph({ children: [new SimpleField("PAGE")] }),
              new Paragraph({
                children: [
                  new ImageRun({
                    data: imageBytes,
                    type: "png",
                    transformation: { width: 20, height: 20 },
                  }),
                ],
              }),
              new Table({
                rows: [
                  new TableRow({
                    children: [
                      new TableCell({ children: [new Paragraph("Item")] }),
                      new TableCell({ children: [new Paragraph("Amount")] }),
                    ],
                  }),
                ],
              }),
            ],
          },
        ],
      }),
    );
    const file = path.join(s.root, "docs", "rich.docx");
    fs.writeFileSync(file, bytes);
    const d = (await s.call("office_open", { path: file })).document;
    const view = { documentId: d.id, revision: d.sourceRevision };
    for (const domain of [
      "footnotes",
      "hyperlinks",
      "fields",
      "images",
      "tables",
      "headers_footers",
    ]) {
      const r = await s.call("docx_read_details", { view, domain });
      assert(r.data.total > 0, domain + " must report present features");
      if (domain === "headers_footers")
        assert(r.data.items.some((i) => i.kind === "header" && i.refId));
    }
    assert.equal(digest(fs.readFileSync(file)), digest(bytes));
  } finally {
    s.service.close();
  }
});

test("Inherited section headers resolve to the actual source part and disclose content truncation", async () => {
  const { Document, Packer, Paragraph, Header } = await import("docx"),
    s = setup();
  try {
    const file = path.join(s.root, "docs", "sections.docx");
    fs.writeFileSync(
      file,
      await Packer.toBuffer(
        new Document({
          sections: [
            {
              headers: {
                default: new Header({
                  children: [new Paragraph("Shared confidential header")],
                }),
              },
              children: [new Paragraph("First section")],
            },
            { children: [new Paragraph("Second section")] },
          ],
        }),
      ),
    );
    const d = (await s.call("office_open", { path: file })).document,
      view = { documentId: d.id, revision: d.sourceRevision };
    const r = await s.call("docx_read_details", {
      view,
      domain: "headers_footers",
      limit: 30,
    });
    const inherited = r.data.items.find(
      (i) =>
        i.sectionIndex === 1 && i.kind === "header" && i.variant === "default",
    );
    assert.equal(inherited.resolution.status, "inherited");
    assert.equal(
      inherited.resolution.resolvedFromSection.sectionId,
      "section-0",
    );
    assert.equal(inherited.source.name, "word/header1.xml");
    assert.equal(inherited.textPreview, "Shared confidential header");
    assert(inherited.contentComplete);
    const invalid = await s.service.call("office_read", {
      view,
      docxDomain: "paragraphs",
      query: "ignored?",
    });
    assert.equal(invalid.payload.error.code, "INVALID_ARGUMENT");
  } finally {
    s.service.close();
  }
});

test("Comment writers are unavailable while existing comments remain readable", async () => {
  const s = setup();
  try {
    const created = await s.call("office_create", {
      operationId: "create",
      document: {
        format: "docx",
        blocks: [{ kind: "paragraph", text: "Keep this document." }],
      },
    });
    assert(
      !s.service.manifest().some((t) => t.name.startsWith("docx_comment_")),
    );
    for (const type of [
      "comment_add",
      "comment_reply",
      "comment_update",
      "comment_status",
      "comment_delete",
    ]) {
      const named = await s.service.call("docx_" + type, {});
      assert(named.isError);
      const shared = await s.service.call("office_edit", {
        view: created.view,
        operationId: type,
        action: { type },
      });
      assert(shared.isError);
      assert.equal(shared.payload.error.code, "INVALID_ARGUMENT");
    }
    const detail = await s.service.call("office_read", {
      view: created.view,
      docxDomain: "comments",
    });
    assert(!detail.isError, JSON.stringify(detail));
    assert.deepEqual(detail.payload.data.items, []);
  } finally {
    s.service.close();
  }
});
