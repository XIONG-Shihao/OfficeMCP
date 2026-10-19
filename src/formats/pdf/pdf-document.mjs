import * as mupdf from "mupdf";
import { PDFDocument, PDFName, degrees } from "pdf-lib";
import { writeFileSync } from "node:fs";
import { digest } from "../../infrastructure/files.mjs";
import { requireThat } from "../../infrastructure/errors.mjs";
import { formFields } from "./pdf-forms.mjs";
import { ocrPage } from "./pdf-ocr.mjs";
import { verifyPdfReplay } from "./pdf-replay.mjs";
import { readWindow } from "../common/read-window.mjs";
export async function readPdf(task, bytes) {
  const doc = new mupdf.PDFDocument(bytes);
  doc.disableJS();
  try {
    const a = task.args,
      count = doc.countPages(),
      items = [],
      features = [];
    requireThat(
      !a.ocr || a.page,
      "INVALID_ARGUMENT",
      "OCR requires an explicit page number.",
    );
    requireThat(
      count <= 2000,
      "LIMIT_EXCEEDED",
      "PDF exceeds the page budget.",
    );
    const outline = doc.loadOutline() ?? [],
      files = doc.getEmbeddedFiles();
    features.push(
      { feature: "outline", count: outline.length, state: "read" },
      {
        feature: "attachments",
        count: Object.keys(files).length,
        state: "read",
        detail:
          "Names and metadata only; embedded contents are not interpreted.",
      },
    );
    if (a.scope === "overview" || !a.page) {
      items.push({
        kind: "document",
        format: "pdf",
        pageCount: count,
        outline,
      });
      for (const [name, ref] of Object.entries(files))
        items.push({
          kind: "attachment",
          name,
          ...doc.getFilespecParams(ref),
          source: { kind: "pdf_attachment", name },
        });
    } else {
      requireThat(
        a.page <= count,
        "TARGET_CONFLICT",
        "Page number does not exist.",
      );
      const page = doc.loadPage(a.page - 1),
        st = page.toStructuredText("preserve-whitespace");
      try {
        const blocks = JSON.parse(st.asJSON()).blocks ?? [];
        items.push(
          ...blocks.map((b, i) => ({
            kind: b.type ?? "text",
            page: a.page,
            block: i,
            ...b,
          })),
        );
      } finally {
        st.destroy();
      }
      for (const annotation of page.getAnnotations()) {
        items.push({
          kind: "annotation",
          page: a.page,
          type: annotation.getType(),
          bounds: annotation.getBounds(),
          contents: annotation.getContents(),
        });
        annotation.destroy();
      }
      features.push({
        feature: `page:${a.page}`,
        count: items.length,
        state: "read",
      });
      page.destroy();
    }
    if (a.ocr) {
      const words = await ocrPage(task, bytes);
      items.push(...words);
      features.push({
        feature: `OCR:page:${a.page}`,
        count: words.length,
        state: "read",
        detail:
          "OCR guesses, with confidence and top-left page-point bounds; verify important text visually.",
      });
    }
    const fields = await formFields(bytes);
    items.push(...fields);
    features.push({
      feature: "form-fields",
      count: fields.length,
      state: "read",
    });
    const window = readWindow(items, a);
    return {
      ...window,
      coverage: {
        features,
        notChecked: [
          a.ocr ? "OCR-accuracy" : "OCR",
          "table-relationships",
          "visual-reading-order",
          "unrequested-pages",
        ],
        completeDocument: false,
        windowComplete: window.nextOffset === null,
      },
      engine: "MuPDF.js 1.28.1",
    };
  } finally {
    doc.destroy();
  }
}
async function rearrange(bytes, args) {
  const doc = await PDFDocument.load(bytes, { updateMetadata: false });
  const original = doc.getPages(),
    fields = doc.getForm().getFields();
  requireThat(
    args.pages.every((p) => p.page <= original.length),
    "TARGET_CONFLICT",
    "Page does not exist.",
  );
  requireThat(
    !doc.catalog.has(PDFName.of("Perms")) &&
      !fields.some((f) => f.constructor.name === "PDFSignature"),
    "UNSUPPORTED_FEATURE",
    "Signed documents cannot be rearranged.",
  );
  const selected = new Set(args.pages.map((p) => p.page));
  const duplicated = args.pages.length !== selected.size;
  const hasAnnotations = original.some(
    (p) => (p.node.Annots()?.size() ?? 0) > 0,
  );
  // Full permutations retain original page references, including form widgets
  // and bookmark destinations. Deleting/duplicating referenced pages needs a
  // separate semantic remapping operation, not a best-effort copy.
  requireThat(
    (selected.size === original.length && !duplicated) ||
      !(
        fields.length ||
        hasAnnotations ||
        doc.catalog.has(PDFName.of("Outlines")) ||
        doc.catalog.has(PDFName.of("StructTreeRoot")) ||
        doc.catalog.has(PDFName.of("Names"))
      ),
    "UNSUPPORTED_FEATURE",
    "This PDF has document-level references. Use a complete page permutation; removal or duplication cannot yet preserve all references.",
  );
  const seen = new Set(),
    pages = [];
  for (const spec of args.pages) {
    const source = original[spec.page - 1];
    const page = seen.has(spec.page)
      ? (await doc.copyPages(doc, [spec.page - 1]))[0]
      : source;
    seen.add(spec.page);
    pages.push({
      page,
      rotation: (source.getRotation().angle + spec.rotation) % 360,
    });
  }
  while (doc.getPageCount()) doc.removePage(doc.getPageCount() - 1);
  for (const { page, rotation } of pages) {
    doc.addPage(page);
    page.setRotation(degrees(rotation));
  }
  return Buffer.from(await doc.save({ updateFieldAppearances: false }));
}
export async function editPdfDocument(task, bytes) {
  if (task.kind === "pdf_rearrange_pages") {
    const out = await rearrange(bytes, task.args);
    writeFileSync(task.output, out, { flag: "wx", mode: 0o600 });
    return {
      revision: digest(out),
      affectedTargets: task.args.pages.map((p) => p.page),
    };
  }
  await verifyPdfReplay(task, bytes, rearrange, "PDF page edit replay differs from the saved candidate.");
  return {
    passed: true,
    revision: digest(bytes),
    profile: "pdf.mupdf.v1",
    checks: ["page-operation-replay", "PDF-reopened"],
    notVerified: [
      "visual-layout",
      "all-link-and-form-semantics-after-page-removal",
    ],
  };
}
