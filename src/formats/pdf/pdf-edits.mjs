import { writeFileSync } from "node:fs";
import { PDFDocument, PDFName, degrees } from "pdf-lib";
import { digest } from "../../infrastructure/files.mjs";
import { requireThat } from "../../infrastructure/errors.mjs";
import { verifyPdfReplay } from "./pdf-replay.mjs";

async function build(bytes, args) {
  const source = await PDFDocument.load(bytes, { updateMetadata: false });
  for (const key of [
    "AcroForm",
    "Outlines",
    "Names",
    "Perms",
    "StructTreeRoot",
    "OCProperties",
  ])
    requireThat(
      !source.catalog.has(PDFName.of(key)),
      "UNSUPPORTED_FEATURE",
      `Page selection cannot preserve PDF ${key}; no edit was made.`,
    );
  for (const page of source.getPages())
    requireThat(
      !page.node.has(PDFName.of("Annots")),
      "UNSUPPORTED_FEATURE",
      "Page selection does not preserve annotation destinations; no edit was made.",
    );
  requireThat(
    args.pages.every((x) => x.page <= source.getPageCount()),
    "TARGET_CONFLICT",
    "Requested page does not exist.",
  );
  const output = await PDFDocument.create({ updateMetadata: false });
  for (const spec of args.pages) {
    const [page] = await output.copyPages(source, [spec.page - 1]);
    page.setRotation(degrees((page.getRotation().angle + spec.rotation) % 360));
    output.addPage(page);
  }
  return Buffer.from(await output.save({ useObjectStreams: false }));
}
export async function editPdf(task, bytes) {
  if (task.kind === "pdf_edit_pages") {
    const output = await build(bytes, task.args);
    writeFileSync(task.output, output, { flag: "wx", mode: 0o600 });
    return {
      revision: digest(output),
      affectedTargets: task.args.pages.map((x) => x.page),
    };
  }
  await verifyPdfReplay(task, bytes, build, "PDF differs from the declared page operations.");
  return {
    passed: true,
    revision: digest(bytes),
    profile: "pdf.pages.v1",
    checks: ["deterministic-page-operation-replay", "PDF-reopened"],
    notVerified: ["visual-layout", "document-metadata-preservation"],
  };
}
