import * as mupdf from "mupdf";
import { writeFileSync } from "node:fs";
import { readPackage } from "../opc/package-readers.mjs";
import { requireThat } from "../../infrastructure/errors.mjs";
export function extractDocument(task, bytes) {
  if (task.args.source.kind === "package_part") {
    requireThat(
      task.format !== "pdf",
      "UNSUPPORTED_FORMAT",
      "PDF attachments use pdf_attachment sources.",
    );
    return readPackage(
      {
        ...task,
        kind: "office_package_extract",
        args: { partName: task.args.source.name },
      },
      bytes,
    );
  }
  requireThat(
    task.format === "pdf",
    "UNSUPPORTED_FORMAT",
    "PDF attachment source requires a PDF view.",
  );
  const doc = new mupdf.PDFDocument(bytes);
  doc.disableJS();
  try {
    const spec = doc.getEmbeddedFiles()[task.args.source.name];
    requireThat(
      spec,
      "NOT_FOUND",
      "Attachment name is unavailable; inspect the PDF overview.",
    );
    const metadata = doc.getFilespecParams(spec),
      buffer = doc.getEmbeddedFileContents(spec);
    requireThat(buffer, "NOT_FOUND", "Attachment has no embedded bytes.");
    try {
      const content = Buffer.from(buffer.asUint8Array());
      requireThat(
        content.length <= 20 * 1024 * 1024,
        "LIMIT_EXCEEDED",
        "Attachment exceeds extraction limit.",
      );
      writeFileSync(task.output, content, { flag: "wx", mode: 0o600 });
      return {
        name: task.args.source.name,
        mediaType: metadata.mimetype || "application/octet-stream",
        size: content.length,
      };
    } finally {
      buffer.destroy();
    }
  } finally {
    doc.destroy();
  }
}
