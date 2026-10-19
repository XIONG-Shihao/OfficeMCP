import * as mupdf from "mupdf";
import { writeFileSync, rmSync } from "node:fs";
import path from "node:path";
import { requireThat } from "../../infrastructure/errors.mjs";
import { runLocal } from "../../engines/local-process.mjs";

// Tesseract consumes pixels. This bounded intermediate is private scratch data;
// the MCP result contains only recognized words, confidence and coordinates.
function writeOcrInput(bytes, pageNumber, output) {
  const document = new mupdf.PDFDocument(bytes);
  try {
    document.disableJS();
    requireThat(pageNumber >= 1 && pageNumber <= document.countPages(), "TARGET_CONFLICT", "Page does not exist.");
    const page = document.loadPage(pageNumber - 1);
    try {
      const [x0, y0, x1, y1] = page.getBounds();
      requireThat(Math.ceil((x1 - x0) * 2) * Math.ceil((y1 - y0) * 2) <= 8_000_000,
        "LIMIT_EXCEEDED", "OCR page exceeds pixel budget.");
      const pixels = page.toPixmap(mupdf.Matrix.scale(2, 2), mupdf.ColorSpace.DeviceRGB, false, true);
      try {
        const png = pixels.asPNG();
        requireThat(png.length <= 8 * 1024 * 1024, "LIMIT_EXCEEDED", "OCR image exceeds its scratch budget.");
        writeFileSync(output, png, { flag: "wx", mode: 0o600 });
      } finally { pixels.destroy(); }
    } finally { page.destroy(); }
  } finally { document.destroy(); }
}

export async function ocrPage(task, bytes) {
  requireThat(task.engines?.tesseract, "ENGINE_UNAVAILABLE",
    "Configure engines.tesseract for OCR; ordinary PDF text extraction remains available.");
  const input = path.join(task.scratch, "ocr.png");
  try {
    writeOcrInput(bytes, task.args.page, input);
    const tsv = await runLocal(task.engines.tesseract, [input, "stdout", "-l", task.args.ocrLanguage ?? "eng", "tsv"]);
    return tsv.split(/\r?\n/).slice(1).filter(Boolean).map((line) => line.split("\t"))
      .filter((cells) => cells.length >= 12 && Number(cells[10]) >= 0 && cells.slice(11).join("\t").trim())
      .map((cells) => ({ kind: "ocr-word", page: task.args.page, text: cells.slice(11).join("\t"),
        confidence: Number(cells[10]), bounds: { x: Number(cells[6]) / 2, y: Number(cells[7]) / 2,
          width: Number(cells[8]) / 2, height: Number(cells[9]) / 2, units: "points", origin: "top-left" } }));
  } finally { rmSync(input, { force: true }); }
}
