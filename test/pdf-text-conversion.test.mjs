import test from "./support/test.mjs";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { tmpdir } from "node:os";
import { PDFDocument } from "pdf-lib";
import { exportDocumentPdf } from "../src/formats/common/document-export.mjs";
import { ocrPage } from "../src/formats/pdf/pdf-ocr.mjs";

const externalScript = { skip: process.platform === "win32" && "Fixture executables use a POSIX shebang" };
function fixture() {
  const root = fs.mkdtempSync(path.join(tmpdir(), "office-pdf-text-"));
  const scratch = path.join(root, "scratch");
  fs.mkdirSync(scratch);
  return { root, scratch, close: () => fs.rmSync(root, { recursive: true, force: true }) };
}
function executable(directory, name, source) {
  const file = path.join(directory, name);
  fs.writeFileSync(file, `#!${process.execPath}\n${source}`, { mode: 0o700 });
  return file;
}
async function pdfBytes(size = [100, 50]) {
  const document = await PDFDocument.create();
  document.addPage(size).drawText("125", { x: 10, y: 15, size: 14 });
  return Buffer.from(await document.save());
}

test("DOCX PDF conversion keeps document export without producing preview images", externalScript, async () => {
  const f = fixture();
  try {
    const bytes = await pdfBytes();
    const input = path.join(f.root, "source.docx"), output = path.join(f.root, "result.pdf");
    fs.writeFileSync(input, "source document bytes");
    const pdf = path.join(f.root, "converted.pdf");
    fs.writeFileSync(pdf, bytes);
    const engine = executable(f.root, "convert.mjs", `
      import fs from "node:fs"; import path from "node:path";
      const args = process.argv.slice(2);
      if (args[args.indexOf("--convert-to") + 1] !== "pdf") throw new Error("Unexpected conversion format");
      const out = args[args.indexOf("--outdir") + 1];
      fs.copyFileSync(${JSON.stringify(pdf)}, path.join(out, path.basename(args.at(-1), ".docx") + ".pdf"));
    `);
    const result = await exportDocumentPdf({ format: "docx", input, output, scratch: f.scratch,
      maxFileBytes: 1024 * 1024, engines: { libreOffice: engine } });
    assert.equal(result.format, "pdf");
    assert.equal(result.pageCount, 1);
    assert.equal(result.size, bytes.length);
    assert.deepEqual(fs.readFileSync(output), bytes);
    assert.equal(fs.readFileSync(input, "utf8"), "source document bytes");
    assert.deepEqual(fs.readdirSync(path.join(f.scratch, "export-pdf-output")), ["source.pdf"]);
    assert(!Object.hasOwn(result, "width"));
    assert(!Object.hasOwn(result, "layoutId"));
  } finally { f.close(); }
});

test("DOCX export refuses unavailable conversion engines and oversized converted PDFs", externalScript, async () => {
  const f = fixture();
  try {
    const task = { format: "docx", input: path.join(f.root, "source.docx"), output: path.join(f.root, "out.pdf"), scratch: f.scratch, maxFileBytes: 10 };
    await assert.rejects(exportDocumentPdf(task), { code: "ENGINE_UNAVAILABLE" });
    const engine = executable(f.root, "oversize.mjs", `
      import fs from "node:fs"; import path from "node:path";
      const args = process.argv.slice(2), out = args[args.indexOf("--outdir") + 1];
      fs.writeFileSync(path.join(out, "source.pdf"), Buffer.alloc(11));
    `);
    await assert.rejects(exportDocumentPdf({ ...task, engines: { libreOffice: engine } }), { code: "LIMIT_EXCEEDED" });
    assert(!fs.existsSync(task.output));
  } finally { f.close(); }
});

test("OCR returns bounded text evidence and deletes its private image intermediate", externalScript, async () => {
  const f = fixture();
  try {
    const tsv = "level\tpage_num\tblock_num\tpar_num\tline_num\tword_num\tleft\ttop\twidth\theight\tconf\ttext\n5\t1\t1\t1\t1\t1\t20\t10\t40\t20\t96\t125\n";
    const engine = executable(f.root, "ocr.mjs", `
      import fs from "node:fs";
      const png = fs.readFileSync(process.argv[2]);
      if (png.readUInt32BE(16) !== 200 || png.readUInt32BE(20) !== 100) throw new Error("Unexpected OCR input size");
      process.stdout.write(${JSON.stringify(tsv)});
    `);
    const result = await ocrPage({ scratch: f.scratch, engines: { tesseract: engine }, args: { page: 1 } }, await pdfBytes());
    assert.deepEqual(result, [{ kind: "ocr-word", page: 1, text: "125", confidence: 96,
      bounds: { x: 10, y: 5, width: 20, height: 10, units: "points", origin: "top-left" } }]);
    assert.deepEqual(fs.readdirSync(f.scratch), []);
    const failed = executable(f.root, "failed.mjs", "process.exit(1);\n");
    await assert.rejects(ocrPage({ scratch: f.scratch, engines: { tesseract: failed }, args: { page: 1 } }, await pdfBytes()), { code: "ENGINE_FAILED" });
    assert.deepEqual(fs.readdirSync(f.scratch), []);
    await assert.rejects(ocrPage({ scratch: f.scratch, engines: { tesseract: engine }, args: { page: 1 } }, await pdfBytes([5000, 5000])), { code: "LIMIT_EXCEEDED" });
    assert.deepEqual(fs.readdirSync(f.scratch), []);
  } finally { f.close(); }
});
