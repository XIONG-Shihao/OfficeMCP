import { writeFileSync, readFileSync } from "node:fs";
import ExcelJS from "exceljs";
import PptxGenJS from "pptxgenjs";
import { PDFDocument, rgb } from "pdf-lib";
import { digest } from "../../infrastructure/files.mjs";
import { requireThat } from "../../infrastructure/errors.mjs";
import { unpack } from "../opc/ooxml.mjs";
import { documentFonts } from "../pdf/pdf-fonts.mjs";
import { createCsvWorkbook } from "#document-core/xlsx/csv-document.mjs";

function picture(block, task) {
  const source = task.assets?.[block.assetId ?? block.path];
  requireThat(source, "ACCESS_DENIED", "Image must be an owned image asset.");
  return readFileSync(source);
}
export async function createDocument(task) {
  const spec = task.args.document;
  requireThat(
    spec.format !== "docx",
    "UNSUPPORTED_FEATURE",
    "DOCX creation is SDK-only.",
  );
  let bytes;
  if (spec.format === "csv") {
    bytes = createCsvWorkbook(spec.rows);
  } else if (spec.format === "xlsx") {
    const book = new ExcelJS.Workbook();
    book.creator = "OfficeMCP";
    book.calcProperties.fullCalcOnLoad = true;
    requireThat(
      new Set(spec.sheets.map((s) => s.name.toLowerCase())).size ===
        spec.sheets.length,
      "INVALID_ARGUMENT",
      "Worksheet names must be unique.",
    );
    for (const [index, s] of spec.sheets.entries()) {
      requireThat(
        !/[\\/*?:\[\]]/.test(s.name),
        "INVALID_ARGUMENT",
        "Worksheet name contains invalid characters.",
      );
      const sheet = book.addWorksheet(s.name);
      for (const row of s.rows)
        sheet.addRow(
          row.map((v) =>
            v && typeof v === "object"
              ? { formula: v.formula.replace(/^=/, "") }
              : v,
          ),
        );
      if (s.header && s.rows.length) {
        sheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
        sheet.getRow(1).fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FF17365D" },
        };
        sheet.views = [{ state: "frozen", ySplit: 1 }];
      }
      sheet.columns.forEach((c, i) => {
        c.width = s.columnWidths?.[i] ?? 18;
      });
      if (s.table) {
        const names = s.rows[0];
        requireThat(
          names?.every((n) => typeof n === "string" && n.length > 0) &&
            new Set(names).size === names.length,
          "INVALID_ARGUMENT",
          "Table headers must be nonempty unique strings.",
        );
        requireThat(
          s.rows.every((r) => r.length === names.length),
          "INVALID_ARGUMENT",
          "Table rows must have the same width.",
        );
        sheet.addTable({
          name: `Table${index + 1}`,
          ref: "A1",
          headerRow: true,
          columns: names.map((name) => ({ name })),
          rows: s.rows.slice(1),
          style: { theme: "TableStyleMedium2", showRowStripes: true },
        });
      }
    }
    bytes = Buffer.from(await book.xlsx.writeBuffer());
  } else if (spec.format === "pptx") {
    const deck = new PptxGenJS();
    deck.layout = "LAYOUT_WIDE";
    deck.title = spec.title;
    deck.author = "OfficeMCP";
    deck.subject = spec.title;
    deck.lang = "en-US";
    for (const s of spec.slides) {
      const slide = deck.addSlide();
      slide.background = { color: "FFFFFF" };
      slide.addText(s.title, {
        x: 0.55,
        y: 0.35,
        w: 12.2,
        h: 0.7,
        fontFace: "Arial",
        fontSize: 30,
        bold: true,
        color: "17365D",
        margin: 0,
        fit: "shrink",
      });
      let y = 1.3;
      for (const b of s.blocks) {
        if (b.kind === "image") {
          const h = b.height / 72;
          requireThat(
            y + h <= 7,
            "LAYOUT_OVERFLOW",
            "Slide image exceeds usable height; split content across slides.",
          );
          slide.addImage({
            data:
              "image/" +
              (picture(b, task)[0] === 0x89 ? "png" : "jpeg") +
              ";base64," +
              picture(b, task).toString("base64"),
            x: 0.6,
            y,
            w: b.width / 72,
            h,
            altText: b.alt,
          });
          y += h + 0.2;
        } else if (b.kind === "table") {
          const h = b.rows.length * 0.38;
          requireThat(
            y + h <= 7 && b.rows.every((r) => r.length === b.rows[0].length),
            "LAYOUT_OVERFLOW",
            "Table must be rectangular and fit on one slide.",
          );
          slide.addTable(b.rows, {
            x: 0.6,
            y,
            w: 12.1,
            rowH: 0.38,
            fontSize: 15,
            border: { color: "D5DDE8", pt: 1 },
            color: "17365D",
            margin: 0.06,
            autoPage: false,
          });
          y += h + 0.2;
        } else {
          const h = Math.max(0.5, Math.ceil(b.text.length / 85) * 0.35);
          requireThat(
            y + h <= 7,
            "LAYOUT_OVERFLOW",
            "Slide text exceeds usable height; split content across slides.",
          );
          slide.addText(b.text, {
            x: 0.65,
            y,
            w: 12,
            h,
            fontFace: "Arial",
            fontSize: b.style?.fontSize ?? (b.kind === "heading" ? 24 : 20),
            bold: b.kind === "heading" || b.style?.bold,
            italic: b.style?.italic,
            color: b.style?.color ?? "233044",
            bullet: b.kind === "bullet" ? { indent: 18 } : undefined,
            margin: 0,
            breakLine: false,
            fit: "shrink",
          });
          y += h + 0.15;
        }
      }
      if (s.notes) slide.addNotes(s.notes);
    }
    bytes = Buffer.from(
      await deck.write({ outputType: "nodebuffer", compression: true }),
    );
  } else {
    const pdf = await PDFDocument.create({ updateMetadata: false });
    pdf.setTitle(spec.title);
    const { regular, bold } = await documentFonts(
      pdf,
      spec.pages.flatMap((p) => [p.title ?? "", ...p.lines]),
    );
    for (const p of spec.pages) {
      const page = pdf.addPage([595.28, 841.89]);
      let y = 790;
      if (p.title) {
        requireThat(
          bold.widthOfTextAtSize(p.title, 20) <= 495,
          "LAYOUT_OVERFLOW",
          "PDF heading is too wide.",
        );
        page.drawText(p.title, {
          x: 50,
          y,
          size: 20,
          font: bold,
          color: rgb(0.09, 0.21, 0.36),
        });
        y -= 40;
      }
      for (const line of p.lines) {
        requireThat(
          regular.widthOfTextAtSize(line, 11) <= 495 && y >= 50,
          "LAYOUT_OVERFLOW",
          "PDF line or page overflows; wrap long lines or add another page.",
        );
        page.drawText(line, { x: 50, y, size: 11, font: regular });
        y -= 17;
      }
    }
    bytes = Buffer.from(await pdf.save({ useObjectStreams: false }));
  }
  requireThat(
    bytes.length <= task.maxFileBytes,
    "LIMIT_EXCEEDED",
    "Created document exceeds the file limit.",
  );
  if (!["pdf", "csv"].includes(spec.format)) unpack(bytes, task.maxExpandedBytes);
  writeFileSync(task.output, bytes, { flag: "wx", mode: 0o600 });
  return {
    revision: digest(bytes),
    format: spec.format,
    bytes: bytes.length,
    warnings:
      spec.format === "xlsx"
        ? [
            "Formulas request calculation on open; cached results are not calculated by the creation library.",
          ]
        : [],
  };
}
