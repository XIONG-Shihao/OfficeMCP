import test from "./support/test.mjs";
import assert from "node:assert/strict";
import * as fs from "node:fs";
import path from "node:path";
import { tmpdir } from "node:os";
import { zipSync, unzipSync, strToU8, strFromU8 } from "fflate";
import { OfficeService } from "../src/application/service.mjs";
import { configFrom } from "../src/infrastructure/files.mjs";
import { xlsxBytes, pptxBytes } from "./four-format-fixtures.mjs";
test("M6 larger corpus: 10000 cells and 1000 shapes with bounded windows and worker deadlines", async (t) => {
  const root = fs.mkdtempSync(path.join(tmpdir(), "office-scale-")),
    docs = path.join(root, "docs");
  fs.mkdirSync(docs);
  const cfg = path.join(root, "config.json");
  fs.writeFileSync(
    cfg,
    JSON.stringify({
      readRoots: [docs],
      writeRoots: { output: docs },
      stateDirectory: path.join(root, "state"),
    }),
  );
  const x = unzipSync(xlsxBytes());
  let xml = strFromU8(x["xl/worksheets/budget.xml"]);
  xml = xml.replace(
    /<sheetData>[\s\S]*?<\/sheetData>/,
    "<sheetData>" +
      Array.from(
        { length: 1000 },
        (_, r) =>
          `<row r="${r + 1}">` +
          Array.from(
            { length: 10 },
            (_, c) =>
              `<c r="${String.fromCharCode(65 + c)}${r + 1}"><v>${r * 10 + c}</v></c>`,
          ).join("") +
          "</row>",
      ).join("") +
      "</sheetData>",
  );
  x["xl/worksheets/budget.xml"] = strToU8(xml);
  fs.writeFileSync(path.join(docs, "large.xlsx"), zipSync(x));
  const p = unzipSync(pptxBytes());
  xml = strFromU8(p["ppt/slides/slide9.xml"]);
  const shape = xml.match(/<p:sp>[\s\S]*?<\/p:sp>/)[0];
  xml = xml.replace(
    shape,
    Array.from({ length: 1000 }, (_, i) =>
      shape
        .replace('id="2"', 'id="' + (100 + i) + '"')
        .replace("Quarterly review", "Item " + i),
    ).join(""),
  );
  p["ppt/slides/slide9.xml"] = strToU8(xml);
  fs.writeFileSync(path.join(docs, "large.pptx"), zipSync(p));
  const s = new OfficeService(configFrom(cfg));
  try {
    for (const [file, name, args, total] of [
      [
        "large.xlsx",
        "xlsx_read",
        { sheetId: "7", range: "A1:J1000", limit: 100 },
        10000,
      ],
      ["large.pptx", "pptx_read", { slide: 1, limit: 100 }, 1004],
    ]) {
      const d = (await s.call("office_open", { path: path.join(docs, file) }))
        .payload.document;
      const start = performance.now();
      const r = await s.call(name, {
        documentId: d.id,
        expectedRevision: d.sourceRevision,
        ...args,
      });
      assert(!r.isError, JSON.stringify(r));
      assert.equal(r.payload.data.total, total);
      assert.equal(r.payload.data.items.length, 100);
      assert(r.payload.nextCursor);
      assert(Buffer.byteLength(JSON.stringify(r)) < 200000);
      const ms = Math.round(performance.now() - start);
      assert(ms < 15000);
      t.diagnostic(
        `${file}: ${total} source items, first 100 returned in ${ms} ms`,
      );
    }
  } finally {
    s.close();
    fs.rmSync(root, { recursive: true, force: true });
  }
});
