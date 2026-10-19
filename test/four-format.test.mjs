import test from "./support/test.mjs";
import assert from "node:assert/strict";
import * as fs from "node:fs";
import path from "node:path";
import { tmpdir } from "node:os";
import { unzipSync, zipSync, strFromU8, strToU8 } from "fflate";
import { OfficeService } from "../src/application/service.mjs";
import { configFrom, digest } from "../src/infrastructure/files.mjs";
import {
  writeFourFormatFixtures,
  xlsxBytes,
  pptxBytes,
  imageBytes,
} from "./four-format-fixtures.mjs";
function fixture() {
  const root = fs.mkdtempSync(path.join(tmpdir(), "officemcp-m4-")),
    docs = path.join(root, "docs");
  fs.mkdirSync(docs);
  writeFourFormatFixtures(docs);
  const cfg = path.join(root, "config.json");
  fs.writeFileSync(
    cfg,
    JSON.stringify({
      readRoots: [docs],
      writeRoots: { output: docs },
      stateDirectory: path.join(root, "state"),
    }),
  );
  const s = new OfficeService(configFrom(cfg));
  return {
    s,
    root,
    docs,
    close() {
      s.close();
      fs.rmSync(root, { recursive: true, force: true });
    },
  };
}
async function call(f, name, args) {
  const r = await f.s.call(name, args);
  assert(!r.isError, JSON.stringify(r.payload));
  return r.payload;
}
async function open(f, name) {
  const d = (await call(f, "office_open", { path: path.join(f.docs, name) }))
    .document;
  return { documentId: d.id, expectedRevision: d.sourceRevision };
}
async function all(f, name, args) {
  const result = [];
  let cursor;
  do {
    const r = await call(f, name, { ...args, ...(cursor ? { cursor } : {}) });
    result.push(...r.data.items);
    cursor = r.nextCursor;
  } while (cursor);
  return result;
}
function modify(bytes, part, change) {
  const p = unzipSync(bytes);
  p[part] = strToU8(change(strFromU8(p[part])));
  return Buffer.from(zipSync(p));
}
test("M4 XLSX formulas/caches/formats, hidden sheets, exact numbers and bounded blank-inclusive paging", async () => {
  const f = fixture();
  try {
    const a = await open(f, "budget.xlsx");
    const sheets = await all(f, "xlsx_inspect", { ...a, limit: 1 });
    assert.deepEqual(
      sheets.map((s) => [s.sheetId, s.state]),
      [
        ["7", "visible"],
        ["11", "veryHidden"],
      ],
    );
    const cells = await all(f, "xlsx_read", {
      ...a,
      sheetId: "7",
      range: "A1:E3",
      limit: 2,
    });
    assert.equal(cells.length, 15);
    assert.equal(cells.find((c) => c.address === "A1").value, "Unit cost");
    assert.equal(
      cells.find((c) => c.address === "E1").value,
      "9007199254740993",
    );
    const b2 = cells.find((c) => c.address === "B2");
    assert.equal(b2.value, null);
    assert.equal(b2.formula.expression, "SUM(B1,25)");
    assert.equal(b2.cachedValue.value, "150");
    assert.equal(b2.cachedValue.freshness, "unverified");
    assert.equal(b2.numberFormat.code, '"$"#,##0.00');
    assert.equal(b2.displayedText, null);
    assert.equal(
      cells.find((c) => c.address === "C3").formula.expression,
      null,
    );
    assert.equal(
      cells.find((c) => c.address === "C3").formula.attributes.si,
      "0",
    );
    assert(cells.find((c) => c.address === "C3").rowHidden);
    assert.equal(
      cells.find((c) => c.address === "D2").cachedValue.freshness,
      "missing",
    );
    assert.equal(cells.find((c) => c.address === "A3").storage, "absent");
    const hidden = await call(f, "xlsx_read", {
      ...a,
      sheetId: "11",
      range: "A1",
    });
    assert.equal(hidden.data.items[0].value, "Hidden reserve 77");
    assert.equal(hidden.data.dateSystem, "1904");
    const search = await call(f, "xlsx_search", {
      ...a,
      sheetId: "7",
      range: "A1:E3",
      query: "SUM",
    });
    assert.equal(search.data.items[0].address, "B2");
    const names = await call(f, "xlsx_inspect", {
      ...a,
      collection: "definedNames",
    });
    assert.equal(names.data.items[0].name, "InvoiceTotal");
    assert.equal(
      digest(fs.readFileSync(path.join(f.docs, "budget.xlsx"))),
      a.expectedRevision,
    );
  } finally {
    f.close();
  }
});
test("M4 PPTX respects slide order, hidden slides, groups, tables, notes and exact media retrieval", async () => {
  const f = fixture();
  try {
    const a = await open(f, "review.pptx");
    const slides = await all(f, "pptx_inspect", { ...a, limit: 1 });
    assert.equal(slides[0].partName, "ppt/slides/slide9.xml");
    assert.equal(slides[0].hidden, true);
    const shapes = await all(f, "pptx_read", { ...a, slide: 1, limit: 1 });
    assert.equal(shapes.find((s) => s.shapeId === "4").text, "Revenue 150");
    assert.deepEqual(shapes.find((s) => s.shapeId === "4").groupPath, ["3"]);
    assert.equal(
      shapes.find((s) => s.shapeId === "6").table[0][0].text,
      "Margin 42",
    );
    const notes = await call(f, "pptx_read", {
      ...a,
      slide: 1,
      domain: "notes",
    });
    assert(notes.data.items.some((s) => s.text.includes("reserve 77")));
    assert.equal(
      (await call(f, "pptx_read", { ...a, slide: 2, domain: "notes" })).data
        .notesState,
      "absent",
    );
    assert.equal(
      (await call(f, "pptx_search", { ...a, slide: 1, query: "Revenue" })).data
        .items[0].shapeId,
      "4",
    );
    const parts = await all(f, "office_package_parts", { ...a, limit: 2 });
    assert(parts.some((p) => p.partName === "ppt/media/logo.png"));
    const asset = (
      await call(f, "office_package_extract", {
        ...a,
        partName: "ppt/media/logo.png",
      })
    ).data;
    const bytes = (
      await call(f, "office_asset_read", { assetId: asset.assetId })
    ).data;
    assert.deepEqual(Buffer.from(bytes.bytes, "base64"), imageBytes);
    assert.equal(asset.hash, digest(imageBytes));
    assert.equal(
      digest(fs.readFileSync(path.join(f.docs, "review.pptx"))),
      a.expectedRevision,
    );
  } finally {
    f.close();
  }
});
test("M4 cursors bind sheet/range/domain and enforce range/format/missing-part limits", async () => {
  const f = fixture();
  try {
    const a = await open(f, "budget.xlsx");
    const r = await call(f, "xlsx_read", {
      ...a,
      sheetId: "7",
      range: "A1:B3",
      limit: 1,
    });
    for (const change of [{ sheetId: "11" }, { range: "A1:C3" }])
      assert.equal(
        (
          await f.s.call("xlsx_read", {
            ...a,
            sheetId: "7",
            range: "A1:B3",
            limit: 1,
            cursor: r.nextCursor,
            ...change,
          })
        ).payload.error.code,
        "CURSOR_CONFLICT",
      );
    assert.equal(
      (
        await f.s.call("xlsx_read", {
          ...a,
          sheetId: "7",
          range: "A1:XFD1048576",
        })
      ).payload.error.code,
      "LIMIT_EXCEEDED",
    );
    assert.equal(
      (await f.s.call("xlsx_read", { ...a, sheetId: "7", range: "XFE1" }))
        .payload.error.code,
      "INVALID_ARGUMENT",
    );
    assert.equal(
      (await f.s.call("pptx_inspect", a)).payload.error.code,
      "UNSUPPORTED_FORMAT",
    );
    assert.equal(
      (
        await f.s.call("office_package_extract", {
          ...a,
          partName: "../outside",
        })
      ).payload.error.code,
      "NOT_FOUND",
    );
    const p = await open(f, "review.pptx");
    const first = await call(f, "pptx_read", { ...p, slide: 1, limit: 1 });
    assert.equal(
      (
        await f.s.call("pptx_read", {
          ...p,
          slide: 1,
          domain: "notes",
          limit: 1,
          cursor: first.nextCursor,
        })
      ).payload.error.code,
      "CURSOR_CONFLICT",
    );
  } finally {
    f.close();
  }
});
test("M4 malformed relationships, shared strings, duplicate coordinates and strict namespaces fail honestly", async () => {
  const f = fixture();
  try {
    const variants = [
      [
        "bad-string.xlsx",
        modify(xlsxBytes(), "xl/worksheets/budget.xml", (s) =>
          s.replace("<v>0</v>", "<v>999</v>"),
        ),
        "xlsx_read",
        { sheetId: "7", range: "A1" },
        "INVALID_DOCUMENT",
      ],
      [
        "duplicate.xlsx",
        modify(xlsxBytes(), "xl/worksheets/budget.xml", (s) =>
          s.replace('r="B1"', 'r="A1"'),
        ),
        "xlsx_read",
        { sheetId: "7", range: "A1" },
        "INVALID_DOCUMENT",
      ],
      [
        "strict.xlsx",
        modify(xlsxBytes(), "xl/workbook.xml", (s) =>
          s.replace(
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            "http://purl.oclc.org/ooxml/spreadsheetml/main",
          ),
        ),
        "xlsx_inspect",
        {},
        "UNSUPPORTED_FEATURE",
      ],
      [
        "escape.pptx",
        modify(pptxBytes(), "ppt/_rels/presentation.xml.rels", (s) =>
          s.replace("slides/slide9.xml", "../../../outside.xml"),
        ),
        "pptx_inspect",
        {},
        "INVALID_DOCUMENT",
      ],
      [
        "missing.pptx",
        modify(pptxBytes(), "ppt/slides/_rels/slide9.xml.rels", (s) =>
          s.replace(
            "../notesSlides/notesSlide1.xml",
            "../notesSlides/missing.xml",
          ),
        ),
        "pptx_read",
        { slide: 1, domain: "notes" },
        "INVALID_DOCUMENT",
      ],
    ];
    for (const [name, bytes, tool, args, code] of variants) {
      fs.writeFileSync(path.join(f.docs, name), bytes);
      const a = await open(f, name);
      assert.equal(
        (await f.s.call(tool, { ...a, ...args })).payload.error.code,
        code,
        name,
      );
    }
  } finally {
    f.close();
  }
});
test("M4 independent producer corpus preserves table/chart inventory, missing formula cache and notes", async () => {
  const f = fixture();
  try {
    for (const name of ["producer.xlsx", "producer.pptx"])
      fs.copyFileSync(
        new URL("./corpus/" + name, import.meta.url),
        path.join(f.docs, name),
      );
    const x = await open(f, "producer.xlsx");
    const sheets = await all(f, "xlsx_inspect", x);
    assert.equal(sheets[1].state, "hidden");
    const cells = await all(f, "xlsx_read", {
      ...x,
      sheetId: sheets[0].sheetId,
      range: "A1:B4",
      limit: 3,
    });
    assert.equal(
      cells.find((c) => c.address === "B4").formula.expression,
      "SUM(B2:B3)",
    );
    assert.equal(
      cells.find((c) => c.address === "B4").cachedValue.freshness,
      "missing",
    );
    const parts = await all(f, "office_package_parts", { ...x, limit: 5 });
    assert(parts.some((p) => p.partName.startsWith("xl/charts/")));
    assert(parts.some((p) => p.partName.startsWith("xl/tables/")));
    const p = await open(f, "producer.pptx");
    const items = await all(f, "pptx_read", { ...p, slide: 1, limit: 2 });
    assert(items.some((i) => i.text.includes("Producer fixture")));
    assert(items.some((i) => i.table?.[1]?.[1]?.text === "150"));
    const notes = await all(f, "pptx_read", {
      ...p,
      slide: 1,
      domain: "notes",
    });
    assert(notes.some((i) => i.text.includes("Independent producer note 77")));
  } finally {
    f.close();
  }
});
test("M4 accepts alternate XML prefixes and reports foreign shape-tree objects without interpreting fallback text", async () => {
  const f = fixture();
  try {
    const xb = modify(xlsxBytes(), "xl/workbook.xml", (s) =>
      s
        .replace('xmlns="', 'xmlns:ss="')
        .replace(/<(\/?)([A-Za-z_][\w.-]*)(?=[\s/>])/g, "<$1ss:$2"),
    );
    fs.writeFileSync(path.join(f.docs, "prefix.xlsx"), xb);
    const x = await open(f, "prefix.xlsx");
    assert.equal(
      (await call(f, "xlsx_inspect", x)).data.items[0].name,
      "Budget",
    );
    const pb = modify(pptxBytes(), "ppt/slides/slide9.xml", (s) =>
      s.replace(
        "</p:spTree>",
        '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Fallback/></mc:AlternateContent></p:spTree>',
      ),
    );
    fs.writeFileSync(path.join(f.docs, "extension.pptx"), pb);
    const p = await open(f, "extension.pptx");
    const shapes = (await call(f, "pptx_read", { ...p, slide: 1 })).data.items;
    assert(
      shapes.some(
        (i) => i.interpretation === "unsupported-object" && i.text === null,
      ),
    );
    await call(f, "pptx_search", { ...p, slide: 1, query: "Revenue" });
  } finally {
    f.close();
  }
});
