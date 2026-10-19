import test from "./support/test.mjs";
import assert from "node:assert/strict";
import * as fs from "node:fs";
import path from "node:path";
import { tmpdir } from "node:os";
import { OfficeService } from "../src/application/service.mjs";
import { configFrom, digest } from "../src/infrastructure/files.mjs";
import { writeFixtures } from "./fixtures.mjs";
test("PDF user workflow searches and reads positioned text without changing source", async () => {
  const root = fs.mkdtempSync(path.join(tmpdir(), "office-pdf-usage-")),
    docs = path.join(root, "docs");
  fs.mkdirSync(docs);
  writeFixtures(docs);
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
  const call = async (n, a) => {
    const r = await s.call(n, a);
    assert(!r.isError, JSON.stringify(r));
    return r.payload;
  };
  try {
    const file = path.join(docs, "invoice.pdf"),
      hash = digest(fs.readFileSync(file)),
      d = (await call("office_open", { path: file })).document,
      a = { documentId: d.id, expectedRevision: d.sourceRevision };
    const found = await call("pdf_search", {
      ...a,
      query: "INVOICE TOTAL",
      pageTo: 1,
    });
    assert.equal(found.data.items.length, 1);
    assert.equal(found.data.items[0].page, 1);
    assert(found.data.items[0].positions.length);
    assert.equal(found.nextCursor, null);
    assert(
      (
        await s.call("pdf_search", {
          ...a,
          query: "invoice",
          pageFrom: 2,
          pageTo: 1,
        })
      ).isError,
    );
    const read = await call("pdf_read", { ...a, page: 1 });
    assert(read.data.items.some((item) => /Invoice total/i.test(item.text)));
    assert(read.data.items.every((item) => Array.isArray(item.transform)));
    assert((await s.call("pdf_read", { ...a, page: 9999 })).isError);
    const outline = await call("pdf_outline", a);
    assert.equal(outline.data.total, 0);
    assert.equal(digest(fs.readFileSync(file)), hash);
  } finally {
    s.close();
    fs.rmSync(root, { recursive: true, force: true });
  }
});
