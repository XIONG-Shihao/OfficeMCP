import test from "./support/test.mjs";
import assert from "node:assert/strict";
import * as fs from "node:fs";
import path from "node:path";
import { tmpdir } from "node:os";
import { OfficeService } from "../src/application/service.mjs";
import { configFrom, digest } from "../src/infrastructure/files.mjs";
import { writeFourFormatFixtures } from "./four-format-fixtures.mjs";
import { editWorkflows } from "./edit-workflows.mjs";
function fixture() {
  const root = fs.mkdtempSync(path.join(tmpdir(), "office-edits-")),
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
    root,
    docs,
    s,
    close() {
      s.close();
      fs.rmSync(root, { recursive: true, force: true });
    },
    async call(name, args) {
      const r = await s.call(name, args);
      assert(!r.isError, JSON.stringify(r.payload));
      return r.payload;
    },
  };
}
test("M5 XLSX/PPTX edit, formula cache invalidation, exact preservation, validation and idempotent publication", async () => {
  const f = fixture();
  try {
    const hashes = ["budget.xlsx", "review.pptx"].map((n) =>
      digest(fs.readFileSync(path.join(f.docs, n))),
    );
    const start = performance.now();
    await editWorkflows(f.call, f.docs);
    assert(performance.now() - start < 15000);
    assert.deepEqual(
      ["budget.xlsx", "review.pptx"].map((n) =>
        digest(fs.readFileSync(path.join(f.docs, n))),
      ),
      hashes,
    );
  } finally {
    f.close();
  }
});
test("M5 rejects stale cells, formula injection, shared formulas and producer charts", async () => {
  const f = fixture();
  try {
    const d = (
      await f.call("office_open", { path: path.join(f.docs, "budget.xlsx") })
    ).document;
    const base = {
      documentId: d.id,
      expectedSourceRevision: d.sourceRevision,
      sheetId: "7",
      preservationProfile: "xlsx.cells.v1",
    };
    for (const [i, formula] of [
      "WEBSERVICE(A1)",
      "SUM([1]Sheet!A1)",
      "cmd|x!A1",
      "=SUM(B1,25)",
      "SUM(B1,25",
      "SUM(XFE1,2)",
    ].entries())
      assert(
        (
          await f.s.call("xlsx_set_formula", {
            ...base,
            operationId: "bad" + i,
            address: "B2",
            expectedRawValue: "150",
            expectedFormula: "SUM(B1,25)",
            formula,
          })
        ).isError,
      );
    assert(
      (
        await f.s.call("xlsx_set_formula", {
          ...base,
          operationId: "shared",
          address: "C2",
          expectedRawValue: "300",
          expectedFormula: "B2*2",
          formula: "B2*3",
        })
      ).isError,
    );
    assert.equal(
      (
        await f.s.call("xlsx_write_cells", {
          ...base,
          operationId: "stale",
          cells: [
            {
              address: "B1",
              expectedRawValue: "124",
              type: "number",
              value: "200",
            },
          ],
        })
      ).payload.error.code,
      "TARGET_CONFLICT",
    );
    fs.copyFileSync(
      new URL("./corpus/producer.xlsx", import.meta.url),
      path.join(f.docs, "producer.xlsx"),
    );
    const p = (
      await f.call("office_open", { path: path.join(f.docs, "producer.xlsx") })
    ).document;
    assert.equal(
      (
        await f.s.call("xlsx_write_cells", {
          ...base,
          documentId: p.id,
          expectedSourceRevision: p.sourceRevision,
          sheetId: "1",
          operationId: "chart",
          cells: [
            {
              address: "B2",
              expectedRawValue: "125",
              type: "number",
              value: "200",
            },
          ],
        })
      ).payload.error.code,
      "UNSUPPORTED_FEATURE",
    );
  } finally {
    f.close();
  }
});

test("M5 candidate cannot change if publication wins while an edit worker is running", async () => {
  const f = fixture();
  let release;
  try {
    const d = (
      await f.call("office_open", { path: path.join(f.docs, "review.pptx") })
    ).document;
    const a = {
      documentId: d.id,
      expectedSourceRevision: d.sourceRevision,
      operationId: "first",
      slide: 1,
      shapeId: "2",
      expectedText: "Quarterly review",
      expectedMatches: 1,
      replacement: "Quarterly results",
      preservationProfile: "pptx.text-replacement.v1",
    };
    const c = (await f.call("pptx_replace_text", a)).data;
    const binding = {
      documentId: d.id,
      candidateId: c.candidateId,
      expectedCandidateRevision: c.candidateRevision,
    };
    const v = (await f.call("office_validate", binding)).data;
    const worker = f.s.worker.bind(f.s);
    let ready;
    const started = new Promise((r) => (ready = r)),
      gate = new Promise((r) => (release = r));
    f.s.worker = async (...args) => {
      const result = await worker(...args);
      if (args[0] === "pptx_replace_text") {
        ready();
        await gate;
      }
      return result;
    };
    const competing = f.s.call("pptx_replace_text", {
      ...a,
      ...binding,
      operationId: "second",
      expectedText: "Quarterly results",
      replacement: "Must not win",
    });
    await started;
    const published = await f.call("office_commit", {
      ...binding,
      expectedSourceRevision: d.sourceRevision,
      validationId: v.validationId,
      operationId: "publish",
      mode: "save_as",
      destination: { rootId: "output", relativePath: "race.pptx" },
      overwrite: false,
    });
    release();
    assert.equal((await competing).payload.error.code, "ALREADY_COMMITTED");
    assert.equal(
      f.s.get("candidate", c.candidateId).revision,
      published.data.outputRevision,
    );
  } finally {
    release?.();
    f.close();
  }
});
