import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import path from "node:path";
import { unzipSync, strFromU8 } from "fflate";
export function verifyEditedFiles(docs) {
  for (const [source, output, changed] of [
    [
      "budget.xlsx",
      "budget-reviewed.xlsx",
      ["xl/worksheets/budget.xml", "xl/workbook.xml"],
    ],
    ["review.pptx", "review-reviewed.pptx", ["ppt/slides/slide9.xml"]],
  ]) {
    const before = unzipSync(readFileSync(path.join(docs, source))),
      after = unzipSync(readFileSync(path.join(docs, output)));
    assert.deepEqual(Object.keys(before).sort(), Object.keys(after).sort());
    for (const part of Object.keys(before))
      if (!changed.includes(part))
        assert.deepEqual(after[part], before[part], part);
    if (source.endsWith("xlsx")) {
      const xml = strFromU8(after["xl/worksheets/budget.xml"]);
      assert.match(xml, /<(?:\w+:)?c\b[^>]*r="B1"[^>]*>[\s\S]*?<v>200<\/v>/);
      assert.match(xml, /SUM\(B1,50\)/);
      assert(!xml.includes("<v>150</v>"));
      assert(!xml.includes("<v>300</v>"));
      assert(!xml.includes("<v>350</v>"));
    } else {
      const xml = strFromU8(after["ppt/slides/slide9.xml"]);
      assert(xml.includes("Quarterly results"));
      assert(xml.includes("Revenue 150"));
    }
  }
}
export async function editWorkflows(call, docs) {
  for (const format of ["xlsx", "pptx"]) {
    const filename = format === "xlsx" ? "budget.xlsx" : "review.pptx";
    const d = (await call("office_open", { path: path.join(docs, filename) }))
      .document;
    const base = { documentId: d.id, expectedSourceRevision: d.sourceRevision };
    let c = (
      await call(format === "xlsx" ? "xlsx_write_cells" : "pptx_replace_text", {
        ...base,
        operationId: format + "-edit",
        ...(format === "xlsx"
          ? {
              sheetId: "7",
              cells: [
                {
                  address: "B1",
                  expectedRawValue: "125",
                  type: "number",
                  value: "200",
                },
              ],
              preservationProfile: "xlsx.cells.v1",
            }
          : {
              slide: 1,
              shapeId: "2",
              expectedText: "Quarterly review",
              expectedMatches: 1,
              replacement: "Quarterly results",
              preservationProfile: "pptx.text-replacement.v1",
            }),
      })
    ).data;
    if (format === "xlsx")
      c = (
        await call("xlsx_set_formula", {
          ...base,
          candidateId: c.candidateId,
          expectedCandidateRevision: c.candidateRevision,
          operationId: "xlsx-formula",
          sheetId: "7",
          address: "B2",
          expectedRawValue: null,
          expectedFormula: "SUM(B1,25)",
          formula: "SUM(B1,50)",
          preservationProfile: "xlsx.cells.v1",
        })
      ).data;
    const binding = {
      documentId: d.id,
      candidateId: c.candidateId,
      expectedCandidateRevision: c.candidateRevision,
    };
    const v = (await call("office_validate", binding)).data;
    const commit = {
      ...binding,
      expectedSourceRevision: d.sourceRevision,
      validationId: v.validationId,
      operationId: format + "-commit",
      mode: "save_as",
      destination: {
        rootId: "output",
        relativePath:
          format === "xlsx" ? "budget-reviewed.xlsx" : "review-reviewed.pptx",
      },
      overwrite: false,
    };
    const receipt = await call("office_commit", commit);
    assert(receipt.data.published);
    assert.deepEqual((await call("office_commit", commit)).data, receipt.data);
  }
  verifyEditedFiles(docs);
}
