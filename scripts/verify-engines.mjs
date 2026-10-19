// Real documents stay local. Output contains aggregate evidence, never extracted text.
import * as fs from "node:fs";
import path from "node:path";
import { tmpdir } from "node:os";
import { OfficeService } from "../src/application/service.mjs";
import { configFrom, digest } from "../src/infrastructure/files.mjs";
const directory = path.resolve(
  process.argv[2] ?? new URL("../test_doc", import.meta.url).pathname,
);
const output = path.join(directory, "officemcp-results");
fs.mkdirSync(output, { recursive: true });
using rootDirectory = fs.mkdtempDisposableSync(
  path.join(tmpdir(), "officemcp-engines-"),
);
const root = rootDirectory.path,
  stamp = Date.now();
const engines = process.env.OFFICEMCP_ASPOSE_PYTHON
  ? {
      python: process.env.OFFICEMCP_ASPOSE_PYTHON,
      ...(process.env.OFFICEMCP_ASPOSE_LICENSE
        ? { asposeLicense: process.env.OFFICEMCP_ASPOSE_LICENSE }
        : { allowEvaluation: true }),
    }
  : {};
fs.writeFileSync(
  path.join(root, "config.json"),
  JSON.stringify({
    readRoots: [directory],
    writeRoots: { output },
    stateDirectory: path.join(root, "state"),
    engines,
    limits: {
      maxFileBytes: 320 * 1024 ** 2,
      maxExpandedBytes: 384 * 1024 ** 2,
      maxStateBytes: 2 * 1024 ** 3,
      workerTimeoutMs: 120000,
    },
  }),
);
const service = new OfficeService(configFrom(path.join(root, "config.json")));
const report = {
  date: new Date().toISOString(),
  localOnly: true,
  status: null,
  files: [],
};
const call = async (name, args) => {
  const r = await service.call(name, args);
  if (r.isError)
    throw Object.assign(new Error(r.payload.error.message), {
      code: r.payload.error.code,
    });
  return r.payload;
};
async function publish(d, c, file) {
  const v = (
    await call("office_validate", {
      documentId: d.id,
      candidateId: c.candidateId,
      expectedCandidateRevision: c.candidateRevision,
    })
  ).data;
  const p = (
    await call("office_commit", {
      documentId: d.id,
      candidateId: c.candidateId,
      expectedSourceRevision: d.sourceRevision,
      expectedCandidateRevision: c.candidateRevision,
      validationId: v.validationId,
      operationId: "publish-" + d.id,
      mode: "save_as",
      overwrite: false,
      destination: { rootId: "output", relativePath: file },
    })
  ).data;
  return { output: p.path, validation: v, engineEvidence: c.engineEvidence };
}
try {
  report.status = (await call("office_engine_status", {})).data;
  for (const file of [
    "response.docx",
    "financial.xlsx",
    "AI_workshop_Shihao.pptx",
  ]) {
    const source = path.join(directory, file);
    if (!fs.existsSync(source)) continue;
    const before = digest(fs.readFileSync(source)),
      start = performance.now(),
      item = { file, sourceRevision: before };
    report.files.push(item);
    try {
      const d = (await call("office_open", { path: source })).document;
      const view = { documentId: d.id, expectedRevision: d.sourceRevision };
      const base = {
        documentId: d.id,
        expectedSourceRevision: d.sourceRevision,
        operationId: "edit-" + d.id,
      };
      if (file.endsWith("docx")) {
        const read = (await call("docx_structure", { ...view, limit: 20 }))
          .data;
        const node = read.nodes.find((x) => x.text.trim().length > 3);
        const c = (
          await call("docx_edit_paragraph", {
            ...base,
            paragraphIndex: node.paragraphIndex,
            expectedText: node.text,
            replacement: node.text + " [OfficeMCP verified edit]",
            preservationProfile: "docx.superdoc.v1",
          })
        ).data;
        Object.assign(
          item,
          await publish(d, c, `response.engine-${stamp}.docx`),
        );
      } else if (file.endsWith("xlsx")) {
        const sheets = (await call("xlsx_inspect", { ...view, limit: 100 }))
          .data.items;
        const sheet = sheets.find(
          (x) => x.partName === "xl/worksheets/sheet10.xml",
        );
        if (!sheet) throw new Error("Sample sheet not found");
        const c = (
          await call("xlsx_edit_range", {
            ...base,
            sheet: sheet.name,
            cells: [
              {
                address: "D4",
                expectedValue: 254063.62,
                expectedFormula: null,
                value: 254064.62,
              },
            ],
            preservationProfile: "xlsx.aspose.v1",
          })
        ).data;
        Object.assign(
          item,
          await publish(d, c, `financial.engine-${stamp}.xlsx`),
        );
      } else {
        const read = (await call("pptx_engine_read", { ...view, slide: 1 }))
          .data;
        item.inspection = { shapeCount: read.shapes.length, slideCount: read.slideCount,
          engine: read.engine, evaluation: read.evaluation };
        const shape = read.shapes.find(
          (s) => typeof s.text === "string" && s.text.trim().length > 3,
        );
        const edited = (
          await call("pptx_edit_shape_text", {
            ...base,
            slide: 1,
            shapeId: shape.shapeId,
            expectedText: shape.text,
            replacement: shape.text + " [OfficeMCP]",
            preservationProfile: "pptx.aspose.v1",
          })
        ).data;
        const notes = (
          await call("pptx_engine_read", {
            documentId: d.id,
            candidateId: edited.candidateId,
            expectedRevision: edited.candidateRevision,
            slide: 1,
          })
        ).data.notes;
        const c = (
          await call("pptx_set_notes", {
            ...base,
            operationId: "notes-" + d.id,
            candidateId: edited.candidateId,
            expectedCandidateRevision: edited.candidateRevision,
            slide: 1,
            expectedText: notes,
            text: notes + "\nOfficeMCP verification note.",
            preservationProfile: "pptx.aspose.v1",
          })
        ).data;
        Object.assign(item, await publish(d, c, `slides.engine-${stamp}.pptx`));
      }
      item.passed = true;
    } catch (e) {
      item.passed = false;
      item.error = {
        code: e.code ?? "VERIFICATION_FAILED",
        message: e.message,
      };
      if (e.code === "ENGINE_LICENSE_REQUIRED" && item.inspection)
        item.licensePending = true;
    }
    item.originalUnchanged = before === digest(fs.readFileSync(source));
    item.elapsedMs = Math.round(performance.now() - start);
    console.log(
      JSON.stringify({
        file,
        passed: item.passed,
        error: item.error,
        elapsedMs: item.elapsedMs,
        originalUnchanged: item.originalUnchanged,
      }),
    );
  }
} finally {
  service.close();
  fs.writeFileSync(
    path.join(output, "engine-integration.json"),
    JSON.stringify(report, null, 2) + "\n",
  );
}
console.log("Evidence: " + path.join(output, "engine-integration.json"));
if (
  report.files.some(
    (x) => (!x.passed && !x.licensePending) || !x.originalUnchanged,
  )
)
  process.exitCode = 1;
