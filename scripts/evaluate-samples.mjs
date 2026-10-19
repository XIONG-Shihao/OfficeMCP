// Real supplied files stay local. Reports contain aggregate evidence, not document text.
import fs from "node:fs";
import path from "node:path";
import { tmpdir } from "node:os";
import { OfficeService } from "../src/application/service.mjs";
import { configFrom, digest } from "../src/infrastructure/files.mjs";
import { VERSION } from "../src/contracts/registry.mjs";

const directory = path.resolve(
  process.argv[2] ?? new URL("../test_doc", import.meta.url).pathname,
);
const output = path.join(directory, "officemcp-results");
fs.mkdirSync(output, { recursive: true });
using temporary = fs.mkdtempDisposableSync(path.join(tmpdir(), "office-samples-"));

const savedConfig = path.join(directory, "officemcp-local.json");
const config = fs.existsSync(savedConfig)
  ? JSON.parse(fs.readFileSync(savedConfig, "utf8"))
  : {};
config.readRoots = [directory];
config.writeRoots = { output };
config.stateDirectory = path.join(temporary.path, "state");
config.toolset = "recommended";
config.limits = {
  maxFileBytes: 320 * 1024 ** 2,
  maxExpandedBytes: 384 * 1024 ** 2,
  maxStateBytes: 2 * 1024 ** 3,
  workerTimeoutMs: 120000,
  ...config.limits,
};
if (config.engines?.python)
  config.engines.python = path.resolve(directory, config.engines.python);
const configPath = path.join(temporary.path, "config.json");
fs.writeFileSync(configPath, JSON.stringify(config));

const service = new OfficeService(configFrom(configPath));
const report = {
  version: VERSION,
  localOnly: true,
  outputDirectory: output,
  files: [],
};
const stamp = Date.now();

async function call(item, name, args = {}) {
  const started = performance.now();
  const response = await service.call(name, args);
  item.steps.push({
    tool: name,
    elapsedMs: Math.round(performance.now() - started),
    status: response.payload.status,
    ...(response.isError ? { error: response.payload.error } : {}),
  });
  if (response.isError)
    throw Object.assign(new Error(response.payload.error.message), {
      code: response.payload.error.code,
    });
  return response.payload;
}

function editFor(format, read) {
  if (format === "docx") {
    const paragraph = read.data.items.find(
      (item) => item.kind === "paragraph" && item.text?.trim() && item.text.length < 2000,
    );
    if (!paragraph) throw new Error("No safe paragraph in the first DOCX window.");
    return {
      type: "replace_text",
      target: paragraph.target,
      expectedText: paragraph.text,
      replacement: paragraph.text + " [OfficeMCP verification]",
    };
  }
  if (format === "xlsx") {
    const cell = read.data.items.find(
      (item) => item.type === "string" && typeof item.value === "string" && !item.formula,
    );
    if (!cell) throw new Error("No literal cell in the sampled XLSX range.");
    return {
      type: "write_cells",
      sheet: read.data.sheet,
      cells: [{
        address: cell.address,
        expectedValue: cell.value,
        expectedFormula: null,
        value: cell.value + " [OfficeMCP verification]",
      }],
    };
  }
  if (format === "pptx") {
    const shape = read.data.items.find((item) => item.kind === "shape" && item.text?.trim() && item.shapeId);
    if (!shape) throw new Error("No editable shape on slide 1.");
    return {
      type: "replace_shape_text",
      slide: 1,
      shapeId: shape.shapeId,
      expectedText: shape.text,
      replacement: shape.text + " [OfficeMCP verification]",
    };
  }
  const pageCount = read.data.items.find((item) => item.kind === "document")?.pageCount;
  if (!pageCount) throw new Error("PDF overview did not return page count.");
  return {
    type: "pages",
    pages: Array.from({ length: pageCount }, (_, index) => ({
      page: pageCount - index,
      rotation: 0,
    })),
  };
}

try {
  for (const name of ["response.docx", "financial.xlsx", "AI_workshop_Shihao.pptx", "response.pdf"]) {
    const source = path.join(directory, name);
    if (!fs.existsSync(source)) continue;
    const sourceHash = digest(fs.readFileSync(source));
    const item = { file: name, bytes: fs.statSync(source).size, sourceHash, steps: [] };
    report.files.push(item);
    try {
      const opened = await call(item, "office_open", { path: source });
      const format = opened.document.format;
      let read = await call(item, "office_read", {
        documentId: opened.document.id,
        limit: 100,
        ...(format === "pptx" ? { slide: 1 } : {}),
        ...(format === "pdf" ? { scope: "overview" } : {}),
      });
      if (format === "xlsx") {
        const sheet = read.data.items.find((entry) => entry.kind === "sheet")?.name;
        if (!sheet) throw new Error("No readable worksheet.");
        read = await call(item, "office_read", {
          view: read.view,
          sheet,
          range: "A1:J20",
          limit: 200,
        });
        read.data.sheet = sheet;
      }
      if (format === "pdf") {
        const page = await call(item, "office_read", { view: read.view, page: 1, limit: 100 });
        const query = page.data.items.find((entry) => entry.text?.trim().length > 3)?.text.trim().slice(0, 30);
        if (query) await call(item, "pdf_search", {
          documentId: opened.document.id,
          expectedRevision: opened.document.sourceRevision,
          query,
          pageTo: Math.min(editFor(format, read).pages.length, 10),
          limit: 10,
        });
      }
      const edited = await call(item, "office_edit", {
        view: read.view,
        operationId: `sample-edit-${opened.document.id}`,
        action: editFor(format, read),
      });
      const saved = await call(item, "office_save", {
        view: edited.view,
        operationId: `sample-save-${opened.document.id}`,
        destination: { rootId: "output", relativePath: `sample-${stamp}-${name}` },
      });
      item.output = saved.data.path;
      item.validation = saved.validation;
      item.passed = true;
      await call(item, "office_close", {
        documentId: opened.document.id,
        disposition: "retain",
      });
    } catch (error) {
      item.passed = false;
      item.error = { code: error.code ?? "EVALUATION_ERROR", message: error.message };
    }
    item.originalUnchanged = digest(fs.readFileSync(source)) === sourceHash;
    console.log(JSON.stringify({
      file: name,
      passed: item.passed,
      originalUnchanged: item.originalUnchanged,
      error: item.error,
    }));
  }
} finally {
  service.close();
  fs.writeFileSync(
    path.join(output, "evaluation.json"),
    JSON.stringify(report, null, 2) + "\n",
  );
}

if (report.files.some((item) => !item.passed || !item.originalUnchanged))
  process.exitCode = 1;
