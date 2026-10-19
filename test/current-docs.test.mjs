import test from "./support/test.mjs";
import assert from "node:assert/strict";
import fs from "node:fs";
import { operationDescriptor } from "../src/contracts/registry.mjs";
import { parseOperationArguments } from "../src/contracts/operation-spec.mjs";
import { DOCX_ACTION_FAMILIES } from "../src/formats/docx/docx-agent-actions.mjs";
import { catalog } from "../src/contracts/catalog.mjs";

const read = (file) => fs.readFileSync(new URL(`../${file}`, import.meta.url), "utf8");

test("current usage keeps rendering outside MCP and documents every discovery family", () => {
  const view = { documentId: "doc", revision: "sha256:" + "a".repeat(64) };
  assert.equal(parseOperationArguments(operationDescriptor("docx-agent", "docx_query"), {
    view, select: { type: "page", layoutId: "layout", page: 12 },
  }).success, false);
  for (const file of ["README.md", "docs/DOCX_OPERATIONS.md", "docs/CAPABILITY_FIRST.md", "docs/USAGE.md", "docs/ENGINES.md"]) {
    const prose = read(file);
    assert(/scripted/i.test(prose), `${file} omits the separate scripted GUI-check boundary`);
    assert(!/docx_render\(\{|includeAnchors: true/.test(prose), `${file} recommends a removed rendering call`);
  }
  const actions = read("docs/DOCX_OPERATIONS.md");
  for (const family of DOCX_ACTION_FAMILIES) assert(actions.includes(`\`${family.id}\``), `Missing ${family.id} family`);
  assert(actions.includes("SDK round trip"));
  assert(actions.includes("Independent package checks"));
  assert(actions.includes("Application GUI checks"));
});

test("current engine guide agrees with the installed PDF dependency boundary", () => {
  const packageJson = JSON.parse(read("package.json"));
  assert(packageJson.dependencies.mupdf);
  assert.equal(packageJson.dependencies["pdfjs-dist"], undefined);
  const engines = read("docs/ENGINES.md");
  assert(engines.includes("MuPDF extraction"));
  assert(!engines.includes("PDF.js"), "The guide advertises the removed PDF.js runtime");
  const pdf = catalog().layers.find(({ id }) => id === "pdf");
  assert(pdf.engine.includes(`MuPDF.js ${packageJson.dependencies.mupdf}`));
  assert(pdf.engine.includes(`pdf-lib ${packageJson.dependencies["pdf-lib"]}`));
  assert(!/\bPDF\.js\b/.test(pdf.engine), "Public discovery advertises the removed PDF.js runtime");
});
