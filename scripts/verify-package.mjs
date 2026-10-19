import { execFileSync } from "node:child_process";
import * as fs from "node:fs";
import path from "node:path";
import { tmpdir } from "node:os";
import {fileURLToPath} from "node:url";
const npmCommand=process.platform==='win32'?process.execPath:'npm';
const npmPrefix=process.platform==='win32'?[process.env.npm_execpath??path.join(path.dirname(process.execPath),'node_modules/npm/bin/npm-cli.js')]:[];
import assert from "node:assert/strict";
import { unzipSync, strFromU8 } from "fflate";
import { DOMParser } from '@xmldom/xmldom';
import { Client } from "@modelcontextprotocol/client";
import { StdioClientTransport } from "@modelcontextprotocol/client/stdio";
import { writeFixtures } from "../test/fixtures.mjs";
import {
  writeFourFormatFixtures,
  imageBytes,
} from "../test/four-format-fixtures.mjs";
import { editWorkflows } from "../test/edit-workflows.mjs";
const project = fileURLToPath(new URL("../", import.meta.url));
using packageDirectory = fs.mkdtempDisposableSync(
  path.join(tmpdir(), "officemcp-package-"),
);
const root = packageDirectory.path,
  install = path.join(root, "install"),
  docs = path.join(root, "docs");
fs.mkdirSync(install);
fs.mkdirSync(docs);
writeFixtures(docs);
writeFourFormatFixtures(docs);
const cache = process.env.npm_config_cache ?? path.join(tmpdir(), "officemcp-npm-cache");
const packed = JSON.parse(
  execFileSync(
    npmCommand,
    [...npmPrefix, "pack", "--json", "--pack-destination", root, "--cache", cache],
    { cwd: project, encoding: "utf8", maxBuffer: 32 * 1024 * 1024 },
  ),
)[0];
const markdownPaths = packed.files
  .map((file) => file.path)
  .filter((name) => name.endsWith(".md") && !name.startsWith("node_modules/"))
  .sort();
for (const removed of [
  "src/formats/common/local-render.mjs",
  "src/formats/pdf/pdf-raster.mjs",
  "src/formats/docx/docx-page-map.mjs",
  "src/formats/docx/docx-render-anchors.mjs",
  "contracts/docx-tools.json",
  "src/formats/docx/docx-contract.mjs",
  "src/formats/docx/docx-results.mjs",
])
  assert(
    !packed.files.some((file) => file.path === removed),
    `Removed artifact was packaged: ${removed}`,
  );
assert.deepEqual(
  markdownPaths,
  [
    "README.md",
    "agent/OfficeMCP_docx.md",
    ...fs
      .readdirSync(path.join(project, "docs"))
      .filter((name) => name.endsWith(".md"))
      .map((name) => `docs/${name}`),
  ].sort(),
  "Package must contain the portable agent guide, current docs and only README at its root",
);
assert(
  packed.files.some((file) => file.path === "verification.json"),
  "Documentation's linked machine evidence must ship",
);
assert(
  packed.files.some(
    (file) => file.path === "contracts/docx-agent-tools.json",
  ),
  "The installed package must ship the agent-facing DOCX contract",
);
const second = JSON.parse(
  execFileSync(
    npmCommand,
    [...npmPrefix, "pack", "--json", "--pack-destination", root, "--cache", cache],
    { cwd: project, encoding: "utf8", maxBuffer: 32 * 1024 * 1024 },
  ),
)[0];
assert.equal(
  second.integrity,
  packed.integrity,
  "Repeated packs must produce identical artifact bytes",
);
assert(
  packed.files.every(
    (f) =>
      !["src/http.mjs", "REMOTE.md", "src/remote-service.mjs"].includes(f.path),
  ),
  "Remote implementation must not ship",
);
assert(
  packed.files.some((f) => f.path === "npm-shrinkwrap.json"),
  "Packed runtime must pin transitive dependencies",
);
assert(
  packed.files.every(
    (f) =>
      !f.path.startsWith("runtime/") &&
      !f.path.startsWith("src/main/") &&
      !f.path.startsWith("test_doc/") &&
      !f.path.startsWith(".engines/") &&
      !f.path.includes("__pycache__") &&
      !f.path.endsWith(".pyc"),
  ),
);
// Install the standalone CLI as the project root of its shipped lockfile.
// Consumer-root overrides/nested-shrinkwrap behavior must not silently select a
// different dependency tree (including older patched transitive packages).
const installedPackage = path.join(install, "node_modules/officemcp-local-preview");
fs.mkdirSync(installedPackage, { recursive: true });
execFileSync("tar", ["-xzf", path.join(root, packed.filename), "-C", installedPackage, "--strip-components=1"]);
execFileSync(
  npmCommand,
  [
    ...npmPrefix, "ci",
    "--omit=dev",
    "--offline",
    "--ignore-scripts",
    "--no-audit",
    "--no-fund",
    "--cache",
    cache,
  ],
  { cwd: installedPackage, stdio: "pipe" },
);
const installedLock = JSON.parse(fs.readFileSync(path.join(installedPackage, "npm-shrinkwrap.json")));
let verifiedRuntimePackages = 0;
for (const [name, entry] of Object.entries(installedLock.packages)) {
  if (!name || entry.dev) continue;
  const file = path.join(installedPackage, name, "package.json");
  if (!fs.existsSync(file)) { assert(entry.optional, `Missing locked runtime dependency: ${name}`); continue; }
  assert.equal(JSON.parse(fs.readFileSync(file)).version, entry.version, `Installed dependency differs from lock: ${name}`);
  verifiedRuntimePackages++;
}
const entry = path.join(
    install,
    "node_modules/officemcp-local-preview/src/cli.mjs",
  ),
  config = path.join(root, "config.json");
fs.writeFileSync(
  config,
  JSON.stringify({
    readRoots: [docs],
    writeRoots: { output: docs },
    stateDirectory: path.join(root, "state"),
  }),
);
const client = new Client(
  { name: "standalone-package-test", version: "1" },
  {
    capabilities: {},
    versionNegotiation: { mode: { pin: "2026-07-28" } },
  },
);
const transport = new StdioClientTransport({
  command: process.execPath,
  args: [entry, "serve", "--config", config],
  cwd: install,
  stderr: "pipe",
  env: { PATH: process.env.PATH ?? "", NODE_PATH: "" },
});
transport.stderr?.on("data", () => {});
async function call(name, args = {}) {
  const r = await client.callTool({ name, arguments: args });
  assert(!r.isError, JSON.stringify(r.structuredContent));
  assert.deepEqual(r.structuredContent, JSON.parse(r.content[0].text));
  return r.structuredContent;
}
try {
  await client.connect(transport);
  const info = await call("office_server_info");
  const d = (
    await call("office_open", { path: path.join(docs, "contract.docx") })
  ).document;
  const engineStatus = await call("office_engine_status");
  assert.equal(engineStatus.data.superdoc.version, "2.2.1");
  const structure = await call("docx_structure", {
    documentId: d.id,
    expectedRevision: d.sourceRevision,
    limit: 2,
  });
  assert.equal(structure.data.nodes[0].text, "Payment within 30 days");
  const engineCandidate = (
    await call("docx_edit_paragraph", {
      documentId: d.id,
      expectedSourceRevision: d.sourceRevision,
      operationId: "package-superdoc",
      paragraphIndex: 0,
      expectedText: "Payment within 30 days",
      replacement: "Payment within 60 days",
      preservationProfile: "docx.superdoc.v1",
    })
  ).data;
  const engineValidation = await call("office_validate", {
    documentId: d.id,
    candidateId: engineCandidate.candidateId,
    expectedCandidateRevision: engineCandidate.candidateRevision,
  });
  assert.equal(engineValidation.data.passed, true);
  const reviewBase = await call("office_create", {
    operationId: "package-review-create",
    document: {
      format: "docx",
      blocks: [{ kind: "paragraph", text: "Review this payment term." }],
    },
  });
  const detail = await call("docx_read_details", {
    view: reviewBase.view,
    domain: "paragraphs",
  });
  assert.equal(detail.data.items[0].text, "Review this payment term.");
  assert(
    !(await client.listTools()).tools.some((t) =>
      t.name.startsWith("docx_comment_"),
    ),
  );
  const c = (
    await call("docx_replace_text", {
      documentId: d.id,
      expectedSourceRevision: d.sourceRevision,
      operationId: "package-edit",
      paragraphId: "p_0",
      expectedText: "30 days",
      expectedMatches: 1,
      replacement: "45 days",
      changeMode: "direct",
      preservationProfile: "docx.text-replacement.v1",
    })
  ).data;
  const v = (
    await call("office_validate", {
      documentId: d.id,
      candidateId: c.candidateId,
      expectedCandidateRevision: c.candidateRevision,
    })
  ).data;
  const committed = await call("office_commit", {
    documentId: d.id,
    candidateId: c.candidateId,
    expectedSourceRevision: d.sourceRevision,
    expectedCandidateRevision: c.candidateRevision,
    validationId: v.validationId,
    operationId: "package-commit",
    mode: "save_as",
    destination: { rootId: "output", relativePath: "reviewed.docx" },
    overwrite: false,
  });
  assert(committed.data.published);
  const pdf = (
    await call("office_open", { path: path.join(docs, "invoice.pdf") })
  ).document;
  assert(
    (
      await call("pdf_read", {
        documentId: pdf.id,
        expectedRevision: pdf.sourceRevision,
        page: 1,
      })
    ).data.items.some((i) => i.text.includes("125 dollars")),
  );
  const pdfView = { documentId: pdf.id, expectedRevision: pdf.sourceRevision };
  assert(
    (await call("pdf_search", { ...pdfView, query: "invoice", pageTo: 1 })).data
      .items.length > 0,
  );
  assert(Array.isArray((await call("pdf_outline", pdfView)).data.items));
  const x = (
    await call("office_open", { path: path.join(docs, "budget.xlsx") })
  ).document;
  const xa = { documentId: x.id, expectedRevision: x.sourceRevision };
  assert.equal(
    (await call("xlsx_inspect", xa)).data.items[1].state,
    "veryHidden",
  );
  const cell = (await call("xlsx_read", { ...xa, sheetId: "7", range: "B2" }))
    .data.items[0];
  assert.equal(cell.formula.expression, "SUM(B1,25)");
  assert.equal(cell.cachedValue.freshness, "unverified");
  const xview = { documentId: x.id, revision: x.sourceRevision };
  const xread = await call("xlsx_query", { view: xview, sheet: "Budget", range: "B1:B2" });
  assert.equal(xread.data.storage, 'rust-indexed-v1', 'Installed package must use its bundled native helper');
  assert.equal(xread.data.items[1].calculation.value, 150);
  const xedited = await call("xlsx_apply", { view: xview, operationId: "package-oss-xlsx",
    commands: [{ target: xread.data.items[0].target, expectedValue: 125, value: 150 }] });
  const xfresh = await call("xlsx_query", { view: xedited.view, sheet: "Budget", range: "B1:B2" });
  assert.equal(xfresh.data.items[1].calculation.value, 175);
  assert.equal(xfresh.data.items[1].formula, "SUM(B1,25)");
  assert.equal(xfresh.data.items[1].calculation.revision, xedited.view.revision);
  const xsaved = await call("office_save", { view: xedited.view, operationId: "package-oss-save",
    destination: { rootId: "output", relativePath: "oss-budget.xlsx" } });
  assert(xsaved.data.published);
  const originalParts = unzipSync(fs.readFileSync(path.join(docs, "budget.xlsx")));
  const savedParts = unzipSync(fs.readFileSync(path.join(docs, "oss-budget.xlsx")));
  for (const [name, value] of Object.entries(originalParts))
    if (!["xl/workbook.xml", "xl/worksheets/budget.xml"].includes(name)) assert.deepEqual(savedParts[name], value, name);
  const savedXml = strFromU8(savedParts["xl/worksheets/budget.xml"]);
  const savedDoc = new DOMParser().parseFromString(savedXml, 'application/xml');
  const savedB2 = Array.from(savedDoc.getElementsByTagNameNS('http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'c')).find(cell => cell.getAttribute('r') === 'B2');
  assert.equal(savedB2.getElementsByTagNameNS('http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'f')[0].textContent, 'SUM(B1,25)');
  // A freshly calculated numeric cache may serialize as 175 or 175.0.
  // Compare exact decimal spelling without converting arbitrary numeric text
  // through Number (which would hide precision loss in large literals).
  assert.ok(!savedB2.hasAttribute('t') || savedB2.getAttribute('t') === 'n');
  assert.match(savedB2.getElementsByTagNameNS('http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'v')[0].textContent, /^175(?:\.0+)?$/);
  assert(savedXml.includes("9007199254740993"));
  const reopenedXlsx = await call("office_open", { path: path.join(docs, "oss-budget.xlsx") });
  const reopenedXlsxRange = await call("xlsx_query", { view: { documentId: reopenedXlsx.document.id,
    revision: reopenedXlsx.document.sourceRevision }, sheet: "Budget", range: "B2" });
  assert.equal(reopenedXlsxRange.data.items[0].calculation.value, 175);
  const csvSource = '\ufeffid;price;note\r\n00123;125;"中文; literal =1+1"\r\n';
  const csvPath = path.join(docs, 'literal-prices.csv'); fs.writeFileSync(csvPath, csvSource);
  const csvOpened = await call('office_open', { path: csvPath });
  assert.equal(csvOpened.document.format, 'csv');
  const csvView = { documentId: csvOpened.document.id, revision: csvOpened.document.sourceRevision };
  const csvRead = await call('xlsx_query', { view: csvView, sheet: 'CSV', range: 'A2:C2' });
  assert.equal(csvRead.data.items[0].value, '00123');
  const csvChanged = await call('xlsx_apply', { view: csvView, operationId: 'package-csv-edit',
    commands: [{ target: csvRead.data.items[1].target, expectedValue: '125', value: '150' }] });
  const csvSaved = await call('office_save', { view: csvChanged.view, operationId: 'package-csv-save',
    destination: { rootId: 'output', relativePath: 'literal-prices-delivered.csv' } });
  assert.equal(fs.readFileSync(csvSaved.data.path, 'utf8'), csvSource.replace('125', '150'));
  assert.equal(fs.readFileSync(csvPath, 'utf8'), csvSource);
  const csvReopened = await call('office_open', { path: csvSaved.data.path });
  const csvReread = await call('xlsx_query', { view: { documentId: csvReopened.document.id, revision: csvReopened.document.sourceRevision }, sheet: 'CSV', range: 'A2:C2' });
  assert.equal(csvReread.data.items[1].value, '150');
  assert.equal(csvReread.data.items[2].value, '中文; literal =1+1');
  const p = (
    await call("office_open", { path: path.join(docs, "review.pptx") })
  ).document;
  const pa = { documentId: p.id, expectedRevision: p.sourceRevision };
  assert((await call("pptx_inspect", pa)).data.items[0].hidden);
  assert(
    (
      await call("pptx_read", { ...pa, slide: 1, domain: "notes" })
    ).data.items.some((i) => i.text.includes("reserve 77")),
  );
  const part = (
    await call("office_package_extract", {
      ...pa,
      partName: "ppt/media/logo.png",
    })
  ).data;
  const extracted = (await call("office_asset_read", { assetId: part.assetId }))
    .data;
  assert.deepEqual(Buffer.from(extracted.bytes, "base64"), imageBytes);
  await editWorkflows(call, docs);
  assert.equal((await call("office_state_info")).data.schemaVersion, 3);
  assert.equal((await call("office_cleanup")).data.publishedOutputsDeleted, 0);
  assert(
    fs.existsSync(
      path.join(
        install,
        "node_modules/officemcp-local-preview/contracts/dependencies.json",
      ),
    ),
  );
  for (const file of packed.files.filter(
    (f) => f.path.startsWith("src/") && !f.path.endsWith("/"),
  ))
    assert.deepEqual(
      fs.readFileSync(path.join(project, file.path)),
      fs.readFileSync(
        path.join(install, "node_modules/officemcp-local-preview", file.path),
      ),
      "Installed runtime differs from source: " + file.path,
    );
  const unicode = await call("office_create", {
    operationId: "package-unicode",
    document: {
      format: "pdf",
      pages: [{ title: "季度报告", lines: ["付款期限"] }],
    },
  });
  const unicodeRead = await call("office_read", {
    view: unicode.view,
    page: 1,
  });
  assert(JSON.stringify(unicodeRead.data.items).includes("季度报告"));
  fs.copyFileSync(
    new URL("./verify-extensions.mjs", import.meta.url),
    path.join(install, "verify-extensions.mjs"),
  );
  execFileSync(
    process.execPath,
    [path.join(install, "verify-extensions.mjs"), root, docs],
    { cwd: install, stdio: "pipe" },
  );
  const catalog = JSON.parse(
    fs.readFileSync(
      path.join(
        install,
        "node_modules/officemcp-local-preview/contracts/catalog.json",
      ),
    ),
  );
  assert.equal(catalog.version, info.data.version);
  assert.equal(catalog.layers.length, 5);
  const modernTools = (await client.listTools()).tools;
  assert(!modernTools.some((tool) => ["office_render", "docx_render", "pdf_render", "pptx_render"].includes(tool.name)));
  const legacyOnlyTools = ["office_job_status", "office_job_cancel"];
  assert(
    legacyOnlyTools.every(
      (name) => !modernTools.some((tool) => tool.name === name),
    ),
  );
  assert.equal(
    catalog.layers.flatMap((l) => l.operations).length,
    modernTools.filter((tool) => !tool.name.startsWith("docx_")).length +
      legacyOnlyTools.length +
      14,
  );
  const evidence = {
    localOnly: true,
    pdfSearchOutline: true,
    operationCatalog: true,
    embeddedSdk: true,
    sqliteStorageAdapter: true,
    xlsxPptxEdits: true,
    version: info.data.version,
    platform: info.data.platform,
    root,
    entry,
    artifact: path.join(root, packed.filename),
    sha512: packed.integrity,
    offlineInstall: true,
    installMethod: "Extract standalone tarball and npm ci --omit=dev --offline using its shipped root shrinkwrap",
    consumerDependencyInstallQualified: false,
    verifiedRuntimePackages,
    repeatedPackIdentical: true,
    dependencyInventory: true,
    docxPublication: true,
    pdfText: true,
    renderingToolsAbsent: true,
    xlsxReads: true,
    xlsxOpenSourceEditCalculateSaveReopen: true,
    pptxReads: true,
    exactPartExtraction: true,
    modernToolCount: modernTools.length,
    legacyToolCount: modernTools.length + legacyOnlyTools.length,
    docxAgentToolCount: 14,
  };
  fs.writeFileSync(
    path.join(root, "evidence.json"),
    JSON.stringify(evidence, null, 2),
  );
  console.log("PASS " + JSON.stringify(evidence));
} finally {
  await client.close();
}
// Verify the compact agent contract from the independently installed tarball.
const agentConfig = path.join(root, "docx-agent-config.json");
fs.writeFileSync(
  agentConfig,
  JSON.stringify({
    readRoots: [docs],
    writeRoots: { output: docs },
    stateDirectory: path.join(root, "docx-agent-state"),
    toolset: "docx-agent",
  }),
);
const agentClient = new Client(
    { name: "standalone-docx-agent-test", version: "1" },
    {
      capabilities: {},
      versionNegotiation: { mode: { pin: "2026-07-28" } },
    },
  ),
  agentTransport = new StdioClientTransport({
    command: process.execPath,
    args: [entry, "serve", "--config", agentConfig],
    cwd: install,
    stderr: "pipe",
    env: { PATH: process.env.PATH ?? "", NODE_PATH: "" },
  });
agentTransport.stderr?.on("data", () => {});
try {
  await agentClient.connect(agentTransport);
  const tools = (await agentClient.listTools()).tools;
  assert.equal(tools.length, 14);
  assert(!tools.some((tool) => ["docx_render", "office_render", "pdf_render", "pptx_render"].includes(tool.name)));
  assert(tools.every((tool) => tool.name.startsWith("docx_")));
  const invoke = async (name, args) => {
    const response = await agentClient.callTool({ name, arguments: args });
    assert(!response.isError, JSON.stringify(response.structuredContent));
    assert.equal(response.structuredContent.contractVersion, "docx/6");
    return response.structuredContent;
  };
  let view = (
    await invoke("docx_create", {
      operationId: "agent-create",
      document: {
        format: "docx",
        blocks: [{ kind: "paragraph", text: "Payment within 30 days" }],
      },
    })
  ).view;
  const inspection = await invoke("docx_inspect", {
    view,
    actions: ["rewrite_block", "crop_image", "insert_caption"],
  });
  assert(inspection.data.availableActions.includes("rewrite_block"));
  assert(inspection.data.availableActions.includes("insert_caption"));
  assert(inspection.data.availableActions.includes("crop_image"));
  assert(!inspection.data.availableActions.includes("undo_changes"));
  assert.equal(inspection.data.qualifiedActions, undefined);
  assert(Array.isArray(inspection.data.storyInventory));
  assert(Array.isArray(inspection.data.objectInventory));
  const firstRead = await invoke("docx_read", {
    view,
    representation: "complete",
    offset: 0,
    limit: 1,
  });
  assert(firstRead.data.readSnapshotId);
  assert.equal(
    firstRead.data.remainingReads[0]?.readSnapshotId,
    firstRead.data.readSnapshotId,
  );
  const target = (
    await invoke("docx_query", {
      view,
      select: { type: "text", pattern: "30 days" },
      require: "exactlyOne",
    })
  ).data.items[0].target;
  view = (
    await invoke("docx_apply", {
      view,
      operationId: "agent-edit",
      action: "rewrite_block",
      selectorTarget: target,
      arguments: { text: "Payment within 45 days" },
    })
  ).view;
  const validation = await invoke("docx_validate", { view });
  assert.equal(validation.data.passed, true);
  assert(validation.data.checks.includes("output-XML-well-formed"));
  await invoke("docx_save", {
    view,
    operationId: "agent-save",
    destination: {
      rootId: "output",
      relativePath: "installed-docx-agent.docx",
    },
  });
  console.log(
    "PASS installed DOCX agent contract: 14 tools, opaque target edit and independent package validation",
  );
} finally {
  await agentClient.close();
}
