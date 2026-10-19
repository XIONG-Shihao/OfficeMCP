import test from "./support/test.mjs";
import assert from "node:assert/strict";
import * as fs from "node:fs";
import path from "node:path";
import { tmpdir } from "node:os";
import { OfficeService } from "../src/application/service.mjs";
import { configFrom } from "../src/infrastructure/files.mjs";
import { serverInfo } from "../src/protocol/server-info.mjs";
import { catalog as executableCatalog } from "../src/contracts/catalog.mjs";
import { manifest, VERSION, CONTRACT } from "../src/contracts/registry.mjs";
import {
  LEGACY_MCP_PROTOCOL_VERSIONS,
  MCP_PROTOCOL_VERSION,
  TASKS_EXTENSION,
} from "../src/protocol/mcp-protocol.mjs";
import {
  DOCX_AGENT_ACTIONS,
} from "../src/formats/docx/docx-agent-actions.mjs";

test("Server discovery describes supported subsets without blanket denials", async () => {
  const root = fs.mkdtempSync(path.join(tmpdir(), "office-info-"));
  let service;
  try {
    const config = path.join(root, "config.json");
    const docs = path.join(root, "docs");
    fs.mkdirSync(docs);
    fs.writeFileSync(
      config,
      JSON.stringify({
        readRoots: [docs],
        writeRoots: { output: docs },
        stateDirectory: path.join(root, "state"),
      }),
    );
    service = new OfficeService(configFrom(config));
    const response = await service.call("office_server_info");
    assert(!response.isError, JSON.stringify(response.payload));
    const info = response.payload.data;
    assert.equal(info.version, VERSION);
    assert.equal(info.contractVersion, CONTRACT);
    assert.equal(info.protocol.current, MCP_PROTOCOL_VERSION);
    assert.deepEqual(
      info.protocol.legacyCompatibility,
      LEGACY_MCP_PROTOCOL_VERSIONS,
    );
    assert.equal(info.protocol.statelessRequests, true);
    assert.deepEqual(info.protocol.extensions, [TASKS_EXTENSION]);
    assert.deepEqual(
      info.operations,
      service.manifest().map((t) => t.name),
    );
    for (const feature of ["tracked-edits", "OCR"])
      assert(!info.unavailable.includes(feature), feature);
    assert.match(
      info.formats.docx.trackedReview,
      /Unsupported formatted or table-bearing ranges fail explicitly/,
    );
    assert.equal(info.formats.docx.rendering, undefined);
    assert(info.unavailable.includes("document-rendering"));
    assert.match(info.formats.pdf.ocr, /Optional local Tesseract.*office_read/);
    assert.match(info.formats.pdf.forms, /XFA.*signature.*refused/);
    assert.equal(info.formats.xlsx.rendering, undefined);
  } finally {
    service?.close();
    fs.rmSync(root, { recursive: true, force: true });
  }
});

test("DOCX-agent discovery retains PDF export and independent validation without rendering", () => {
  const root = fs.mkdtempSync(path.join(tmpdir(), "office-docx-agent-info-"));
  let service;
  try {
    const config = path.join(root, "config.json");
    const docs = path.join(root, "docs");
    fs.mkdirSync(docs);
    fs.writeFileSync(
      config,
      JSON.stringify({
        readRoots: [docs],
        writeRoots: { output: docs },
        stateDirectory: path.join(root, "state"),
        toolset: "docx-agent",
        engines: { libreOffice: process.execPath },
      }),
    );
    service = new OfficeService(configFrom(config));
    const info = serverInfo(service).data;
    assert.match(
      info.formats.docx.edit,
      new RegExp(`\\b${DOCX_AGENT_ACTIONS.length}\\b`),
    );
    assert(!/experimental|known-limited|fixture-tested/.test(info.formats.docx.edit));
    assert.equal(info.formats.docx.rendering, undefined);
    assert.match(info.formats.docx.pdfExport, /LibreOffice PDF export is configured/);
    assert.match(info.formats.docx.validation, /independent XML.*OPC.*binary-member/);
    assert(!info.unavailable.includes("DOCX-page-targets"));
    assert.equal(info.formats.docx.layout.tool, "docx_layout");
    assert.equal(info.formats.docx.layout.available, true);
  } finally {
    service?.close();
    fs.rmSync(root, { recursive: true, force: true });
  }
});

test("DOCX publication metadata declares retained engines without render tools", () => {
  const operations = new Map(manifest("docx-agent").map((tool) => [
    tool.name,
    tool._meta["io.officemcp/operation"],
  ]));
  assert.equal(operations.get("docx_export").risk, "publish");
  assert.deepEqual(operations.get("docx_export").supportedEngines, [
    "libreoffice",
    "mupdf",
  ]);
  assert.equal(operations.get("docx_render"), undefined);
});

test("Current release and usage counts agree with generated discovery", () => {
  const read = (name) =>
    fs.readFileSync(new URL(`../${name}`, import.meta.url), "utf8");
  const release = read("docs/RELEASE.md");
  assert(
    release.includes(
      `Version ${VERSION}; core contract ${CONTRACT} and agent DOCX contract docx/6;`,
    ),
  );
  const count = manifest("docx-agent").length;
  assert(
    read("docs/USAGE.md").includes(`${count}-tool agent-facing DOCX contract`),
  );
  assert(
    read("docs/CAPABILITIES.md").includes(
      `exposes ${count} agent-facing tools`,
    ),
  );
  const catalog = executableCatalog();
  const total = catalog.layers.reduce(
    (sum, layer) => sum + layer.operations.length,
    0,
  );
  assert(
    read("docs/CAPABILITIES.md").includes(
      `catalog contains ${total} entries`,
    ),
  );
  assert(!read("docs/USAGE.md").includes("No PDF OCR, editing, forms"));
});

test("Protocol dependencies, schemas and coordinator split stay current", () => {
  const read = (name) =>
    fs.readFileSync(new URL(`../${name}`, import.meta.url), "utf8");
  const packageJson = JSON.parse(read("package.json"));
  assert.equal(
    packageJson.dependencies["@modelcontextprotocol/server"],
    "2.0.0",
  );
  assert.equal(
    packageJson.devDependencies["@modelcontextprotocol/client"],
    "2.0.0",
  );
  assert.equal(
    packageJson.dependencies["@modelcontextprotocol/sdk"],
    undefined,
  );
  for (const name of [
    "contracts/tools.json",
    "contracts/docx-agent-tools.json",
  ])
    for (const tool of JSON.parse(read(name))) {
      assert.equal(
        tool.inputSchema.$schema,
        "https://json-schema.org/draft/2020-12/schema",
      );
      assert.equal(
        tool.outputSchema.$schema,
        "https://json-schema.org/draft/2020-12/schema",
      );
    }
  assert(read("src/application/service.mjs").split("\n").length < 900);
  assert(read("src/formats/docx/docx-host.mjs").split("\n").length < 1000);
  assert(read("src/contracts/results.mjs").split("\n").length < 700);
  for (const name of [
    "src/protocol/server-info.mjs",
    "src/application/tool-dispatch.mjs",
    "src/application/coordinator-state.mjs",
    "src/workers/worker-runner.mjs",
    "src/application/task-store.mjs",
    "src/contracts/catalog.mjs",
    "src/formats/docx/docx-canonical.mjs",
    "src/formats/docx/docx-agent-service.mjs",
    "src/formats/docx/docx-agent-inspect.mjs",
    "src/formats/docx/docx-agent-query.mjs",
    "src/formats/docx/docx-package-inventory.mjs",
    "src/formats/docx/docx-extended-actions.mjs",
    "src/formats/docx/docx-validation.mjs",
  ])
    assert(fs.existsSync(new URL(`../${name}`, import.meta.url)));
  for (const tool of manifest("docx-agent")) {
    const metadata = tool._meta?.["io.officemcp/operation"];
    assert(metadata, `${tool.name} lacks neutral operation metadata`);
    assert.equal(metadata.name, tool.name);
    assert(["read", "mutate", "publish"].includes(metadata.risk));
    assert(["synchronous", "taskEligible"].includes(metadata.execution));
    assert(metadata.supportedEngines.length > 0);
    assert(metadata.supportedEngines.every((engine) => typeof engine === "string"));
    assert.equal(metadata.resultContract, "docx/6");
  }
});

test("Modern task records and legacy job records remain separate", async () => {
  const root = fs.mkdtempSync(path.join(tmpdir(), "office-task-kinds-"));
  let service;
  try {
    const config = path.join(root, "config.json");
    const docs = path.join(root, "docs");
    fs.mkdirSync(docs);
    fs.writeFileSync(
      config,
      JSON.stringify({
        readRoots: [docs],
        writeRoots: { output: docs },
        stateDirectory: path.join(root, "state"),
      }),
    );
    service = new OfficeService(configFrom(config));
    const taskId = service.startTask("office_server_info", {});
    const jobId = service.startBackground("office_server_info", {});
    assert.match(taskId, /^task_/);
    assert.match(jobId, /^job_/);
    assert.equal(service.rows("task").some((record) => record.id === taskId), true);
    assert.equal(service.rows("job").some((record) => record.id === jobId), true);
    assert.throws(() => service.get("job", taskId));
    assert.throws(() => service.taskStore.get(jobId));
    for (let attempt = 0; attempt < 100; attempt++) {
      if (
        service.get("task", taskId).status !== "running" &&
        service.get("job", jobId).status !== "running"
      )
        break;
      await new Promise((resolve) => setTimeout(resolve, 10));
    }
    assert.equal(service.get("task", taskId).status, "completed");
    assert.equal(service.get("job", jobId).status, "completed");
  } finally {
    service?.close();
    fs.rmSync(root, { recursive: true, force: true });
  }
});

test("Documentation lives under docs and local links resolve", () => {
  const root = new URL("../", import.meta.url);
  assert.deepEqual(
    fs.readdirSync(root).filter((name) => /\.md$/i.test(name)),
    ["README.md"],
  );
  const pages = [
    "README.md",
    ...fs
      .readdirSync(new URL("docs/", root))
      .filter((name) => name.endsWith(".md"))
      .map((name) => `docs/${name}`),
  ];
  for (const page of pages) {
    const url = new URL(page, root);
    const markdown = fs.readFileSync(url, "utf8");
    for (const match of markdown.matchAll(/\[[^\]]*\]\(([^)]+)\)/g)) {
      const target = match[1];
      if (/^[a-z][a-z\d+.-]*:/i.test(target)) continue;
      const resolved = new URL(target, url);
      assert(fs.existsSync(resolved), `${page}: missing ${target}`);
      if (resolved.hash && resolved.pathname.endsWith(".md")) {
        const headings = fs
          .readFileSync(resolved, "utf8")
          .split("\n")
          .filter((line) => /^#{1,6} /.test(line))
          .map((line) =>
            line
              .replace(/^#+ /, "")
              .toLowerCase()
              .replace(/[^\p{L}\p{N}_\s-]/gu, "")
              .replace(/\s/g, "-"),
          );
        assert(
          headings.includes(decodeURIComponent(resolved.hash.slice(1))),
          `${page}: missing heading ${target}`,
        );
      }
    }
  }
});
