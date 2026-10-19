import test from "./support/test.mjs";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { tmpdir } from "node:os";
import { once } from "node:events";
import { formatAdapterId, runFormatAdapter } from "../src/workers/format-adapters.mjs";
import { runWorker } from "../src/workers/worker-runner.mjs";

test("worker dispatch uses exact operation names and explicit validation profiles", async () => {
  for (const [kind, id] of [
    ["office_read", "common.read"], ["xlsx_read", "opc.read"], ["xlsx_search", "opc.read"],
    ["pptx_read", "opc.read"], ["pdf_read", "pdf.pdfjs"],
    ["xlsx_edit_range", "aspose"],
    ["pptx_replace_frame", "presentation.structure"], ["pdf_fill_form", "pdf.forms"],
    ["xlsx_validate", "opc.edit"], ["pdf_validate", "pdf.package-edit"],
  ]) assert.equal(formatAdapterId({ kind }), id, kind);
  for (const [kind, profile, id] of [
    ["pdf_validate", "pdf.forms.v1", "pdf.forms"],
    ["pdf_validate", "pdf.mupdf.v1", "pdf.document"],
    ["pptx_validate", "pptx.structure.v1", "presentation.structure"],
    ["xlsx_validate", "xlsx.aspose.v1", "engine.validation"],
  ]) assert.equal(formatAdapterId({ kind, operations: [{ preservationProfile: profile, engineEvidence: {} }] }), id);
  for (const kind of ["pdf_typo", "xlsx_unknown", "pptx_unknown", "office_package_missing", "unknown_validate", "office_render", "office_visual_compare", "pdf_render", "pptx_render"]) {
    assert.equal(formatAdapterId({ kind, operations: [{ engineEvidence: {} }] }), null);
    assert.deepEqual(await runFormatAdapter({ kind, input: "/path/that/must/not/be/read" }), { handled: false });
  }
});

test("failed worker IPC serialization releases its process and scratch directory", async () => {
  const directory = fs.mkdtempSync(path.join(tmpdir(), "office-worker-cleanup-"));
  let closed;
  const service = {
    config: { stateDirectory: directory, workerTimeoutMs: 5000, maxExpandedBytes: 1024 },
    activeWorkers: new Set(),
    killWorker(child) { closed = once(child, "close"); child.kill("SIGKILL"); },
  };
  try {
    await assert.rejects(runWorker(service, "office_engine_status", null, { unserializable: 1n }), { code: "WORKER_FAILED" });
    await closed;
    assert.equal(service.activeWorkers.size, 0);
    assert.deepEqual(fs.readdirSync(directory), []);
  } finally { fs.rmSync(directory, { recursive: true, force: true }); }
});
