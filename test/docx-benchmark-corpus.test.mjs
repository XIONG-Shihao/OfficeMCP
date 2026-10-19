import test from "./support/test.mjs";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { tmpdir } from "node:os";
import { benchmarkOptions, benchmarkReportPath, writeBenchmarkReport } from "../scripts/benchmark-options.mjs";
import {
  DOCX_AGENT_BENCHMARK_SCENARIOS,
  benchmarkScenario,
} from "../scripts/docx-agent-benchmark-scenarios.mjs";

test("DOCX agent benchmark corpus covers the eight required reliability jobs", () => {
  assert.equal(DOCX_AGENT_BENCHMARK_SCENARIOS.length, 8);
  assert.equal(
    new Set(DOCX_AGENT_BENCHMARK_SCENARIOS.map((item) => item.category)).size,
    8,
  );
  for (const scenario of DOCX_AGENT_BENCHMARK_SCENARIOS) {
    assert.equal(benchmarkScenario(scenario.id), scenario);
    assert(scenario.instruction.length > 40);
    assert(scenario.expectedSignals.length >= 2);
  }
});

import { toolMetrics } from '../scripts/benchmark-tool-metrics.mjs';
test('benchmark distinguishes pagination and revisions from identical requests', () => {
  const event = input => ({ part: { tool: 'docx_read', state: { input } } });
  const result = toolMetrics([
    event({ view: { revision: 'one', documentId: 'doc' }, offset: 0 }),
    event({ offset: 0, view: { documentId: 'doc', revision: 'one' } }),
    event({ view: { revision: 'one', documentId: 'doc' }, offset: 30 }),
    event({ view: { revision: 'two', documentId: 'doc' }, offset: 0 }),
    event(undefined),
  ]);
  assert.equal(result.repeatedToolNames, 4);
  assert.equal(result.repeatedRequests, 1);
  assert.equal(result.requestsWithArguments, 4);
});

test("benchmark measures captured UTF-8 text bytes and reports incomplete outputs", () => {
  const input = { text: "合同 😀" };
  const output = "已保存 ✅";
  const result = toolMetrics([
    { part: { tool: "docx_apply", state: { input, output } } },
    { part: { tool: "docx_read", state: { input: {}, output: "" } } },
    { part: { tool: "docx_save", state: { input: null } } },
    { part: { tool: "docx_render", state: { output: { image: "uncounted" } } } },
  ]);
  assert.equal(result.metricVersion, 2);
  assert.equal(result.toolPayloadMetricVersion, 1);
  assert.equal(result.toolInputJsonBytes, Buffer.byteLength(JSON.stringify(input), "utf8") + 2 + 4);
  assert.equal(result.toolOutputTextBytes, Buffer.byteLength(output, "utf8"));
  assert.equal(result.toolOutputsWithText, 2, "empty captured text is present, not missing");
  assert.equal(result.toolOutputsMissing, 1);
  assert.equal(result.toolOutputsNonText, 1);
  assert.equal(result.toolInputJsonComplete, false);
  assert.equal(result.toolOutputTextComplete, false);
  assert.equal(toolMetrics([{ part: { tool: "docx_read", state: { input: {}, output: "" } } }]).toolOutputTextComplete, true);
});

test("benchmark source and labels are validated and evidence writes never overwrite", () => {
  const root = fs.mkdtempSync(path.join(tmpdir(), "officemcp-benchmark-options-"));
  try {
    for (const directory of ["src/contracts", "node_modules"]) fs.mkdirSync(path.join(root, directory), { recursive: true });
    fs.writeFileSync(path.join(root, "package.json"), JSON.stringify({ name: "officemcp-local-preview", type: "module" }));
    fs.writeFileSync(path.join(root, "src/cli.mjs"), "");
    fs.writeFileSync(path.join(root, "src/contracts/registry.mjs"), "");
    const selected = benchmarkOptions(root, { OFFICEMCP_BENCH_SOURCE: root, OFFICEMCP_BENCH_LABEL: "before-review" });
    assert.equal(selected.sourceRoot, fs.realpathSync(root));
    assert.equal(selected.label, "before-review");
    assert.throws(() => benchmarkOptions(root, { OFFICEMCP_BENCH_SOURCE: "./relative" }), /absolute/);
    assert.throws(() => benchmarkOptions(root, { OFFICEMCP_BENCH_LABEL: "../overwrite" }), /LABEL/);
    const report = benchmarkReportPath(root, "docx-agent-corpus", selected.label);
    writeBenchmarkReport(report, { preserved: true });
    assert.throws(() => writeBenchmarkReport(report, { preserved: false }), { code: "EEXIST" });
    assert.equal(JSON.parse(fs.readFileSync(report, "utf8")).preserved, true);
    assert.notEqual(benchmarkOptions(root, {}).label, benchmarkOptions(root, {}).label);
    fs.unlinkSync(path.join(root, "src/cli.mjs"));
    assert.throws(() => benchmarkOptions(root, {}), /missing src\/cli.mjs/);
  } finally { fs.rmSync(root, { recursive: true, force: true }); }
});
