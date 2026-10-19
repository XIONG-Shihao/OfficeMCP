import test from "./support/test.mjs";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { createHash } from "node:crypto";
import { unzipSync, strFromU8 } from "fflate";
import { DOMParser } from "@xmldom/xmldom";
import { createServiceFixture } from "./service-fixture.mjs";
import { OfficeService } from "../src/application/service.mjs";
import { configFrom } from "../src/infrastructure/files.mjs";
import { createWorkbook, editWorkbook } from "#document-core/xlsx/workbook-package.mjs";
import { boundWorkbookMutationEvidence } from "../src/application/workbook-mutation-evidence.mjs";

const sha256 = value => createHash("sha256").update(value).digest("hex");
const bytes = value => Buffer.byteLength(JSON.stringify(value));
function omitted(record, field, original) {
  assert.equal(record[field], undefined);
  assert.deepEqual(record.omittedFields[field], {
    omitted: true, type: "string", chars: [...original].length,
    bytes: Buffer.byteLength(original), sha256: sha256(original), encoding: "utf-8",
  });
}
function calls(fixture) {
  let service = fixture.service;
  return {
    get service() { return service; },
    async call(name, args) {
      const result = await service.call(name, args);
      assert(!result.isError, `${name}: ${JSON.stringify(result.payload)}`);
      assert(bytes(result.payload) * 2 < service.config.maxResponseBytes);
      return result.payload;
    },
    restart() { service.close(); service = new OfficeService(configFrom(fixture.config)); },
    close() { service.close(); fs.rmSync(fixture.root, { recursive: true, force: true }); },
  };
}
async function open(f, file) {
  const result = await f.call("office_open", { path: file });
  return { documentId: result.document.id, revision: result.document.sourceRevision };
}

test("long CSV literals complete with bounded durable evidence, reusable targets and exact saved bytes", async () => {
  const before = "A".repeat(40000), after = "B".repeat(40000);
  const source = `id,note\r\n001,${before}\r\n`;
  const fixture = createServiceFixture({ prepare: docs => fs.writeFileSync(path.join(docs, "long.csv"), source) });
  const f = calls(fixture);
  try {
    const file = path.join(fixture.docs, "long.csv"), view = await open(f, file);
    const args = { view, operationId: "long-csv", commands: [{ sheet: "CSV", cell: "B2", expectedValue: before, expectedRawValue: before, value: after }] };
    const changed = await f.call("xlsx_apply", args);
    assert.equal(changed.status, "completed");
    assert.equal(changed.data.postconditions.affectedCellCount, 1);
    assert.deepEqual(changed.data.evidence.affectedTargets, { total: 1, returned: 1, omitted: 0 });
    const item = changed.data.affectedTargets[0];
    omitted(item.before, "value", before); omitted(item.after, "value", after);
    assert.equal(item.target.revision, changed.view.revision);
    f.restart();
    const recovered = await f.call("office_operation_status", { operationId: args.operationId });
    assert.deepEqual(recovered.data.data.evidence, changed.data.evidence);
    const replay = await f.call("xlsx_apply", args);
    assert.deepEqual(replay.data, JSON.parse(JSON.stringify(changed.data))); assert.deepEqual(replay.view, changed.view);
    const next = await f.call("xlsx_apply", { view: replay.view, operationId: "long-csv-target",
      commands: [{ target: replay.data.affectedTargets[0].target, expectedValue: after, value: after + "!" }] });
    const saved = await f.call("office_save", { view: next.view, operationId: "save-long-csv",
      destination: { rootId: "output", relativePath: "saved.csv" } });
    assert.deepEqual(fs.readFileSync(saved.data.path), Buffer.from(`id,note\r\n001,${after}!\r\n`));
    assert.deepEqual(fs.readFileSync(file), Buffer.from(source));
  } finally { f.close(); }
});

test("many long XLSX literals preserve all edits while receipt item count follows the response budget", async () => {
  const old = "O".repeat(30000), replacement = "N".repeat(30000), count = 24;
  const original = editWorkbook(createWorkbook(), Array.from({ length: count }, (_, i) => ({ sheet: "Sheet1", cell: `A${i + 1}`, value: old + i })));
  const fixture = createServiceFixture({ limits: { maxResponseBytes: 16000 },
    prepare: docs => fs.writeFileSync(path.join(docs, "long.xlsx"), original) });
  const f = calls(fixture);
  try {
    const file = path.join(fixture.docs, "long.xlsx"), view = await open(f, file);
    const args = { view, operationId: "long-xlsx", commands: Array.from({ length: count }, (_, i) => ({
      sheet: "Sheet1", cell: `A${i + 1}`, expectedValue: old + i, value: replacement + i,
    })) };
    const changed = await f.call("xlsx_apply", args), returned = changed.data.affectedTargets.length;
    assert(returned > 0 && returned < count);
    assert.equal(changed.data.postconditions.affectedCellCount, count);
    assert.equal(changed.data.postconditions.affectedCellsTruncated, true);
    assert.deepEqual(changed.data.evidence.affectedTargets, { total: count, returned, omitted: count - returned });
    omitted(changed.data.affectedTargets[0].before, "value", old + "0");
    omitted(changed.data.affectedTargets[0].after, "value", replacement + "0");
    f.restart();
    const recovered = await f.call("office_operation_status", { operationId: args.operationId });
    assert.equal(recovered.data.data.postconditions.affectedCellCount, count);
    const replay = await f.call("xlsx_apply", args);
    assert.deepEqual(replay.data, JSON.parse(JSON.stringify(changed.data)));
    const saved = await f.call("office_save", { view: replay.view, operationId: "save-long-xlsx",
      destination: { rootId: "output", relativePath: "saved.xlsx" } });
    const parts = unzipSync(fs.readFileSync(saved.data.path));
    const xml = new DOMParser().parseFromString(strFromU8(parts["xl/worksheets/sheet1.xml"]), "application/xml");
    const values = new Map(Array.from(xml.getElementsByTagName("c"), cell => [cell.getAttribute("r"), cell.getElementsByTagName("t")[0]?.textContent]));
    for (let i = 0; i < count; i++) assert.equal(values.get(`A${i + 1}`), replacement + i);
    for (const [name, part] of Object.entries(unzipSync(original))) {
      if (!["xl/workbook.xml", "xl/worksheets/sheet1.xml"].includes(name)) assert.deepEqual(parts[name], part, name);
    }
    assert.deepEqual(fs.readFileSync(file), Buffer.from(original));
    const oversizedRead = await f.service.call("xlsx_query", { view: replay.view, sheet: "Sheet1", range: "A1", limit: 1 });
    assert.equal(oversizedRead.payload.error.code, "LIMIT_EXCEEDED");
  } finally { f.close(); }
});

test("evidence preserves Unicode hashes, full counts and immutable small receipts", () => {
  const value = "中文🙂\\\n".repeat(1000);
  const result = { operationId: "unicode", data: { candidateId: "c", candidateRevision: "r", sourceRevision: "s", applied: true,
    affectedTargets: [{ sheet: "S", cell: "A1", after: { value } }], postconditions: { affectedCellCount: 101, affectedCellsTruncated: true } } };
  const original = structuredClone(result);
  const bounded = boundWorkbookMutationEvidence(result, { documentId: "d", maxResponseBytes: 16000 });
  omitted(bounded.data.affectedTargets[0].after, "value", value);
  assert.deepEqual(bounded.data.evidence.affectedTargets, { total: 101, returned: 1, omitted: 100 });
  assert.deepEqual(result, original);
  const small = structuredClone(result); small.data.affectedTargets[0].after.value = "small";
  small.data.postconditions = { affectedCellCount: 1, affectedCellsTruncated: false };
  assert.deepEqual(boundWorkbookMutationEvidence(small, { documentId: "d", maxResponseBytes: 200000 }), small);
});

test("an insufficient mandatory receipt budget fails before candidate acceptance", async () => {
  const f = createServiceFixture({ prepare: docs => fs.writeFileSync(path.join(docs, "tiny.csv"), "id,value\n1,old\n") });
  try {
    const view = await open(f, path.join(f.docs, "tiny.csv"));
    f.service.config.maxResponseBytes = 1000;
    const failed = await f.service.call("xlsx_apply", { view, operationId: "too-small-receipt", commands: [{ sheet: "CSV", cell: "B2", value: "new" }] });
    assert.equal(failed.payload.error.code, "LIMIT_EXCEEDED");
    assert.match(failed.payload.error.message, /candidate was not accepted/);
    assert.equal(f.service.rows("candidate").length, 0);
    assert.equal(f.service.get("operation", "too-small-receipt").status, "failed");
    assert.equal(fs.readFileSync(path.join(f.docs, "tiny.csv"), "utf8"), "id,value\n1,old\n");
  } finally { f.close(); }
});
