import test from "./support/test.mjs";
import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { DatabaseSync } from "./support/raw-sqlite.mjs";
import { OfficeService } from "../src/application/service.mjs";
import { CONTRACT } from "../src/contracts/registry.mjs";
import { configFrom, digest } from "../src/infrastructure/files.mjs";

test("frozen 0.12 release state and receipts survive upgrade without repeating edits or publication", async () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "officemcp-state-upgrade-"));
  const stored = fs.readFileSync(new URL("./fixtures/state-0.12.0.json", import.meta.url), "utf8");
  const fixture = JSON.parse(stored.replaceAll("$ROOT", root));
  assert.equal(fixture.provenance.version, "0.12.0-preview.1");
  assert.match(fixture.provenance.commit, /^[a-f0-9]{40}$/);
  let service;
  try {
    fs.mkdirSync(path.join(root, "docs"));
    fs.mkdirSync(path.join(root, "state"));
    for (const file of fixture.files) {
      const target = path.join(root, file.path), bytes = Buffer.from(file.base64, "base64");
      assert.equal(digest(bytes), file.sha256, "Frozen source bytes must retain their recorded digest");
      fs.mkdirSync(path.dirname(target), { recursive: true }); fs.writeFileSync(target, bytes);
    }
    const db = new DatabaseSync(path.join(root, "state/state.sqlite"));
    db.exec("CREATE TABLE records (kind TEXT NOT NULL,id TEXT NOT NULL,value TEXT NOT NULL,PRIMARY KEY(kind,id))");
    for (const record of fixture.records)
      db.prepare("INSERT INTO records VALUES (?,?,?)").run(record.kind, record.id, JSON.stringify(record.value));
    db.close();
    const config = path.join(root, "config.json");
    fs.writeFileSync(config, JSON.stringify({ readRoots: [path.join(root, "docs")],
      writeRoots: { output: path.join(root, "docs") }, stateDirectory: path.join(root, "state") }));
    const now = Math.max(...fixture.records.map(record => record.value.updatedAt ?? 0)) + 1000;
    service = new OfficeService(configFrom(config), { now: () => now });
    const worker = service.worker.bind(service);
    service.worker = () => { throw new Error("Historical retries must not invoke a mutation engine"); };
    for (const name of ["docx_replace_text", "office_commit"]) {
      const prior = fixture.calls.find(call => call.name === name);
      const current = await service.call(name, prior.args);
      assert(!current.isError, JSON.stringify(current));
      const { requestId: _old, contractVersion: oldContract, ...expected } = prior.payload;
      assert.equal(oldContract, "0.12");
      const { requestId: _new, contractVersion, ...actual } = current.payload;
      assert.equal(contractVersion, CONTRACT);
      assert.deepEqual(actual, expected);
    }
    service.worker = worker;
    const published = fixture.calls.find(call => call.name === "office_commit").payload.data;
    assert.equal(digest(fs.readFileSync(published.path)), published.outputRevision);
    const opened = await service.call("office_open", { path: published.path });
    assert(!opened.isError, JSON.stringify(opened));
    const document = opened.payload.document;
    const read = await service.call("docx_structure", { documentId: document.id, expectedRevision: document.sourceRevision });
    assert(!read.isError, JSON.stringify(read));
    assert(read.payload.data.nodes.some(node => node.text.includes("45 days")));
    assert.equal(service.rows("operation").length, 2);
    assert.equal(service.rows("candidate").length, 1);
  } finally { service?.close(); fs.rmSync(root, { recursive: true, force: true }); }
});
