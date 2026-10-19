import test from "./support/test.mjs";
import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { clientFor, saveChecked } from "../src/formats/docx/docx-sdk.mjs";
import { Document, Packer, Paragraph } from "docx";
import { digest } from "../src/infrastructure/files.mjs";

function fixture({ changed = false, openFailure = false } = {}) {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "docx-save-lifecycle-"));
  const input = path.join(root, "source.docx"), output = path.join(root, "candidate.docx");
  fs.writeFileSync(input, "immutable source");
  const events = [];
  const makeDocument = (name, text) => {
    const doc = { name, closes: 0,
      get: async () => { events.push(`${name}:snapshot`); return { body: [{ kind: "paragraph", paragraph: { inlines: [{ kind: "run", run: { text } }] } }] }; },
      save: async () => { events.push(`${name}:save`); fs.writeFileSync(output, "saved candidate", { flag: "wx" }); },
      close: async () => { events.push(`${name}:close`); doc.closes++; },
      headerFooters: { parts: { list: async () => ({ items: [], total: 0 }) } },
    };
    for (const domain of ["images", "sections", "footnotes", "fields", "hyperlinks", "contentControls", "trackChanges", "comments", "captions", "crossRefs"])
      doc[domain] = { list: async () => ({ items: [], total: 0 }) };
    return doc;
  };
  const original = makeDocument("original", "Exact content");
  const reopened = makeDocument("reopened", changed ? "Unexpected content" : "Exact content");
  let owned = original;
  const client = { open: async (args) => {
    events.push("client:open");
    assert.equal(args.doc, output);
    assert.equal(args.userName, "Editor");
    if (openFailure) throw new Error("reopen failed");
    return reopened;
  } };
  return { original, reopened, client, events, input, output,
    get owned() { return owned; },
    setDocument: next => { events.push(next ? "owner:verified" : "owner:released"); owned = next; },
    task: { input, output, maxFileBytes: 1024, args: { steps: [] } },
    cleanup: () => fs.rmSync(root, { recursive: true, force: true }),
  };
}

test("sequential DOCX save closes the old document before reopening and transfers only a verified handle", async () => {
  const f = fixture();
  try {
    const inputRevision = digest(fs.readFileSync(f.input));
    const result = await saveChecked(f.client, f.original, f.task, inputRevision, {
      sequentialReopen: true, setDocument: f.setDocument, userName: "Editor",
    });
    assert.deepEqual(f.events, ["original:snapshot", "original:save", "original:close", "owner:released", "client:open", "reopened:snapshot", "owner:verified"]);
    assert.equal(f.owned, f.reopened);
    assert.equal(f.original.closes, 1);
    assert.equal(f.reopened.closes, 0);
    assert.equal(result.engineEvidence.inputRevision, inputRevision);
    assert.equal(result.revision, digest(fs.readFileSync(f.output)));
    assert.equal(result.engineEvidence.outputRevision, result.revision);
    assert.deepEqual(Object.keys(result).sort(), ["engine", "engineEvidence", "revision"]);
    assert.equal(digest(fs.readFileSync(f.input)), inputRevision);
    await f.owned.close({ discard: true }); // Creation/host cleanup owns exactly this handle now.
    assert.equal(f.original.closes, 1);
    assert.equal(f.reopened.closes, 1);
  } finally { f.cleanup(); }
});

test("sequential DOCX verification failure closes the rejected handle without adopting it", async () => {
  const f = fixture({ changed: true });
  try {
    await assert.rejects(saveChecked(f.client, f.original, f.task, digest(fs.readFileSync(f.input)), {
      sequentialReopen: true, setDocument: f.setDocument, userName: "Editor",
    }), { code: "VALIDATION_FAILED" });
    assert.equal(f.owned, null);
    assert.equal(f.original.closes, 1);
    assert.equal(f.reopened.closes, 1);
    assert.equal(f.events.includes("owner:verified"), false);
    assert.equal(fs.readFileSync(f.input, "utf8"), "immutable source");
  } finally { f.cleanup(); }
});

test("a failed sequential reopen leaves no old or unverified document owned by the caller", async () => {
  const f = fixture({ openFailure: true });
  try {
    await assert.rejects(saveChecked(f.client, f.original, f.task, null, {
      sequentialReopen: true, setDocument: f.setDocument, userName: "Editor",
    }), /reopen failed/);
    assert.equal(f.owned, null);
    assert.equal(f.original.closes, 1);
    assert.equal(f.reopened.closes, 0);
  } finally { f.cleanup(); }
});

test("small DOCX saves retain the original session and close their verification session", async () => {
  const f = fixture();
  try {
    await saveChecked(f.client, f.original, f.task, null, { setDocument: f.setDocument, userName: "Editor" });
    assert.equal(f.owned, f.original);
    assert.equal(f.original.closes, 0);
    assert.equal(f.reopened.closes, 1);
    assert.equal(f.events.some(event => event.startsWith("owner:")), false);
  } finally { f.cleanup(); }
});

test("a sequentially reopened real SDK document remains editable and keeps source bytes unchanged", async () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "docx-sequential-sdk-"));
  const input = path.join(root, "source.docx");
  const source = await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph("Source text")] }] }));
  fs.writeFileSync(input, source);
  const task = { input, output: path.join(root, "first.docx"), scratch: path.join(root, "sdk"), timeoutMs: 30000, maxFileBytes: 8 * 1024 * 1024, args: { steps: [] } };
  const client = clientFor(task);
  let doc;
  try {
    doc = await client.open({ doc: input, runtime: "v2", userName: "Editor" });
    const first = await saveChecked(client, doc, task, digest(source), {
      sequentialReopen: true, userName: "Editor", setDocument: next => { doc = next; },
    });
    const found = await doc.query.match({ select: { type: "text", pattern: "Source text" }, require: "exactlyOne" });
    await doc.replace({ target: found.items[0].target, text: "Edited through adopted session", changeMode: "direct" });
    const secondTask = { ...task, input: task.output, output: path.join(root, "second.docx") };
    const second = await saveChecked(client, doc, secondTask, first.revision, {
      sequentialReopen: true, userName: "Editor", setDocument: next => { doc = next; },
    });
    assert.equal(second.engineEvidence.inputRevision, first.revision);
    assert.match(await doc.getText({}), /Edited through adopted session/);
    assert.equal(second.revision, digest(fs.readFileSync(secondTask.output)));
    assert.deepEqual(fs.readFileSync(input), source);
  } finally {
    try { if (doc) await doc.close({ discard: true }); }
    finally { await client.dispose(); fs.rmSync(root, { recursive: true, force: true }); }
  }
});
