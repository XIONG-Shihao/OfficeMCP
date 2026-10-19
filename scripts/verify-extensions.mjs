// Copied to the clean install directory by verify-package; resolves only installed public exports.
import { OfficeClient } from "officemcp-local-preview";
import { SQLiteObjectStore } from "officemcp-local-preview/storage";
import { createHash } from "node:crypto";
import * as fs from "node:fs";
import path from "node:path";
import assert from "node:assert/strict";
const root = process.argv[2],
  docs = process.argv[3],
  hash = (b) => "sha256:" + createHash("sha256").update(b).digest("hex");
const store = new SQLiteObjectStore(path.join(root, "sdk-objects.sqlite"));
const sdk = new OfficeClient({
  directory: path.join(root, "sdk-office"),
  storage: store,
});
async function workflow(call, source) {
  const d = (await call("office_open", { assetId: source.assetId })).document;
  const c = (
    await call("docx_replace_text", {
      documentId: d.id,
      expectedSourceRevision: d.sourceRevision,
      operationId: "ext-edit",
      paragraphId: "p_0",
      expectedText: "30 days",
      replacement: "45 days",
      expectedMatches: 1,
      changeMode: "direct",
      preservationProfile: "docx.text-replacement.v1",
    })
  ).data;
  const b = {
      documentId: d.id,
      candidateId: c.candidateId,
      expectedCandidateRevision: c.candidateRevision,
    },
    v = (await call("office_validate", b)).data;
  return (
    await call("office_commit", {
      ...b,
      expectedSourceRevision: d.sourceRevision,
      validationId: v.validationId,
      operationId: "ext-save",
      mode: "save_as",
      destination: { rootId: "output", relativePath: "result.docx" },
      overwrite: false,
    })
  ).data;
}
const bytes = fs.readFileSync(path.join(docs, "contract.docx"));
await store.putIfAbsent({ key: "source", bytes, operationId: "seed" });
const embedded = await workflow(
  async (name, args) => {
    const r = await sdk.call(name, args);
    assert(!r.isError, JSON.stringify(r));
    return r.payload;
  },
  await sdk.importSource("source", "docx"),
);
await sdk.exportAsset({
  assetId: embedded.assetId,
  key: "reviewed",
  operationId: "export",
});
assert.equal((await store.read("reviewed")).revision, embedded.outputRevision);
sdk.close();
store.close();
console.log("PASS installed local SDK and SQLite storage adapter");
