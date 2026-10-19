import test from "./support/test.mjs";
import assert from "node:assert/strict";
import * as fs from "node:fs";
import path from "node:path";
import { tmpdir } from "node:os";
import { OfficeClient } from "../src/sdk.mjs";
import { SQLiteObjectStore } from "../src/storage.mjs";
import { docxBytes } from "./fixtures.mjs";
import { digest } from "../src/infrastructure/files.mjs";
test("M8 embedded API and SQLite storage share mutation/validation receipts, no-overwrite and restart semantics", async () => {
  const root = fs.mkdtempSync(path.join(tmpdir(), "office-sdk-"));
  let store = new SQLiteObjectStore(path.join(root, "objects.sqlite")),
    sdk = new OfficeClient({
      directory: path.join(root, "office"),
      storage: store,
    });
  const originalBytes = docxBytes();
  try {
    assert.equal(
      (await sdk.call("office_server_info")).payload.data.transport,
      "embedded",
    );
    await store.putIfAbsent({
      key: "source",
      bytes: originalBytes,
      operationId: "seed",
    });
    const source = await sdk.importSource("source", "docx");
    const call = async (name, args) => {
      const r = await sdk.call(name, args);
      assert(!r.isError, JSON.stringify(r));
      return r.payload;
    };
    const d = (await call("office_open", { assetId: source.assetId })).document;
    const c = (
      await call("docx_replace_text", {
        documentId: d.id,
        expectedSourceRevision: d.sourceRevision,
        operationId: "edit",
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
    };
    const v = (await call("office_validate", b)).data;
    const out = (
      await call("office_commit", {
        ...b,
        expectedSourceRevision: d.sourceRevision,
        validationId: v.validationId,
        operationId: "commit",
        mode: "save_as",
        destination: { rootId: "output", relativePath: "out.docx" },
        overwrite: false,
      })
    ).data;
    const args = {
      assetId: out.assetId,
      key: "reviewed",
      operationId: "export",
    };
    const receipt = await sdk.exportAsset(args);
    assert.equal((await store.read("reviewed")).revision, out.outputRevision);
    assert.equal((await store.read("source")).revision, digest(originalBytes));
    sdk.close();
    store.close();
    store = new SQLiteObjectStore(path.join(root, "objects.sqlite"));
    sdk = new OfficeClient({
      directory: path.join(root, "office"),
      storage: store,
    });
    assert.deepEqual(await sdk.exportAsset(args), receipt);
    await assert.rejects(sdk.exportAsset({ ...args, key: "different" }), {
      code: "OPERATION_ID_CONFLICT",
    });
    await assert.rejects(sdk.exportAsset({ ...args, operationId: "other" }), {
      code: "DESTINATION_CONFLICT",
    });
    // Simulate lost adapter reply: the same storage transaction receipt remains recoverable.
    const bytes = (await store.read("reviewed")).bytes;
    const saved = await store.putIfAbsent({
      key: "lost-reply",
      bytes,
      operationId: "lost",
    });
    assert.deepEqual(
      await store.putIfAbsent({
        key: "lost-reply",
        bytes,
        operationId: "lost",
      }),
      saved,
    );
  } finally {
    sdk.close();
    store.close();
    fs.rmSync(root, { recursive: true, force: true });
  }
});

test("M8 zero-byte objects and remote upload metadata remain quota bounded", async () => {
  const root = fs.mkdtempSync(path.join(tmpdir(), "office-empty-quota-"));
  const store = new SQLiteObjectStore(path.join(root, "objects.sqlite"), {
    maxObjects: 1,
  });
  const sdk = new OfficeClient({
    directory: path.join(root, "office"),
    storage: store,
    limits: { maxStateBytes: 1 },
  });
  try {
    await store.putIfAbsent({
      key: "empty",
      bytes: Buffer.alloc(0),
      operationId: "first",
    });
    await assert.rejects(
      store.putIfAbsent({
        key: "another",
        bytes: Buffer.alloc(0),
        operationId: "second",
      }),
      { code: "LIMIT_EXCEEDED" },
    );
    await assert.rejects(sdk.importSource("empty", "docx"), {
      code: "LIMIT_EXCEEDED",
    });
  } finally {
    sdk.close();
    store.close();
    fs.rmSync(root, { recursive: true, force: true });
  }
});
