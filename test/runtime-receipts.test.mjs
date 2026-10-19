import test from "./support/test.mjs";
import assert from "node:assert/strict";
import fs from "node:fs";
import { createServiceFixture } from "./service-fixture.mjs";

test("DOCX save recovery returns the published revision and expired mutation IDs stay expired", async () => {
  const s = createServiceFixture({ toolset: "full" });
  try {
    const created = await s.call("docx_create", { operationId: "create", document: { format: "docx",
      blocks: [{ kind: "paragraph", text: "Amount 125" }] } });
    const args = { view: created.view, operationId: "edit", action: "replace_text",
      arguments: { edits: [{ find: "125", replace: "150" }], caseSensitive: true } };
    const edited = await s.call("docx_apply", args);
    const saved = await s.call("docx_save", { view: edited.view, operationId: "save",
      destination: { rootId: "output", relativePath: "saved.docx" } });
    const recovered = await s.call("docx_operation_get", { operationId: "save" });
    assert.deepEqual(recovered.view, saved.view);
    assert.equal(recovered.view.revision, saved.data.outputRevision);
    assert.equal(recovered.data.path, saved.data.path);
    const completedAt = s.service.get("operation", "edit").updatedAt;
    s.service.now = () => completedAt + s.service.config.operationRetentionMs + 1;
    const expired = await s.service.call("docx_apply", args);
    assert.equal(expired.payload.error.code, "OPERATION_EXPIRED");
    assert.equal(s.service.get("operation", "edit").status, "expired");
  } finally { s.close(); }
});

test("an image response that exceeds its budget does not leave a broken asset handle", async () => {
  const s = createServiceFixture({ toolset: "full" });
  try {
    const created = await s.call("docx_create", { operationId: "create", document: { format: "docx",
      blocks: [{ kind: "paragraph", text: "Artifact fixture" }] } });
    const assetsBefore = s.service.rows("asset").length;
    let output;
    s.service.worker = async (_kind, _input, _args, context) => {
      output = context.output;
      fs.writeFileSync(output, Buffer.alloc(8 * 1024 * 1024 + 1));
      return { mediaType: "image/png" };
    };
    await assert.rejects(s.service.artifacts.produce({ kind: "office_extract", input: "fixture",
      documentId: created.view.documentId, mediaType: "image/png", includeImage: true }),
      error => error.code === "LIMIT_EXCEEDED");
    assert.equal(s.service.rows("asset").length, assetsBefore);
    assert(!fs.existsSync(output));
  } finally { s.close(); }
});
