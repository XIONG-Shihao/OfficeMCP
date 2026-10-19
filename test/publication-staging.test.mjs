import test from "./support/test.mjs";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { PDFDocument } from "pdf-lib";
import { createServiceFixture } from "./service-fixture.mjs";
import { OfficeService } from "../src/application/service.mjs";
import { digest, fileRevision } from "../src/infrastructure/files.mjs";

async function setup() {
  const fixture = createServiceFixture({ toolset: "full" });
  const { view } = await fixture.call("docx_create", { operationId: "create",
    document: { format: "docx", blocks: [{ kind: "paragraph", text: "Publication staging fixture" }] } });
  const pdf = await PDFDocument.create(); pdf.addPage();
  const bytes = Buffer.from(await pdf.save());
  const artifact = path.join(fixture.service.config.stateDirectory, "exports", "artifact.pdf");
  const args = { view, operationId: "export", format: "pdf", destination: { rootId: "output", relativePath: "result.pdf" } };
  return { ...fixture, view, bytes, artifact, args };
}

test("PDF export stages on the destination filesystem and replays the same receipt", async (t) => {
  const s = await setup();
  try {
    let produced = 0;
    const link = fs.linkSync;
    t.mock.method(fs, "linkSync", (source, destination) => {
      // Portable simulation of distinct state/output devices: linking across them must fail.
      if (path.dirname(source) !== path.dirname(destination))
        throw Object.assign(new Error("Cross-device link"), { code: "EXDEV" });
      return link(source, destination);
    });
    const produce = () => { produced++; fs.writeFileSync(s.artifact, s.bytes); return {
      file: s.artifact, mediaType: "application/pdf", engine: "fixture", notVerified: ["layout"] };
    };
    const result = await s.service.documents.exportArtifact("docx_export", s.args, "pdf", produce);
    assert.equal(result.data.published, true);
    assert.equal(fileRevision(result.data.path, s.bytes.length), digest(s.bytes));
    assert.deepEqual(await s.service.documents.exportArtifact("docx_export", s.args, "pdf", produce), result);
    assert.equal(produced, 1);
    assert(!fs.existsSync(s.artifact));
    assert.deepEqual(fs.readdirSync(s.docs).sort(), ["result.pdf", "result.pdf.artifact.json"]); // Directory order is not defined.
    assert.deepEqual(JSON.parse(fs.readFileSync(result.data.artifactDescriptor,"utf8")), {
      schema:"mcp-artifact/1",sha256:digest(s.bytes).replace(/^sha256:/,""),inputs:[]
    });
    await assert.rejects(s.service.documents.exportArtifact("docx_export", { ...s.args, operationId: "another" }, "pdf", produce), { code: "DESTINATION_CONFLICT" });
    assert.equal(produced, 1);
  } finally { s.close(); }
});

test("failed staging leaves no output or temporary file", async () => {
  const s = await setup();
  try {
    fs.writeFileSync(s.artifact, s.bytes);
    assert.throws(() => s.service.documents.stagePublication(s.artifact, path.join(s.docs, "result.pdf"), digest("wrong bytes")), { code: "INTEGRITY_FAILED" });
    assert.deepEqual(fs.readdirSync(s.docs), []);
    assert.equal(fileRevision(s.artifact, s.bytes.length), digest(s.bytes));
  } finally { s.close(); }
});

test("converted DOCX artifacts use the hosted publication port and replay its receipt", async () => {
  const s = await setup();
  try {
    const candidate = s.service.get("candidate", s.view.candidateId);
    const artifact = path.join(s.root, "converted.docx");
    const destination = path.join(fs.realpathSync(s.docs), "converted.docx");
    const args = { ...s.args, operationId: "convert", format: "docx", destination: { rootId: "output", relativePath: "converted.docx" } };
    let admissions = 0, publications = 0, produced = 0;
    s.service.publication = {
      async admit(input) { admissions++; assert.equal(input.destination, destination); assert.equal(input.document, null); return "host-ticket"; },
      async publish(ticket, staged) { publications++; assert.equal(ticket, "host-ticket"); fs.copyFileSync(staged, destination, fs.constants.COPYFILE_EXCL); return { revision: fileRevision(destination, s.service.config.maxFileBytes).replace(/^sha256:/, "") }; },
      async cancel() {},
    };
    const produce = () => { produced++; fs.copyFileSync(candidate.path, artifact); return { file: artifact, mediaType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", engine: "fixture", notVerified: ["layout"] }; };
    const result = await s.service.documents.exportArtifact("office_convert", args, "docx", produce);
    assert(result.data.published);
    assert.equal(fileRevision(destination, s.service.config.maxFileBytes), candidate.revision);
    assert.deepEqual(await s.service.documents.exportArtifact("office_convert", args, "docx", produce), result);
    assert.equal(admissions, 1); assert.equal(publications, 1); assert.equal(produced, 1);
    assert(!fs.existsSync(artifact));
  } finally { s.close(); }
});

test("journal failure and a concurrent destination do not leak unowned staging files", async (t) => {
  const s = await setup();
  try {
    const produce = () => { fs.writeFileSync(s.artifact, s.bytes); return { file: s.artifact, mediaType: "application/pdf" }; };
    const publishing = s.service.durable.publishing.bind(s.service.durable);
    s.service.durable.publishing = () => { throw new Error("Injected journal failure"); };
    await assert.rejects(s.service.documents.exportArtifact("docx_export", s.args, "pdf", produce), /Injected journal failure/);
    assert.deepEqual(fs.readdirSync(s.docs), []);
    s.service.durable.publishing = publishing;
    t.mock.method(fs, "linkSync", (_source, destination) => {
      fs.writeFileSync(destination, "Concurrent owner", { flag: "wx" });
      throw Object.assign(new Error("Already exists"), { code: "EEXIST" });
    });
    const request = { ...s.args, operationId: "concurrent" };
    await assert.rejects(s.service.documents.exportArtifact("docx_export", request, "pdf", produce), { code: "DESTINATION_CONFLICT" });
    assert.deepEqual(fs.readdirSync(s.docs), ["result.pdf"]);
    assert.equal(fs.readFileSync(path.join(s.docs, "result.pdf"), "utf8"), "Concurrent owner");
    assert.equal(s.service.get("operation", request.operationId).error.code, "DESTINATION_CONFLICT");
    await assert.rejects(s.service.documents.exportArtifact("docx_export", request, "pdf", () => { throw new Error("Must not rerun"); }), { code: "DESTINATION_CONFLICT" });
  } finally { s.close(); }
});

for (const phase of ["publication-before-link", "publication-after-link", "publication-after-receipt"])
  test(`PDF export restart recovery at ${phase}`, async () => {
    const s = await setup();
    let service = s.service;
    try {
      fs.writeFileSync(s.artifact, s.bytes);
      const request = path.join(s.root, "export-request.json");
      fs.writeFileSync(request, JSON.stringify({ args: s.args, file: s.artifact }));
      service.close();
      const killed = spawnSync(process.execPath,
        [new URL("./fixtures/export-crash-child.mjs", import.meta.url).pathname, s.config, phase, request],
        { encoding: "utf8", timeout: 15000 });
      assert.equal(killed.signal, "SIGKILL", killed.stderr);
      service = new OfficeService(service.config);
      const operation = service.get("operation", "export");
      if (phase === "publication-before-link") {
        assert.equal(operation.status, "failed");
        assert.equal(operation.error.code, "PUBLICATION_UNCERTAIN");
        assert(!fs.existsSync(path.join(s.docs, "result.pdf")));
      } else {
        assert.equal(operation.status, "completed");
        assert(fs.existsSync(operation.result.data.artifactDescriptor),"Recovery must publish the portable descriptor before completing its receipt");
        assert.equal(fileRevision(operation.result.data.path, s.bytes.length), digest(s.bytes));
        assert.deepEqual(await service.documents.exportArtifact("docx_export", s.args, "pdf", () => { throw new Error("Must replay"); }), operation.result);
        assert(!fs.existsSync(operation.publication.tmp));
      }
    } finally { service.close(); fs.rmSync(s.root, { recursive: true, force: true }); }
  });
