import test from "./support/test.mjs";
import assert from "node:assert/strict";
import * as fs from "node:fs";
import path from "node:path";
import { tmpdir } from "node:os";
import { Client } from "@modelcontextprotocol/client";
import { StdioClientTransport } from "@modelcontextprotocol/client/stdio";
import { OfficeService } from "../src/application/service.mjs";
import { configFrom, digest } from "../src/infrastructure/files.mjs";
import { writeFixtures, docxBytes } from "./fixtures.mjs";
function fixture() {
  const root = fs.mkdtempSync(path.join(tmpdir(), "officemcp-test-"));
  const docs = path.join(root, "docs");
  fs.mkdirSync(docs);
  writeFixtures(docs);
  const config = path.join(root, "config.json");
  fs.writeFileSync(
    config,
    JSON.stringify({
      readRoots: [docs],
      writeRoots: { output: docs },
      stateDirectory: path.join(root, "state"),
    }),
  );
  const service = new OfficeService(configFrom(config));
  return {
    root,
    docs,
    config,
    service,
    close() {
      service.close();
      fs.rmSync(root, { recursive: true, force: true });
    },
  };
}
async function call(s, name, args = {}) {
  const r = await s.call(name, args);
  assert(!r.isError, JSON.stringify(r.payload));
  return r.payload;
}
async function opened(f, file = "contract.docx") {
  return (
    await call(f.service, "office_open", { path: path.join(f.docs, file) })
  ).document;
}
const edit = (d, op = "edit_1") => ({
  documentId: d.id,
  expectedSourceRevision: d.sourceRevision,
  operationId: op,
  paragraphId: "p_0",
  expectedText: "30 days",
  expectedMatches: 1,
  replacement: "45 days",
  preservationProfile: "docx.text-replacement.v1",
  changeMode: "direct",
});
test("M1 real stdio discovery, matching text/structured results, paths, snapshots and restart", async () => {
  const f = fixture();
  f.service.close();
  const client = new Client(
    { name: "independent-preview-test", version: "1" },
    {
      capabilities: {},
      versionNegotiation: { mode: { pin: "2026-07-28" } },
    },
  );
  const transport = new StdioClientTransport({
    command: process.execPath,
    args: [
      new URL("../src/cli.mjs", import.meta.url).pathname,
      "serve",
      "--config",
      f.config,
    ],
    stderr: "pipe",
  });
  let stderr = "";
  transport.stderr?.on("data", (b) => (stderr += b.toString()));
  try {
    await client.connect(transport);
    assert.equal(client.getProtocolEra(), "modern");
    const listing = await client.listTools();
    assert(listing.tools.some((t) => t.name === "office_server_info"));
    const info = await client.callTool({
      name: "office_server_info",
      arguments: {},
    });
    assert.deepEqual(JSON.parse(info.content[0].text), info.structuredContent);
    const denied = await client.callTool({
      name: "office_open",
      arguments: { path: "/etc/hosts" },
    });
    assert.equal(denied.isError, true);
    assert.equal(denied.structuredContent.error.code, "ACCESS_DENIED");
    const doc = await client.callTool({
      name: "office_open",
      arguments: { path: path.join(f.docs, "contract.docx") },
    });
    assert(doc.structuredContent.document.sourceRevision.startsWith("sha256:"));
    assert(!stderr.includes("Payment"));
  } finally {
    await client.close();
    fs.rmSync(f.root, { recursive: true, force: true });
  }
});
test("M1 confines symlinks, rejects shared live state, immutable revisions and invalid cursors", async () => {
  const f = fixture();
  try {
    fs.symlinkSync("/etc/hosts", path.join(f.docs, "escape.docx"));
    assert.equal(
      (
        await f.service.call("office_open", {
          path: path.join(f.docs, "escape.docx"),
        })
      ).payload.error.code,
      "ACCESS_DENIED",
    );
    assert.throws(() => new OfficeService(configFrom(f.config)), /live server/);
    const d = await opened(f);
    const a = {
      documentId: d.id,
      expectedRevision: d.sourceRevision,
      limit: 1,
    };
    const r = await call(f.service, "docx_read", a);
    assert(r.nextCursor);
    assert.equal(r.data.items[0].text, "Payment within 30 days");
    assert.equal(
      (
        await f.service.call("docx_read", {
          ...a,
          limit: 2,
          cursor: r.nextCursor,
        })
      ).payload.error.code,
      "CURSOR_CONFLICT",
    );
    assert.equal(
      (await call(f.service, "docx_read", { ...a, cursor: r.nextCursor }))
        .nextCursor,
      null,
    );
  } finally {
    f.close();
  }
});
test("M2 DOCX edit, independent preservation validation, save-as, duplicate calls and restart receipt", async () => {
  const f = fixture();
  try {
    const original = fs.readFileSync(path.join(f.docs, "contract.docx"));
    const d = await opened(f);
    const a = edit(d);
    const first = await call(f.service, "docx_replace_text", a);
    const c = first.data;
    assert.deepEqual((await call(f.service, "docx_replace_text", a)).data, c);
    assert.equal(
      (
        await f.service.call("docx_replace_text", {
          ...a,
          replacement: "60 days",
        })
      ).payload.error.code,
      "OPERATION_ID_CONFLICT",
    );
    assert.deepEqual(
      fs.readFileSync(path.join(f.docs, "contract.docx")),
      original,
    );
    const read = await call(f.service, "docx_read", {
      documentId: d.id,
      candidateId: c.candidateId,
      expectedRevision: c.candidateRevision,
    });
    assert.equal(read.data.items[0].text, "Payment within 45 days");
    assert.equal(
      read.data.items[1].text,
      "Unrelated paragraph must remain unchanged.",
    );
    const v = await call(f.service, "office_validate", {
      documentId: d.id,
      candidateId: c.candidateId,
      expectedCandidateRevision: c.candidateRevision,
    });
    assert(v.data.passed);
    const commit = {
      documentId: d.id,
      candidateId: c.candidateId,
      expectedSourceRevision: d.sourceRevision,
      expectedCandidateRevision: c.candidateRevision,
      validationId: v.data.validationId,
      operationId: "commit_1",
      mode: "save_as",
      destination: { rootId: "output", relativePath: "reviewed.docx" },
      overwrite: false,
    };
    const receipt = await call(f.service, "office_commit", commit);
    assert(receipt.data.published);
    assert.equal(
      digest(fs.readFileSync(path.join(f.docs, "reviewed.docx"))),
      c.candidateRevision,
    );
    assert.deepEqual(
      fs.readFileSync(path.join(f.docs, "contract.docx")),
      original,
    );
    f.service.close();
    f.service = new OfficeService(configFrom(f.config));
    assert.deepEqual(
      (await call(f.service, "office_commit", commit)).data,
      receipt.data,
    );
    assert(
      (
        await call(f.service, "office_operation_status", {
          operationId: "commit_1",
        })
      ).data.data.published,
    );
  } finally {
    f.close();
  }
});
test("M2 rejects stale validation, external source changes and overwrite", async () => {
  const f = fixture();
  try {
    const d = await opened(f);
    const first = (await call(f.service, "docx_replace_text", edit(d))).data;
    const v = (
      await call(f.service, "office_validate", {
        documentId: d.id,
        candidateId: first.candidateId,
        expectedCandidateRevision: first.candidateRevision,
      })
    ).data;
    const a = {
      ...edit(d, "edit_2"),
      candidateId: first.candidateId,
      expectedCandidateRevision: first.candidateRevision,
      expectedText: "45 days",
      replacement: "60 days",
    };
    const second = (await call(f.service, "docx_replace_text", a)).data;
    const commit = {
      documentId: d.id,
      candidateId: second.candidateId,
      expectedSourceRevision: d.sourceRevision,
      expectedCandidateRevision: second.candidateRevision,
      validationId: v.validationId,
      operationId: "bad_commit",
      mode: "save_as",
      destination: { rootId: "output", relativePath: "reviewed.docx" },
      overwrite: false,
    };
    assert.equal(
      (await f.service.call("office_commit", commit)).payload.error.code,
      "VALIDATION_FAILED",
    );
    const fresh = (
      await call(f.service, "office_validate", {
        documentId: d.id,
        candidateId: second.candidateId,
        expectedCandidateRevision: second.candidateRevision,
      })
    ).data;
    fs.writeFileSync(path.join(f.docs, "reviewed.docx"), "existing");
    assert.equal(
      (
        await f.service.call("office_commit", {
          ...commit,
          operationId: "exists",
          validationId: fresh.validationId,
        })
      ).payload.error.code,
      "DESTINATION_CONFLICT",
    );
    fs.writeFileSync(
      path.join(f.docs, "contract.docx"),
      docxBytes(["Changed by an external writer"]),
    );
    assert.equal(
      (
        await f.service.call("office_commit", {
          ...commit,
          operationId: "changed",
          destination: {rootId:"output",relativePath:"external-change.docx"},
          validationId: fresh.validationId,
        })
      ).payload.error.code,
      "SOURCE_REVISION_CONFLICT",
    );
  } finally {
    f.close();
  }
});
test("M2 reconciles a synthetic interrupted publication journal after restart", async () => {
  const f = fixture();
  try {
    const d = await opened(f);
    const c = (await call(f.service, "docx_replace_text", edit(d))).data;
    const candidate = f.service.get("candidate", c.candidateId);
    const dest = path.join(f.docs, "recovered.docx"),
      temp = path.join(f.docs, ".publication.tmp");
    fs.copyFileSync(candidate.path, temp);
    fs.linkSync(temp, dest);
    f.service.put("operation", "lost_commit", {
      name: "office_commit",
      status: "publishing",
      fingerprint: "fixture",
      publication: {
        dest,
        tmp: temp,
        hash: c.candidateRevision,
        candidateId: c.candidateId,
        documentId: d.id,
      },
    });
    f.service.close();
    f.service = new OfficeService(configFrom(f.config));
    assert(
      (
        await call(f.service, "office_operation_status", {
          operationId: "lost_commit",
        })
      ).data.data.published,
    );
    assert(!fs.existsSync(temp));
    assert.equal(digest(fs.readFileSync(dest)), c.candidateRevision);
  } finally {
    f.close();
  }
});
test("M3 PDF bounded positioned text and background jobs", async () => {
  const f = fixture();
  try {
    const d = await opened(f, "invoice.pdf");
    const a = { documentId: d.id, expectedRevision: d.sourceRevision };
    assert.equal((await call(f.service, "pdf_inspect", a)).data.pageCount, 1);
    const text = await call(f.service, "pdf_read", { ...a, page: 1 });
    assert(text.data.items.some((x) => x.text.includes("125 dollars")));
    assert.equal(text.coverage.completeForDocument, false);
    const job = await call(f.service, "pdf_read", {
      ...a, page: 1, background: true,
    });
    let result;
    for (let i = 0; i < 100; i++) {
      await new Promise((r) => setTimeout(r, 100));
      result = await call(f.service, "office_job_status", { jobId: job.jobId });
      if (result.status === "completed") break;
    }
    assert.equal(result.status, "completed");
    assert(result.data.data.items.some(item => item.text.includes("125 dollars")));
    assert.equal(
      (await f.service.call("pdf_read", { ...a, page: 2 })).payload.error.code,
      "TARGET_CONFLICT",
    );
  } finally {
    f.close();
  }
});
test("M3 worker deadlines are bounded and sources remain unchanged", async () => {
  const f = fixture();
  try {
    f.service.config.workerTimeoutMs = 1;
    const d = await opened(f);
    const r = await f.service.call("docx_replace_text", edit(d));
    assert(r.isError);
    assert.equal(r.payload.error.code, "WORKER_TIMEOUT");
    assert.equal(
      digest(fs.readFileSync(path.join(f.docs, "contract.docx"))),
      d.sourceRevision,
    );
  } finally {
    f.close();
  }
});
test("M2 rejects ambiguous run matches and validates unrelated package bytes independently", async () => {
  const { unzipSync, zipSync, strToU8 } = await import("fflate");
  const f = fixture();
  try {
    fs.writeFileSync(
      path.join(f.docs, "ambiguous.docx"),
      docxBytes(["30 days and 30 days"]),
    );
    const d = await opened(f, "ambiguous.docx");
    assert.equal(
      (await f.service.call("docx_replace_text", edit(d))).payload.error.code,
      "TARGET_CONFLICT",
    );
    const good = await opened(f);
    const c = (await call(f.service, "docx_replace_text", edit(good, "good")))
      .data;
    const row = f.service.get("candidate", c.candidateId);
    const parts = unzipSync(fs.readFileSync(row.path));
    parts["docProps/custom.xml"] = strToU8("<tampered/>");
    const bytes = Buffer.from(zipSync(parts));
    fs.writeFileSync(row.path, bytes);
    const revision = digest(bytes);
    f.service.put("candidate", c.candidateId, { ...row, revision });
    assert.equal(
      (
        await f.service.call("office_validate", {
          documentId: good.id,
          candidateId: c.candidateId,
          expectedCandidateRevision: revision,
        })
      ).payload.error.code,
      "VALIDATION_FAILED",
    );
  } finally {
    f.close();
  }
});
test("M3 rejects unreadable SDK bodies, archive traversal and expansion limits; cancellation is terminal", async () => {
  const { unzipSync, zipSync, strToU8 } = await import("fflate");
  const f = fixture();
  try {
    const parts = unzipSync(docxBytes());
    parts["word/document.xml"] = strToU8(
      '<!DOCTYPE x [<!ENTITY leak SYSTEM "file:///etc/hosts">]><x>&leak;</x>',
    );
    fs.writeFileSync(path.join(f.docs, "unsafe.docx"), zipSync(parts));
    let d = await opened(f, "unsafe.docx");
    const unsafeInspection = await f.service.call("docx_inspect", {
      documentId: d.id,
      expectedRevision: d.sourceRevision,
    });
    assert.equal(unsafeInspection.payload.error?.code, "INVALID_DOCUMENT", JSON.stringify(unsafeInspection.payload));
    const bad = unzipSync(docxBytes());
    bad["../escape.xml"] = strToU8("<bad/>");
    fs.writeFileSync(path.join(f.docs, "traversal.docx"), zipSync(bad));
    d = await opened(f, "traversal.docx");
    assert.equal(
      (
        await f.service.call("docx_inspect", {
          documentId: d.id,
          expectedRevision: d.sourceRevision,
        })
      ).payload.error.code,
      "INVALID_DOCUMENT",
    );
    d = await opened(f);
    f.service.config.maxExpandedBytes = 10;
    assert.equal(
      (
        await f.service.call("docx_inspect", {
          documentId: d.id,
          expectedRevision: d.sourceRevision,
        })
      ).payload.error.code,
      "LIMIT_EXCEEDED",
    );
    f.service.config.maxExpandedBytes = 64 * 1024 * 1024;
    const job = await call(f.service, "docx_inspect", {
      documentId: d.id,
      expectedRevision: d.sourceRevision,
      background: true,
    });
    await call(f.service, "office_job_cancel", { jobId: job.jobId });
    await new Promise((r) => setTimeout(r, 30));
    assert.equal(
      (await f.service.call("office_job_status", { jobId: job.jobId })).payload
        .error.code,
      "CANCELLED",
    );
  } finally {
    f.close();
  }
});
