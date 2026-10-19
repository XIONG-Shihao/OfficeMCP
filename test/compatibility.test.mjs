import test from "./support/test.mjs";
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { createInterface } from "node:readline";
import * as fs from "node:fs";
import path from "node:path";
import { tmpdir } from "node:os";
import { writeFixtures } from "./fixtures.mjs";
for (const protocolVersion of ["2025-06-18", "2025-11-25"])
  test(
    "M6 independent text-only JSON-RPC client " + protocolVersion,
    async () => {
      const root = fs.mkdtempSync(path.join(tmpdir(), "office-wire-")),
        docs = path.join(root, "docs");
      fs.mkdirSync(docs);
      writeFixtures(docs);
      const cfg = path.join(root, "config.json");
      fs.writeFileSync(
        cfg,
        JSON.stringify({
          readRoots: [docs],
          writeRoots: { output: docs },
          stateDirectory: path.join(root, "state"),
        }),
      );
      const child = spawn(
        process.execPath,
        [
          new URL("../src/cli.mjs", import.meta.url).pathname,
          "serve",
          "--config",
          cfg,
        ],
        { stdio: ["pipe", "pipe", "pipe"] },
      );
      child.stderr.resume();
      let id = 0;
      const pending = new Map();
      const lines = createInterface({ input: child.stdout });
      lines.on("line", (line) => {
        const m = JSON.parse(line);
        if (m.id) {
          const p = pending.get(m.id);
          assert(p);
          pending.delete(m.id);
          clearTimeout(p.timer);
          p.resolve(m);
        }
      });
      function send(method, params) {
        return new Promise((resolve, reject) => {
          const n = ++id,
            timer = setTimeout(
              () => reject(new Error("Protocol timeout")),
              10000,
            );
          pending.set(n, { resolve, timer });
          child.stdin.write(
            JSON.stringify({ jsonrpc: "2.0", id: n, method, params }) + "\n",
          );
        });
      }
      async function call(name, args = {}) {
        const r = await send("tools/call", { name, arguments: args });
        assert(!r.error && !r.result.isError, JSON.stringify(r));
        const text = r.result.content.find((c) => c.type === "text").text;
        return JSON.parse(text);
      }
      try {
        const init = await send("initialize", {
          protocolVersion,
          capabilities: {},
          clientInfo: { name: "independent-text-client", version: "1" },
        });
        assert.equal(init.result.protocolVersion, protocolVersion);
        child.stdin.write(
          JSON.stringify({
            jsonrpc: "2.0",
            method: "notifications/initialized",
          }) + "\n",
        );
        const listed = await send("tools/list", {});
        assert(listed.result.tools.some((t) => t.name === "office_cleanup"));
        const d = (
          await call("office_open", { path: path.join(docs, "contract.docx") })
        ).document;
        const c = (
          await call("docx_replace_text", {
            documentId: d.id,
            expectedSourceRevision: d.sourceRevision,
            operationId: "wire-edit",
            paragraphId: "p_0",
            expectedText: "30 days",
            replacement: "45 days",
            expectedMatches: 1,
            changeMode: "direct",
            preservationProfile: "docx.text-replacement.v1",
          })
        ).data;
        const binding = {
          documentId: d.id,
          candidateId: c.candidateId,
          expectedCandidateRevision: c.candidateRevision,
        };
        const v = (await call("office_validate", binding)).data;
        assert(
          (
            await call("office_commit", {
              ...binding,
              expectedSourceRevision: d.sourceRevision,
              validationId: v.validationId,
              operationId: "wire-save",
              mode: "save_as",
              destination: { rootId: "output", relativePath: "out.docx" },
              overwrite: false,
            })
          ).data.published,
        );
        assert(
          (await send("tools/call", { name: "missing", arguments: {} })).error,
        );
        assert.equal((await call("office_state_info")).data.schemaVersion, 3);
      } finally {
        for (const p of pending.values()) clearTimeout(p.timer);
        child.stdin.end();
        await new Promise((resolve) => child.once("exit", resolve));
        lines.close();
        fs.rmSync(root, { recursive: true, force: true });
      }
    },
  );

test("2026-07-28 stateless discovery, result envelopes and Tasks extension", async () => {
  const root = fs.mkdtempSync(path.join(tmpdir(), "office-wire-modern-")),
    docs = path.join(root, "docs");
  fs.mkdirSync(docs);
  writeFixtures(docs);
  const cfg = path.join(root, "config.json");
  fs.writeFileSync(
    cfg,
    JSON.stringify({
      readRoots: [docs],
      writeRoots: { output: docs },
      stateDirectory: path.join(root, "state"),
    }),
  );
  const child = spawn(
    process.execPath,
    [
      new URL("../src/cli.mjs", import.meta.url).pathname,
      "serve",
      "--config",
      cfg,
    ],
    { stdio: ["pipe", "pipe", "pipe"] },
  );
  child.stderr.resume();
  let id = 0;
  const pending = new Map();
  const lines = createInterface({ input: child.stdout });
  lines.on("line", (line) => {
    const message = JSON.parse(line);
    if (message.id === undefined) return;
    const request = pending.get(message.id);
    assert(request);
    pending.delete(message.id);
    clearTimeout(request.timer);
    request.resolve(message);
  });
  const capabilities = {
    extensions: { "io.modelcontextprotocol/tasks": {} },
  };
  const meta = (clientCapabilities = capabilities) => ({
    "io.modelcontextprotocol/protocolVersion": "2026-07-28",
    "io.modelcontextprotocol/clientInfo": {
      name: "independent-modern-client",
      version: "1",
    },
    "io.modelcontextprotocol/clientCapabilities": clientCapabilities,
  });
  function send(method, params = {}, clientCapabilities = capabilities) {
    return new Promise((resolve, reject) => {
      const requestId = ++id,
        timer = setTimeout(
          () => reject(new Error(`Protocol timeout: ${method}`)),
          10000,
        );
      pending.set(requestId, { resolve, timer });
      child.stdin.write(
        JSON.stringify({
          jsonrpc: "2.0",
          id: requestId,
          method,
          params: { ...params, _meta: meta(clientCapabilities) },
        }) + "\n",
      );
    });
  }
  async function call(name, args = {}) {
    const response = await send("tools/call", { name, arguments: args });
    assert(!response.error, JSON.stringify(response));
    assert.equal(response.result.resultType, "complete");
    return response.result;
  }
  try {
    const discovery = await send("server/discover");
    assert.equal(discovery.result.resultType, "complete");
    assert.deepEqual(discovery.result.supportedVersions, ["2026-07-28"]);
    assert.deepEqual(
      discovery.result.capabilities.extensions["io.modelcontextprotocol/tasks"],
      {},
    );
    assert.equal(discovery.result.ttlMs, 60000);
    assert.equal(discovery.result.cacheScope, "private");
    assert.equal(
      discovery.result._meta["io.modelcontextprotocol/serverInfo"].name,
      "officemcp-local-preview",
    );

    const listing = await send("tools/list");
    assert.equal(listing.result.resultType, "complete");
    assert.equal(listing.result.ttlMs, 60000);
    assert.equal(listing.result.cacheScope, "private");
    assert(
      listing.result.tools.every(
        (tool) =>
          tool.inputSchema.$schema ===
            "https://json-schema.org/draft/2020-12/schema" &&
          tool.outputSchema.$schema ===
            "https://json-schema.org/draft/2020-12/schema",
      ),
    );
    assert(
      !listing.result.tools.some((tool) =>
        ["office_job_status", "office_job_cancel"].includes(tool.name),
      ),
    );
    assert(
      listing.result.tools.every(
        (tool) => tool.inputSchema.properties?.background === undefined,
      ),
    );
    const info = await call("office_server_info");
    assert.deepEqual(
      info.structuredContent.data.operations,
      listing.result.tools.map((tool) => tool.name),
    );

    const opened = await call("office_open", {
      path: path.join(docs, "contract.docx"),
    });
    const document = opened.structuredContent.document;
    const created = await send("tools/call", {
      name: "docx_structure",
      arguments: {
        documentId: document.id,
        expectedRevision: document.sourceRevision,
        offset: 0,
        limit: 1,
      },
    });
    assert.equal(created.result.resultType, "task");
    assert.equal(created.result.status, "working");
    assert.match(created.result.taskId, /^task_/);
    assert.equal(created.result.ttlMs, 7 * 86400000);

    let task;
    for (let attempt = 0; attempt < 100; attempt++) {
      const response = await send("tasks/get", {
        taskId: created.result.taskId,
      });
      assert.equal(response.result.resultType, "complete");
      task = response.result;
      if (task.status !== "working") break;
      await new Promise((resolve) => setTimeout(resolve, 20));
    }
    assert.equal(task.status, "completed");
    assert.equal(task.result.resultType, "complete");
    assert.equal(task.result.isError, false);
    assert.equal(task.result.structuredContent.data.nodes.length, 1);

    const openedPdf = await call("office_open", {
      path: path.join(docs, "invoice.pdf"),
    });
    const pdfDocument = openedPdf.structuredContent.document;
    const readRequest = await send("tools/call", {
      name: "office_read",
      arguments: {
        view: {
          documentId: pdfDocument.id,
          revision: pdfDocument.sourceRevision,
        },
        page: 1,
      },
    });
    assert.equal(
      readRequest.result.resultType,
      "task",
      JSON.stringify(readRequest.result),
    );
    let readTask;
    for (let attempt = 0; attempt < 100; attempt++) {
      const response = await send("tasks/get", {
        taskId: readRequest.result.taskId,
      });
      readTask = response.result;
      if (readTask.status !== "working") break;
      await new Promise((resolve) => setTimeout(resolve, 20));
    }
    assert.equal(readTask.status, "completed");
    assert.equal(readTask.result.content[0].type, "text");
    const readLines = readTask.result.structuredContent.data.items.flatMap(item => item.lines ?? []);
    assert(readLines.some(line => line.text?.includes("125 dollars")));

    const update = await send("tasks/update", {
      taskId: created.result.taskId,
      inputResponses: {},
    });
    assert.equal(update.result.resultType, "complete");
    const cancel = await send("tasks/cancel", {
      taskId: created.result.taskId,
    });
    assert.equal(cancel.result.resultType, "complete");

    const synchronousFallback = await send(
      "tools/call",
      {
        name: "docx_structure",
        arguments: {
          documentId: document.id,
          expectedRevision: document.sourceRevision,
          offset: 0,
          limit: 1,
          background: true,
        },
      },
      {},
    );
    assert.equal(synchronousFallback.result.resultType, "complete");
    assert.equal(synchronousFallback.result.isError, false);
    assert.equal(
      synchronousFallback.result.structuredContent.data.nodes.length,
      1,
    );
    const missingCapability = await send("tasks/get", {
      taskId: created.result.taskId,
    }, {});
    assert.equal(missingCapability.error.code, -32021);
    const malformedEnvelope = await send(
      "tasks/get",
      { taskId: created.result.taskId },
      null,
    );
    assert.equal(malformedEnvelope.error.code, -32602);
  } finally {
    for (const request of pending.values()) clearTimeout(request.timer);
    child.stdin.end();
    await new Promise((resolve) => child.once("exit", resolve));
    lines.close();
    fs.rmSync(root, { recursive: true, force: true });
  }
});
