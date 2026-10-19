import test from "./support/test.mjs";
import assert from "node:assert/strict";
import * as fs from "node:fs";
import path from "node:path";
import { tmpdir } from "node:os";
import { spawnSync } from "node:child_process";
import { DatabaseSync } from "./support/raw-sqlite.mjs";
import { OfficeService } from "../src/application/service.mjs";
import { configFrom, digest } from "../src/infrastructure/files.mjs";
import { writeFixtures } from "./fixtures.mjs";
function fixture() {
  const root = fs.mkdtempSync(path.join(tmpdir(), "office-life-")),
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
  return { root, docs, cfg, config: configFrom(cfg) };
}
async function call(s, n, a) {
  const r = await s.call(n, a);
  assert(!r.isError, JSON.stringify(r.payload));
  return r.payload;
}
async function prepare(f, s) {
  const d = (
    await call(s, "office_open", { path: path.join(f.docs, "contract.docx") })
  ).document;
  const edit = {
    documentId: d.id,
    expectedSourceRevision: d.sourceRevision,
    operationId: "edit",
    paragraphId: "p_0",
    expectedText: "30 days",
    replacement: "45 days",
    expectedMatches: 1,
    changeMode: "direct",
    preservationProfile: "docx.text-replacement.v1",
  };
  return { d, edit };
}
test("private-state connections refuse attached databases and enforce their bounds", () => {
  const f = fixture();
  const s = new OfficeService(f.config);
  try {
    if (process.versions.bun) {
      // Bun exposes no SQLite limits; the wrapper admits only known statements.
      assert.throws(() => s.db.exec("ATTACH DATABASE ':memory:' AS unexpected"), { code: "UNSUPPORTED_STATE" });
      assert.throws(() => s.db.prepare("SELECT load_extension(?)"), { code: "UNSUPPORTED_STATE" });
      assert.throws(() => s.db.exec("PRAGMA writable_schema=ON"), { code: "UNSUPPORTED_STATE" });
      return;
    }
    assert.equal(s.db.limits.attach, 0);
    assert.equal(s.db.limits.sqlLength, 64 * 1024);
    assert.equal(s.db.limits.variableNumber, 32);
    assert.throws(
      () => s.db.exec("ATTACH DATABASE ':memory:' AS unexpected"),
      /attached|too many/i,
    );
  } finally {
    s.close();
    fs.rmSync(f.root, { recursive: true, force: true });
  }
});
function kill(f, phase, name, args) {
  const req = path.join(f.root, "request.json");
  fs.writeFileSync(req, JSON.stringify({ name, args }));
  const r = spawnSync(
    process.execPath,
    [new URL("./crash-child.mjs", import.meta.url).pathname, f.cfg, phase, req],
    { timeout: 15000, stdio: "pipe" },
  );
  assert.equal(r.signal, "SIGKILL", r.stderr?.toString());
}
for (const phase of [
  "candidate-before-commit",
  "candidate-after-commit",
  "publication-before-link",
  "publication-after-link",
  "publication-after-receipt",
])
  test("M6 actual SIGKILL recovery: " + phase, async () => {
    const f = fixture();
    let s = new OfficeService(f.config);
    try {
      const { d, edit } = await prepare(f, s);
      let name = "docx_replace_text",
        args = edit;
      if (phase.startsWith("publication")) {
        const c = (await call(s, name, args)).data;
        const b = {
          documentId: d.id,
          candidateId: c.candidateId,
          expectedCandidateRevision: c.candidateRevision,
        };
        const v = (await call(s, "office_validate", b)).data;
        name = "office_commit";
        args = {
          ...b,
          expectedSourceRevision: d.sourceRevision,
          validationId: v.validationId,
          operationId: "commit",
          mode: "save_as",
          destination: { rootId: "output", relativePath: "published.docx" },
          overwrite: false,
        };
      }
      s.close();
      kill(f, phase, name, args);
      s = new OfficeService(f.config);
      assert.equal(
        digest(fs.readFileSync(path.join(f.docs, "contract.docx"))),
        d.sourceRevision,
      );
      const result = await s.call(name, args);
      if (phase === "candidate-before-commit") {
        assert.equal(result.payload.error.code, "PARTIAL_MUTATION");
        assert.equal(s.rows("candidate").length, 0);
      } else if (phase === "publication-before-link") {
        assert.equal(result.payload.error.code, "PUBLICATION_UNCERTAIN");
        assert(!fs.existsSync(path.join(f.docs, "published.docx")));
      } else {
        assert(!result.isError, JSON.stringify(result));
        if (phase.startsWith("publication"))
          assert.equal(
            digest(fs.readFileSync(path.join(f.docs, "published.docx"))),
            result.payload.data.outputRevision,
          );
        else assert.equal(s.rows("candidate").length, 1);
      }
    } finally {
      s.close();
      fs.rmSync(f.root, { recursive: true, force: true });
    }
  });
test("M6 migration rollback survives SIGKILL; unknown schema/checksum cannot claim ownership", () => {
  const f = fixture();
  let s = new OfficeService(f.config);
  s.close();
  const file = path.join(f.config.stateDirectory, "state.sqlite");
  let db = new DatabaseSync(file);
  db.prepare(
    "UPDATE records SET value=? WHERE kind='meta' AND id='schema'",
  ).run('{"version":1}');
  db.prepare(
    "DELETE FROM records WHERE kind='meta' AND id='migration-2'",
  ).run();
  db.close();
  const r = spawnSync(
    process.execPath,
    [
      new URL("./crash-child.mjs", import.meta.url).pathname,
      f.cfg,
      "migration-before-commit",
    ],
    { timeout: 15000, stdio: "pipe" },
  );
  assert.equal(r.signal, "SIGKILL");
  db = new DatabaseSync(file);
  assert.equal(
    JSON.parse(
      db
        .prepare("SELECT value FROM records WHERE kind='meta' AND id='schema'")
        .get().value,
    ).version,
    1,
  );
  db.close();
  s = new OfficeService(f.config);
  assert.equal(s.stateInfo().schemaVersion, 3);
  s.put("meta", "schema", { version: 99 });
  s.close();
  assert.throws(() => new OfficeService(f.config), {
    code: "UNSUPPORTED_STATE",
  });
  db = new DatabaseSync(file);
  assert.equal(
    db.prepare("SELECT count(*) AS n FROM records WHERE kind='owner'").get().n,
    0,
  );
  db.prepare(
    "UPDATE records SET value=? WHERE kind='meta' AND id='schema'",
  ).run('{"version":2}');
  db.prepare(
    "UPDATE records SET value=? WHERE kind='meta' AND id='migration-2'",
  ).run('{"checksum":"tampered"}');
  db.close();
  assert.throws(() => new OfficeService(f.config), {
    code: "UNSUPPORTED_STATE",
  });
  fs.rmSync(f.root, { recursive: true, force: true });
});
test("M6 lease expiry, cleanup, orphan removal and permanent operation tombstones preserve sources and outputs", async () => {
  const f = fixture();
  let time = Date.now();
  const s = new OfficeService(
    { ...f.config, artifactRetentionMs: 100, operationRetentionMs: 200 },
    { now: () => time },
  );
  try {
    const { d, edit } = await prepare(f, s);
    const c = (await call(s, "docx_replace_text", edit)).data;
    const b = {
      documentId: d.id,
      candidateId: c.candidateId,
      expectedCandidateRevision: c.candidateRevision,
    };
    const v = (await call(s, "office_validate", b)).data;
    const out = (
      await call(s, "office_commit", {
        ...b,
        expectedSourceRevision: d.sourceRevision,
        validationId: v.validationId,
        operationId: "save",
        mode: "save_as",
        destination: { rootId: "output", relativePath: "published.docx" },
        overwrite: false,
      })
    ).data;
    fs.writeFileSync(s.internal("assets", "orphan.bin"), "orphan");
    time += 101;
    assert.equal(
      (
        await s.call("docx_read", {
          documentId: d.id,
          expectedRevision: d.sourceRevision,
        })
      ).payload.error.code,
      "HANDLE_EXPIRED",
    );
    const cleaned = (await call(s, "office_cleanup", {})).data;
    assert.equal(cleaned.expiredDocuments, 1);
    assert(cleaned.removedFiles >= 3);
    assert.equal(digest(fs.readFileSync(out.path)), out.outputRevision);
    assert.equal(
      digest(fs.readFileSync(path.join(f.docs, "contract.docx"))),
      d.sourceRevision,
    );
    time += 100;
    await call(s, "office_cleanup", {});
    assert.equal(
      (await s.call("docx_replace_text", edit)).payload.error.code,
      "OPERATION_EXPIRED",
    );
    assert.equal(
      (await s.call("office_operation_status", { operationId: "edit" })).payload
        .error.code,
      "OPERATION_EXPIRED",
    );
    assert.equal(s.get("operation", "edit").status, "expired");
  } finally {
    s.close();
    fs.rmSync(f.root, { recursive: true, force: true });
  }
});

test("M6 snapshot reads survive a removed original; publication refuses and cleanup waits for workers", async () => {
  const f = fixture();
  const s = new OfficeService(f.config);
  try {
    const { d, edit } = await prepare(f, s);
    const c = (await call(s, "docx_replace_text", edit)).data;
    const b = {
      documentId: d.id,
      candidateId: c.candidateId,
      expectedCandidateRevision: c.candidateRevision,
    };
    const v = (await call(s, "office_validate", b)).data;
    fs.unlinkSync(path.join(f.docs, "contract.docx"));
    assert(
      (
        await call(s, "docx_read", {
          documentId: d.id,
          expectedRevision: d.sourceRevision,
        })
      ).data.items.some((p) => p.text.includes("30 days")),
    );
    assert(
      (
        await s.call("office_commit", {
          ...b,
          expectedSourceRevision: d.sourceRevision,
          validationId: v.validationId,
          operationId: "missing-source",
          mode: "save_as",
          destination: { rootId: "output", relativePath: "refused.docx" },
          overwrite: false,
        })
      ).isError,
    );
    assert(!fs.existsSync(path.join(f.docs, "refused.docx")));
    const pending = s.call("docx_read", {
      documentId: d.id,
      expectedRevision: d.sourceRevision,
    });
    assert.equal(
      (await s.call("office_cleanup", {})).payload.error.code,
      "BUSY",
    );
    const readResult = await pending;
    assert(!readResult.isError, JSON.stringify(readResult.payload));
  } finally {
    s.close();
    fs.rmSync(f.root, { recursive: true, force: true });
  }
});
