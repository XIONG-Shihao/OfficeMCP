import test from "./support/test.mjs";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";

const root = new URL("../src/", import.meta.url);
const rootPath = root.pathname;

function files(directory) {
  return fs.readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const target = path.join(directory, entry.name);
    return entry.isDirectory() ? files(target) : entry.isFile() ? [target] : [];
  });
}

test("source root contains only stable public entry points", () => {
  const actual = fs
    .readdirSync(root, { withFileTypes: true })
    .filter((entry) => entry.isFile())
    .map((entry) => entry.name)
    .sort();
  assert.deepEqual(actual, [
    ".npmignore",
    "authoring.mjs",
    "cli.mjs",
    "sdk.d.mts",
    "sdk.mjs",
    "storage.d.mts",
    "storage.mjs",
  ]);
  const manifest=JSON.parse(fs.readFileSync(new URL('../package.json',import.meta.url),'utf8'));
  assert.equal(manifest.exports['./authoring'],'./src/authoring.mjs','Host guide schemas must use a declared public export');
});

test("foundation and format modules cannot depend on orchestration or transport", () => {
  const guarded = ["infrastructure", "formats", "engines", "workers"];
  for (const layer of guarded) {
    for (const file of files(path.join(rootPath, layer)).filter((name) => name.endsWith(".mjs"))) {
      const source = fs.readFileSync(file, "utf8");
      assert(
        !/["'](?:\.\.\/)+application\//.test(source),
        `${path.relative(rootPath, file)} imports application orchestration`,
      );
      assert(
        !/["'](?:\.\.\/)+protocol\//.test(source),
        `${path.relative(rootPath, file)} imports MCP transport`,
      );
      assert(
        !/["'](?:\.\.\/)+contracts\/registry\.mjs/.test(source),
        `${path.relative(rootPath, file)} imports the global tool registry`,
      );
    }
  }
});

test("format lifecycle facades cannot write coordinator records or own transactions", () => {
  for (const relative of ["formats/common/document-workflow.mjs", "formats/docx/docx-service.mjs", "formats/docx/docx-agent-service.mjs"]) {
    const source = fs.readFileSync(path.join(rootPath, relative), "utf8");
    assert(!/service\.(?:put|delete|rows)\s*\(|service\.db\b/.test(source),
      `${relative} must delegate persisted handles and transactions to the document runtime`);
  }
});
