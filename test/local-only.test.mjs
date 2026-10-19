import test from "./support/test.mjs";
import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { readFileSync, existsSync } from "node:fs";
test("Local-only release rejects HTTP CLI mode and removes public network entry points", () => {
  const p = JSON.parse(
    readFileSync(new URL("../package.json", import.meta.url)),
  );
  assert(!p.exports["./http"]);
  assert(!existsSync(new URL("../src/http.mjs", import.meta.url)));
  const result = spawnSync(
    process.execPath,
    [
      new URL("../src/cli.mjs", import.meta.url).pathname,
      "http",
      "--config",
      "/not-read.json",
    ],
    { encoding: "utf8", timeout: 5000 },
  );
  assert.equal(result.status, 1);
  assert.match(result.stderr, /Usage:.*serve/);
  assert.equal(result.stdout, "");
});
