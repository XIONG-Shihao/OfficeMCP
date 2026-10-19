import test from "./support/test.mjs";
import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

test("formula policy distinguishes structured references from external workbooks", (t) => {
  const worker = fileURLToPath(new URL("../src/engines/aspose-worker.py", import.meta.url));
  const source = [
    "import importlib.util,json,sys",
    "s=importlib.util.spec_from_file_location('worker',sys.argv[1])",
    "m=importlib.util.module_from_spec(s);s.loader.exec_module(m)",
    "print(json.dumps([m.formula_uses_external_resources(x) for x in sys.argv[2:]]))",
  ].join(";");
  const result = spawnSync("python3", ["-c", source, worker,
    "=SUM(Table1[Amount])", "=Table1[file.xlsx]", '="A[B]"', "=SUM(Sheet1!Table1[Amount])",
    "=SUM([1]Sheet!A1)", "='[Book.xlsx]Sheet 1'!A1", "=WEBSERVICE(\"https://example.invalid\")",
  ], { encoding: "utf8" });
  if (result.error?.code === "ENOENT") return t.skip("python3 unavailable");
  assert.equal(result.status, 0, result.stderr);
  assert.deepEqual(JSON.parse(result.stdout), [false, false, false, false, true, true, true]);
});
