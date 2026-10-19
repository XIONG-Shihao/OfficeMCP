import assert from "node:assert/strict";
import test from "./support/test.mjs";
import {
  assertSupportedRuntime,
  isSupportedBunVersion,
  isSupportedNodeVersion,
  javascriptRuntime,
} from "../src/infrastructure/node-version.mjs";

test("OfficeMCP accepts Node 24.12 through 24.x or Bun 1.3.13 through 1.x", () => {
  for (const supported of ["24.12.0", "24.20.0", "24.99.1"]) assert.equal(isSupportedNodeVersion(supported), true);
  for (const rejected of ["22.23.2", "24.11.99", "25.0.0", "invalid"]) assert.equal(isSupportedNodeVersion(rejected), false);
  for (const supported of ["1.3.13", "1.4.0", "1.99.1"]) assert.equal(isSupportedBunVersion(supported), true);
  for (const rejected of ["1.3.12", "1.2.99", "2.0.0", "", undefined]) assert.equal(isSupportedBunVersion(rejected), false);
  assert.doesNotThrow(() => assertSupportedRuntime({ node: "24.12.0" }));
  assert.doesNotThrow(() => assertSupportedRuntime({ bun: "1.3.13", node: "24.3.0" }), "Under Bun the Node compatibility version is not the gate");
  assert.throws(() => assertSupportedRuntime({ bun: "1.2.0", node: "24.12.0" }), { code: "UNSUPPORTED_RUNTIME" });
  assert.throws(
    () => assertSupportedRuntime({ node: "22.0.0" }),
    (error) => error.code === "UNSUPPORTED_RUNTIME" && /Node >=24\.12\.0 <25 or Bun >=1\.3\.13 <2/.test(error.message),
  );
  assert.deepEqual(javascriptRuntime({ bun: "1.3.13", node: "24.3.0" }), { kind: "bun", version: "1.3.13" });
  assert.deepEqual(javascriptRuntime({ node: "24.12.0" }), { kind: "node", version: "24.12.0" });
});
