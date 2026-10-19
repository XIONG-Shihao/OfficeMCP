import { operationBinding } from "../src/application/tool-dispatch.mjs";
import test from "./support/test.mjs";
import assert from "node:assert/strict";
import {
  operationRegistry,
  operationDescriptor,
} from "../src/contracts/registry.mjs";
import { formatAdapterIds } from "../src/workers/format-adapters.mjs";
import {
  DOCX_ACTION_REGISTRY,
  matchesDocxActionTarget,
} from "../src/formats/docx/docx-agent-actions.mjs";

test("executable operation specs own transport and runtime contracts", () => {
  for (const preset of ["all", "recommended", "docx-agent", "full"])
    for (const [name, operation] of operationRegistry(preset)) {
      assert.equal(operation.name, name);
      assert(operation.input);
      assert.equal(operation.output.$schema, "https://json-schema.org/draft/2020-12/schema");
      assert(operation.policy.handler);
      const binding = operationBinding(preset, name);
      assert.equal(binding.operation, operation);
      assert.equal(typeof binding.execute, "function");
      assert(["synchronous", "taskEligible"].includes(operation.policy.execution));
      assert.equal(
        operationDescriptor(preset, name).policy.handler,
        operation.policy.handler,
      );
    }
});

test("worker adapters have unique explicit identities", () => {
  assert.equal(new Set(formatAdapterIds).size, formatAdapterIds.length);
  assert(formatAdapterIds.includes("opc.read"));
  assert(formatAdapterIds.includes("pdf.pdfjs"));
  assert(formatAdapterIds.includes("aspose"));
});

test("DOCX action descriptors drive target matching", () => {
  assert.equal(Object.keys(DOCX_ACTION_REGISTRY).length, 79);
  assert(matchesDocxActionTarget("image", { kind: "image" }));
  assert(matchesDocxActionTarget("block", {
    kind: "agent_selector",
    address: { nodeType: "paragraph" },
  }));
  assert(!matchesDocxActionTarget("tableCell", {
    kind: "agent_selector",
    address: { nodeType: "paragraph" },
  }));
});
