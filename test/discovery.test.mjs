import test from "./support/test.mjs";
import assert from "node:assert/strict";
import { catalog } from "../src/contracts/catalog.mjs";
import { operationRegistry, operationDescriptor } from "../src/contracts/registry.mjs";
import { operationManifestEntry, parseOperationArguments } from "../src/contracts/operation-spec.mjs";
import { createServiceFixture } from "./service-fixture.mjs";
import {
  DOCX_AGENT_ACTIONS,
  DOCX_ACTION_REGISTRY,
  DOCX_ACTION_FAMILIES,
  docxActionDiscovery,
  parseDocxActionArguments,
} from "../src/formats/docx/docx-agent-actions.mjs";

test("every executable declaration invokes one named application command", async () => {
  for (const preset of ["all", "recommended", "docx-agent", "full"])
    for (const spec of operationRegistry(preset).values()) {
      const calls = [];
      const args = { identity: Symbol(spec.name) };
      const signal = new AbortController().signal;
      const commands = new Proxy({}, { get: (_target, command) => (...values) => {
        calls.push({ command, values });
        return args;
      } });
      assert.equal(await spec.execute({ commands }, spec.name, args, signal), args);
      assert.equal(calls.length, 1, spec.name);
      assert.match(calls[0].command, /^[a-z][a-zA-Z]+$/);
      assert.deepEqual(calls[0].values, [spec.name, args, signal]);
    }
});

test("public catalog derives every operation from its executable declaration in one pass", () => {
  const actual = catalog();
  assert.equal(actual.layers.length, 5);
  for (const layer of actual.layers) {
    const preset = layer.id === "docx" ? "docx-agent" : "all";
    const specs = [...operationRegistry(preset).values()].filter((spec) => spec.name.startsWith(layer.id + "_"));
    assert.deepEqual(layer.operations, specs.map((spec) => ({ ...operationManifestEntry(spec), ...spec.discovery })));
    assert.equal(new Set(layer.operations.map(({ name }) => name)).size, layer.operations.length);
  }
  assert.equal(actual.layers.flatMap(({ operations }) => operations).length, 58);
  assert.equal(actual.layers[0].operations.find(({ name }) => name === "docx_image_get").resultType, "JSON + artifact + native MCP image");
  assert(!actual.layers.some(({ operations }) => operations.some(({ name }) => name.endsWith("_render"))));
  assert.equal(actual.layers[2].operations.find(({ name }) => name === "pptx_set_notes").status, "Requires a configured licensed Aspose runtime");
});

test("legacy render tools remain absent; page selectors require a revision-bound layout and literal text", () => {
  for (const preset of ["all", "recommended", "docx-agent", "full"])
    for (const name of ["office_render", "docx_render", "pdf_render", "pptx_render"])
      assert.equal(operationDescriptor(preset, name), null);
  const view = { documentId: "doc", revision: "sha256:" + "a".repeat(64) };
  for (const name of ["docx_validate", "docx_compare"]) {
    const spec = operationDescriptor("docx-agent", name);
    assert.equal(parseOperationArguments(spec, { view }).success, true);
    assert.equal(parseOperationArguments(spec, { view, visualEvidence: { pages: [1] } }).success, false);
    assert.equal(spec.output.properties.data.properties.visualEvidence, undefined);
  }
  assert.equal(parseOperationArguments(operationDescriptor("docx-agent", "docx_query"), {
    view, select: { type: "page", layoutId: "layout", page: 1 },
  }).success, false);
  assert.equal(parseOperationArguments(operationDescriptor("docx-agent", "docx_query"), { view, select: { type: "page", layoutId: "layout", number: 1, pattern: "A unique clause." } }).success, true);
  assert(operationDescriptor("docx-agent", "docx_layout"));
  assert(operationDescriptor("docx-agent", "docx_export"));
  assert(operationDescriptor("docx-agent", "docx_image_get"));
  assert.equal(parseOperationArguments(operationDescriptor("docx-agent", "docx_apply"), {
    view, mode: "preview", action: "rewrite_block", arguments: { text: "Dry-run content" },
  }).success, true);
});

test("family discovery covers each action once and keeps descriptor selection bounded", () => {
  const all = docxActionDiscovery();
  assert.deepEqual(all.availableActions, DOCX_AGENT_ACTIONS);
  assert.deepEqual(all.actionDescriptors, {});
  const grouped = DOCX_ACTION_FAMILIES.flatMap(({ actions }) => actions);
  assert.deepEqual([...grouped].sort(), [...DOCX_AGENT_ACTIONS].sort());
  assert.equal(new Set(grouped).size, 79);
  for (const family of DOCX_ACTION_FAMILIES) {
    const found = docxActionDiscovery({ actionFamily: family.id });
    assert.equal(found.selectedActionFamily, family.id);
    assert.deepEqual(found.availableActions, family.actions);
    assert.deepEqual(Object.keys(found.actionDescriptors), family.actions);
    assert(family.actions.length <= 20);
    for (const descriptor of Object.values(found.actionDescriptors)) {
      assert.equal(descriptor.family, family.id);
      assert.equal(descriptor.argumentSchema.$schema, "https://json-schema.org/draft/2020-12/schema");
    }
  }
  const images = docxActionDiscovery({ actionFamily: "images", actions: ["resize_image"] });
  assert.deepEqual(Object.keys(images.actionDescriptors), ["resize_image"]);
  assert.throws(() => docxActionDiscovery({ actionFamily: "invented" }), /Unknown/);
  assert.throws(() => docxActionDiscovery({ actions: Array(21).fill("resize_image") }), /At most 20/);
});

test("inspection adds an optional family selector without changing existing defaults", () => {
  const spec = operationDescriptor("docx-agent", "docx_inspect");
  const view = { documentId: "doc", revision: "sha256:" + "a".repeat(64) };
  assert.deepEqual(parseOperationArguments(spec, { view }).data, { view });
  assert.equal(parseOperationArguments(spec, { view, actionFamily: "tables" }).success, true);
  assert.equal(parseOperationArguments(spec, { view, actionFamily: "invented" }).success, false);
  assert(spec.output.properties.data.properties.actionFamilies);
  assert(spec.output.properties.data.properties.selectedActionFamily);
});

test("public DOCX inspection carries selected family descriptors through the SDK worker", async () => {
  const fixture = createServiceFixture({ toolset: "docx-agent" });
  try {
    const created = await fixture.call("docx_create", {
      operationId: "family-document",
      document: { format: "docx", blocks: [{ kind: "paragraph", text: "Family discovery" }] },
    });
    const selected = await fixture.call("docx_inspect", { view: created.view, actionFamily: "text" });
    assert.equal(selected.data.selectedActionFamily, "text");
    assert.equal(selected.data.availableActions.length, 6);
    assert.equal(selected.data.actionDescriptors.delete_text.name, "delete_text");
    assert.equal(selected.data.actionDescriptors.insert_heading.name, "insert_heading");
    assert.equal(selected.data.actionDescriptors.delete_text.aliasOf, undefined);
    assert.equal(selected.data.actionDescriptors.insert_heading.aliasOf, undefined);
    assert.equal(selected.data.actionFamilies.length, 12);
    assert.deepEqual(Object.keys(selected.data.actionDescriptors), selected.data.availableActions);
  } finally { fixture.close(); }
});

test("similar action names retain their distinct occurrence and placement semantics", () => {
  assert.deepEqual(parseDocxActionArguments("delete_text", {
    finds: ["obsolete", "temporary"], caseSensitive: true, changeMode: "tracked",
  }), {
    finds: ["obsolete", "temporary"], caseSensitive: true, changeMode: "tracked",
  });
  assert.deepEqual(parseDocxActionArguments("insert_heading", {
    text: "Terms", level: 2, placement: { at: "before" }, changeMode: "tracked",
  }), {
    text: "Terms", level: 2, placement: { at: "before" }, changeMode: "tracked",
  });
  for (const action of ["delete_text", "insert_heading"]) {
    assert.equal(DOCX_ACTION_REGISTRY[action].aliasOf, undefined);
  }
  assert.throws(() => parseDocxActionArguments("insert_heading", { text: "Terms", level: 7 }), /Invalid arguments/);
  assert.throws(() => parseDocxActionArguments("delete_text", { finds: [""] }), /Invalid arguments/);
});
