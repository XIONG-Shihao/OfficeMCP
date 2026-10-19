import test from "./support/test.mjs";
import assert from "node:assert/strict";
import { executeDocxHostTask } from "../src/formats/docx/docx-host-process.mjs";

test("unknown native commands fail before opening a document or falling through to a read", async () => {
  const handlers = new Proxy({}, { get: () => () => { throw new Error("No handler may run"); } });
  for (const kind of ["docx_native_typo", "docx_native_agent_edit", "toString", "constructor"])
    await assert.rejects(executeDocxHostTask(handlers, { kind, args: {} }), { code: "UNKNOWN_TOOL" });
});

test("native creation resets ownership and preview keeps its discard flag", async () => {
  const events = [], task = { kind: "docx_native_create", args: { marker: 1 } };
  const handlers = {
    dispose: async () => events.push("dispose"), load: async value => events.push(["load", value]),
    create: value => { events.push(["create", value]); return "created"; },
    agentApply: (value, preview) => { events.push(["apply", value, preview]); return "previewed"; },
  };
  assert.equal(await executeDocxHostTask(handlers, task), "created");
  assert.deepEqual(events, ["dispose", ["create", task]]);
  events.length = 0; task.kind = "docx_native_agent_preview";
  assert.equal(await executeDocxHostTask(handlers, task), "previewed");
  assert.deepEqual(events, [["load", task], ["apply", task, true]]);
});
