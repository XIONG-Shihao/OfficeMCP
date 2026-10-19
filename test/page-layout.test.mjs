import test from "./support/test.mjs";
import assert from "node:assert/strict";
import { createServiceFixture } from "./service-fixture.mjs";
import { manifest } from "../src/contracts/registry.mjs";

test("combined surface preserves DOCX agent contracts and four-format tools", async () => {
  const s = createServiceFixture({ toolset: "full" });
  try {
    for (const name of ["docx_apply", "docx_export", "xlsx_read", "pptx_read", "pdf_read", "office_server_info"])
      assert(manifest("full").some((t) => t.name === name), name);
    const created = await s.call("docx_create", { operationId: "full-create", document: { format: "docx", blocks: [{ kind: "paragraph", text: "Combined tool surface document." }] } });
    assert.equal(created.contractVersion, "docx/6");
    assert.equal((await s.call("docx_read", { view: created.view })).contractVersion, "docx/6");
    assert.notEqual((await s.call("office_server_info")).contractVersion, "docx/6");
    const failure = await s.service.call("docx_query", {});
    assert.equal(failure.payload.contractVersion, "docx/6");
  } finally { s.close(); }
});
test("removed DOCX visual tools and selectors fail before semantic execution", async () => {
  const s = createServiceFixture({ toolset: "full" });
  try {
    const view = (await s.call("docx_create", {
      operationId: "removed-visual-create",
      document: { format: "docx", blocks: [{ kind: "paragraph", text: "Semantic document." }] },
    })).view;
    assert(!manifest("full").some((tool) => tool.name === "docx_render"));
    const render = await s.service.call("docx_render", { view, page: 1 });
    assert.equal(render.payload.error.code, "UNKNOWN_TOOL");
    const pageSelect = { type: "page", page: 1, layoutId: "removed-layout" };
    for (const [name, args] of [
      ["docx_query", { view, select: pageSelect }],
      ["docx_apply", { view, operationId: "removed-page-apply", action: "rewrite_block", select: pageSelect, arguments: { text: "Must not replace content" } }],
      ["docx_validate", { view, visualEvidence: { pages: [1] } }],
      ["docx_compare", { view, visualEvidence: { mode: "changedPages" } }],
    ]) {
      const rejected = await s.service.call(name, args);
      assert.equal(rejected.payload.error.code, "INVALID_ARGUMENT", name);
    }
    const read = await s.call("docx_read", { view, representation: "text" });
    assert(read.data.items.some((item) => item.text === "Semantic document."));
    assert.equal((await s.call("docx_validate", { view })).data.passed, true);
    const compared = await s.call("docx_compare", { view });
    assert.equal(compared.data.visualEvidence, undefined);
  } finally { s.close(); }
});
test("paragraph insertion accepts table-cell paragraph targets without changing adjacent cells", async () => {
  const s = createServiceFixture({ toolset: "full" });
  try {
    let view = (await s.call("docx_create", { operationId: "table-anchor-create", document: { format: "docx", blocks: [
      {kind:"table",rows:[["Table cell anchor paragraph.","Adjacent cell must survive."]]}
    ] } })).view;
    const queried = await s.call("docx_query", {view,select:{type:"text",pattern:"Table cell anchor paragraph."},require:"exactlyOne"});
    view = (await s.call("docx_apply", {view,operationId:"table-paragraph-insert",action:"insert_paragraphs",selectorTarget:queried.data.items[0].target,arguments:{text:"Inserted paragraph in the same cell.",placement:{at:"after"}}})).view;
    const tables = await s.call("docx_read", {view,representation:"domain",domain:"tables"});
    const text=JSON.stringify(tables.data);
    assert(text.includes("Inserted paragraph in the same cell."));
    assert(text.includes("Adjacent cell must survive."));
  } finally {s.close();}
});
test("completed quota reservations are reclaimed without discarding in-flight budgets", () => {
  const s=createServiceFixture({limits:{maxStateBytes:4*1024**2,maxFileBytes:1024**2}});
  try {
    const fill=()=>Object.assign(s.service.quotaSnapshot,{checkedAt:Date.now(),bytes:0,reserved:3*1024**2});
    fill();
    assert.doesNotThrow(()=>s.service.checkQuota(2*1024**2));
    const worker={};
    s.service.activeWorkers.add(worker);
    fill();
    assert.throws(()=>s.service.checkQuota(2*1024**2),/quota reached/);
    s.service.activeWorkers.delete(worker);
  }finally{s.close();}
});
test("combined task results keep their operation-specific contract", async () => {
  const s=createServiceFixture({toolset:"full"});
  try {
    const task=s.service.startTask("docx_create",{operationId:"full-task",document:{format:"docx",blocks:[{kind:"paragraph",text:"Combined task contract check."}]}});
    const id=typeof task==="string"?task:task.taskId;
    const until=Date.now()+30000;
    while(s.service.get("task",id).status==="running" && Date.now()<until) await new Promise(resolve=>setTimeout(resolve,20));
    const record=s.service.get("task",id);
    assert.equal(record.status,"completed",JSON.stringify(record.result));
    assert.equal(record.resultContract,"docx/6");
    assert.equal(record.result.contractVersion,"docx/6");
  } finally{s.close();}
});
