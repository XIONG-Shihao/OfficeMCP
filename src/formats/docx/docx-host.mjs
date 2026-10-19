import { readDocxDomain } from "./docx-domain-items.mjs";
import fs from "node:fs";
import { openFilePackage } from "../opc/file-package.mjs";
import { validateXml } from "../opc/ooxml.mjs";
import { createAgentToolkit } from "@superdoc/sdk";
import {
  clientFor,
  paragraphs,
  textOf,
  list,
  headers,
  createSdk,
  ENGINE,
  tables,
} from "./docx-sdk.mjs";
import { mutate } from "./docx-sdk-actions.mjs";
import { requireThat } from "../../infrastructure/errors.mjs";
import { storyLabel, storyOf } from "./docx-story.mjs";
import { inspectDocxPackage } from "./docx-package-inventory.mjs";
import { inspectAgentDocument } from "./docx-agent-inspect.mjs";
import {
  directTextReplacement,
  queryAgentDocument,
  resolveInlineLocator,
} from "./docx-agent-query.mjs";
import {
  canonicalItem,
  paged,
  pagedWindow,
  structuredNode,
} from "./docx-canonical.mjs";
import { installDocxHostProcess } from "./docx-host-process.mjs";
import { validateDocxCandidate } from "./docx-validation.mjs";
import { executeDocxAction } from "./docx-action-engine.mjs";
import { executeDocxMutationPlan } from "./docx-mutation-plan.mjs";
import { semanticWorkflow } from "./docx-semantic-workflows.mjs";
import { captureBodyChanges, bodyReadback } from "./docx-change-readback.mjs";
import { DocxReadCache } from "../../infrastructure/docx-read-cache.mjs";
import { digest } from "../../infrastructure/files.mjs";
import {tracePhase} from "../../infrastructure/docx-phase-trace.mjs";
import { sdkCellEdit } from './docx-cell-sdk.mjs';
let client,
  doc,
  agentToolkit,
  currentPath,
  currentMaxExpandedBytes,
  packageInventory,
  author = "OfficeMCP",
  loads = 0,
  currentRevision,
  sdkRevision,
  readGeneration = 0;
const readCache = new DocxReadCache();
const readSnapshots = new DocxReadCache();
function invalidateReads() {readGeneration++;readCache.clear();readSnapshots.clear();}
async function cachedRead(kind,args,load) {
  const identity={kind,args,path:currentPath,revision:currentRevision,sdkRevision,model:loads,generation:readGeneration};
  const cached=readCache.get('sdk-read',identity);
  if(cached!==undefined)return cached;
  const result=await load();readCache.set('sdk-read',identity,result);return result;
}
async function dispose() {
  if (doc) await doc.close({ discard: true }).catch(() => {});
  doc = null;
  currentPath = null;
  currentMaxExpandedBytes = null;
  packageInventory = null;
  currentRevision = null;
  sdkRevision = null;
  invalidateReads();
}
async function load(task) {
  client ??= clientFor(task);
  const authors = [
    ...new Set(
      (task.args.steps ?? [])
        .filter((s) => s.changeMode === "tracked")
        .map((s) => s.author),
    ),
  ];
  requireThat(
    authors.length <= 1,
    "INVALID_ARGUMENT",
    "Use one tracked-change author per batch.",
  );
  const requested = task.args.author ?? authors[0] ?? author;
  if(doc&&currentPath===task.input&&requested===author) {
    try{const info=await doc.info({}),revision=String(info.revision??info.document?.revision);if(sdkRevision!==revision)invalidateReads();sdkRevision=revision;}
    catch(error){throw Error(`SDK binding unavailable; this worker must be discarded before retrying: ${error.message}`);}
  }
  if (!doc || currentPath !== task.input || requested !== author) {
    await dispose();
    const archive=openFilePackage(task.input,task.maxExpandedBytes,task.maxFileBytes);
    try {
      currentRevision=archive.revision;
      // The engine accepts a body with a DTD or entity declarations and opens it
      // as an empty document. Refuse it here: such a body is never a Word file
      // this server may read from or write back.
      const body=archive.parts['word/document.xml'];
      requireThat(body,'INVALID_DOCUMENT','DOCX main document part is missing.');
      validateXml(body);
    } finally { archive.close(); }
    doc = await tracePhase('sdkOpen',()=>client.open({
      doc: task.input,
      runtime: "v2",
      userName: requested,
    }));
    requireThat(
      doc.openResult.runtime === "v2",
      "UNSUPPORTED_FEATURE",
      "SuperDoc v2 is required.",
    );
    // Read the SDK body only when the requested
    // operation actually needs the model; query/info paths use their own APIs.
    currentPath = task.input;
    currentMaxExpandedBytes = task.maxExpandedBytes;
    author = requested;
    loads++;
    const openedRevision=doc.openResult.document?.revision;
    if(Number.isSafeInteger(openedRevision)&&openedRevision>=0)sdkRevision=String(openedRevision);
    else {const info=await tracePhase('sdkBindingInfo',()=>doc.info({}));sdkRevision=String(info.revision??info.document?.revision);}
  }
}

function currentPackageInventory() {
  packageInventory ??= inspectDocxPackage(
    currentPath,
    currentMaxExpandedBytes,
  );
  return packageInventory;
}

async function toolkit() {
  agentToolkit ??= await createAgentToolkit({
    provider: "generic",
    preset: "core",
  });
  return agentToolkit;
}

async function agentInspect(a, task) {
  const result = await inspectAgentDocument(doc, await toolkit(), a, task);
  packageInventory = {
    storyInventory: result.storyInventory,
    objectInventory: result.objectInventory,
    templateReport: result.templateReport,
    textboxInventory: result.textboxInventory,
  };
  return result;
}

async function agentRead(a) {
  if (a.representation === "domain" && a.domain === "objects") {
    const inventory = packageInventory ??= inspectDocxPackage(currentPath, currentMaxExpandedBytes);
    return paged(inventory.complexObjects, a, "objects", { engine: "read-only OOXML complex object inventory", notInterpreted: ["embedded application payloads", "chart workbook formulas", "SmartArt layout", "equation evaluation"] });
  }
  if (a.representation === "complete") {
    let snapshot = a.readSnapshotId
      ? readSnapshots.get('complete',{id:a.readSnapshotId,revision:currentRevision,generation:readGeneration})
      : null;
    if (!snapshot) {
      const scopes = [];
      const body = await agentRead({
        ...a,
        representation: "structured",
        story: { type: "body" },
        offset: 0,
        limit: 10000,
      });
      scopes.push({ scope: "body", items: body.items });
      for (const domain of [
        "headers_footers",
        "footnotes",
        "fields",
        "hyperlinks",
        "bookmarks",
        "content_controls",
        "images",
        "tracked_changes",
        "comments",
        "sections",
        "toc",
        "captions",
        "cross_references",
      ]) {
        const result = await agentRead({
          ...a,
          representation: "domain",
          story: undefined,
          domain,
          offset: 0,
          limit: 10000,
        });
        scopes.push({ scope: domain, items: result.items });
      }
      const inventory = currentPackageInventory();
      if (inventory.textboxInventory?.length) {
        scopes.push({
          scope: "textbox_inventory",
          state: inventory.textboxInventory.some((item) => item.truncated)
            ? "partial"
            : "complete",
          items: inventory.textboxInventory.map((item) => ({
            kind: "textboxInventory",
            nodeId: item.id,
            story: { type: "textbox", id: item.id },
            text: item.textPreview,
            partName: item.partName,
            textLength: item.textLength,
            truncated: item.truncated,
            contentSource: item.contentSource,
            mutationTargetAvailable: false,
          })),
        });
      }
      for (const story of inventory.storyInventory.filter(
        (item) => item.type === "endnote" && item.readable,
      )) {
        try {
          const result = await agentRead({
            ...a,
            representation: "structured",
            story: story.story,
            offset: 0,
            limit: 10000,
          });
          scopes.push({
            scope: `endnote:${story.id}`,
            items: result.items,
            state: "complete",
          });
        } catch {
          scopes.push({
            scope: `endnote:${story.id}`,
            items: [],
            state: "unsupported",
          });
        }
      }
      const all = scopes.flatMap(({ scope, items }) =>
        items.map((item) => ({
          ...item,
          details: { ...(item.details ?? {}), completeReadScope: scope },
        })),
      );
      requireThat(
        all.length <= 20000,
        "LIMIT_EXCEEDED",
        "Complete read exceeds 20,000 canonical items. Read individual domains instead.",
      );
      snapshot = { scopes, all, inventory };
      if (a.readSnapshotId) readSnapshots.set('complete',{id:a.readSnapshotId,revision:currentRevision,generation:readGeneration},snapshot);
    }
    const response = paged(snapshot.all, a, "complete", {
      ...(a.readSnapshotId ? { readSnapshotId: a.readSnapshotId } : {}),
      includedScopes: snapshot.scopes.map(({ scope, items }) => ({
        scope,
        count: items.length,
      })),
    });
    const returnedAll = response.nextOffset === null;
    response.coverage.contentComplete = returnedAll;
    response.coverage.matrix = Object.fromEntries([
      ...snapshot.scopes.map(({ scope, state }) => [
        scope,
        state ?? (returnedAll ? "complete" : "partial"),
      ]),
      [
        "endnotes",
        snapshot.inventory.storyInventory.some((item) => item.type === "endnote")
          ? snapshot.scopes.some(
                (item) =>
                  item.scope.startsWith("endnote:") &&
                  item.state === "unsupported",
              )
            ? "unsupported"
            : returnedAll
              ? "complete"
              : "partial"
          : "complete",
      ],
      [
        "textboxes",
        snapshot.inventory.storyInventory.some((item) => item.type === "textbox")
          ? snapshot.inventory.textboxInventory?.some((item) => item.truncated)
            ? "partial"
            : "complete"
          : "complete",
      ],
      ["visualLayout", "not_checked"],
    ]);
    response.coverage.scopes = snapshot.scopes.map(({ scope, items, state }) => ({
      scope,
      state: state ?? (returnedAll ? "complete" : "partial"),
      count: items.length,
    }));
    response.coverage.dimensions = {
      sdkVisibleContent: returnedAll ? "complete" : "partial",
      storyInventory: "complete",
      annotations: a.includeAnnotations === false ? "not_checked" : "complete",
      visualLayout: "not_checked",
      unsupportedObjects: snapshot.inventory.objectInventory.some(
        (item) => item.count > 0,
      )
        ? "unsupported"
        : "complete",
    };
    response.coverage.notChecked = [
      ...(response.coverage.matrix.endnotes === "unsupported"
        ? ["endnotes-not-exposed-by-SuperDoc"]
        : []),
      ...(snapshot.inventory.storyInventory.some(
        (item) => item.type === "textbox",
      )
        ? ["textbox-mutation-targets-not-exposed-by-SuperDoc"]
        : []),
      "visual-layout",
      "content-not-exposed-by-SuperDoc",
    ];
    return response;
  }
  if (a.representation === "inventory") {
    const info = await doc.info();
    const entries = Object.entries(info.counts ?? {}).map(([name, count]) => ({
      kind: "inventory",
      nodeId: name,
      name,
      text: String(count),
      count,
      path: ["inventory", name],
      story: a.story,
    }));
    const readingPlan = entries
      .filter((entry) => Number(entry.count) > 0)
      .map((entry) => ({
        representation: "domain",
        domain: {
          tables: "tables",
          images: "images",
          comments: "comments",
          trackedChanges: "tracked_changes",
          fields: "fields",
          contentControls: "content_controls",
        }[entry.name] ?? "paragraphs",
        offset: 0,
        limit: a.limit ?? 30,
      }));
    const response = paged(entries, a, "inventory", {
      outline: info.outline ?? [],
      readingPlan,
    });
    response.coverage.matrix = Object.fromEntries([
      ...Object.entries(response.coverage.matrix),
      ...entries
        .filter((entry) => Number(entry.count) > 0)
        .map((entry) => [entry.name, "detected_not_read"]),
    ]);
    response.coverage.remainingReads = readingPlan;
    response.remainingReads = readingPlan;
    return response;
  }
  if (a.representation === "structured" || a.story) {
    const selectedStory = storyOf(a.story);
    const model = await doc.get({ in: selectedStory });
    const nodes = model.body ?? model.content ?? model.blocks ?? [];
    const structured = nodes.map((node, index) =>
      structuredNode(node, [storyLabel(a.story), index], null, a.story),
    );
    if (a.representation === "structured")
      return paged(structured, a, "structured", {
        story: selectedStory,
      });
    const linear = [];
    const visit = (item) => {
      linear.push({ ...item, children: undefined });
      for (const child of item.children ?? []) visit(child);
    };
    for (const item of structured) visit(item);
    return paged(linear, a, "linear", { story: selectedStory });
  }
  if (a.representation === "linear") {
    const extracted = await doc.extract();
    const items = [
      ...extracted.blocks.map((block) => ({
        ...block,
        kind: block.type,
        nativeTarget: {
          kind: "agent_selector",
          selector: { kind: "nodeId", nodeId: block.nodeId },
          address: {
            kind: "block",
            nodeType: ["paragraph", "heading", "listItem"].includes(block.type)
              ? block.type
              : "paragraph",
            nodeId: block.nodeId,
          },
        },
      })),
      ...(a.includeAnnotations
        ? [
            ...extracted.comments.map((comment) => ({
              ...comment,
              kind: "comment",
              nativeTarget: {
                kind: "agent_comment",
                commentId: comment.entityId,
              },
            })),
            ...extracted.trackedChanges.map((change) => ({
              ...change,
              kind: "trackedChange",
              text: change.excerpt ?? "",
              nativeTarget: {
                kind: "agent_change",
                changeId: change.entityId,
              },
            })),
          ]
        : []),
    ];
    return paged(items, a, "linear", {
      documentRevision: extracted.revision,
      annotationsIncluded: a.includeAnnotations,
    });
  }
  if (["markdown", "text"].includes(a.representation)) {
    const value =
      a.representation === "markdown"
        ? await doc.getMarkdown({ in: storyOf(a.story) })
        : await doc.getText({ in: storyOf(a.story) });
    const items = String(value)
      .split(a.representation === "markdown" ? /\n{2,}/ : /\n/)
      .map((text, order) => ({
        kind: a.representation,
        nodeId: `${a.representation}-${order}`,
        text,
      }));
    return paged(items, a, a.representation);
  }
  const result = await read({
    ...a,
    mode: "read",
    domain: a.domain ?? "paragraphs",
  });
  return pagedWindow(
    (result.items ?? []).map(item => {
      const target = item.nativeTarget;
      if (target?.kind !== 'paragraph' || !target.address) return item;
      // Domain paragraph reads used the legacy span locator. Public agent
      // actions need the same typed block locator as linear/query reads.
      return { ...item, nativeTarget: { kind: 'agent_selector',
        selector: { kind: 'nodeId', nodeId: target.nodeId }, address: target.address } };
    }),
    a,
    a.domain ?? "paragraphs",
    result.total ?? result.items?.length ?? 0,
    result.nextOffset ?? null,
    { engine: result.engine },
  );
}

async function agentQuery(a) {
  if (a.packageCellEdit) return sdkCellEdit(doc, currentPath, currentMaxExpandedBytes, a.packageCellEdit.locator, a.packageCellEdit.arguments);
  if (a.select.type === "entity" && a.select.domain === "objects") {
    requireThat(!a.story && !a.withinLocator, "UNSUPPORTED_FEATURE", "Complex object queries use package identities; story/container scopes are not SDK addresses.");
    const inventory = packageInventory ??= inspectDocxPackage(currentPath, currentMaxExpandedBytes);
    const match = (actual, expected) => expected === undefined || (a.select.caseSensitive ? String(actual ?? "") === expected : String(actual ?? "").toLocaleLowerCase() === expected.toLocaleLowerCase());
    const items = inventory.complexObjects.filter(item => match(item.id, a.select.id) && match(item.name, a.select.name) && match(item.text, a.select.text));
    requireThat(a.require !== "exactlyOne" || items.length === 1, "TARGET_CONFLICT", "Complex object selector must resolve exactly once.");
    return paged(items, a, "objects", { engine: "read-only OOXML complex object inventory" });
  }
  return queryAgentDocument(doc, a);
}

async function inlineLocator(step) {
  return resolveInlineLocator(doc, step);
}

async function directSingleTextReplacement(step, locator) {
  return directTextReplacement(doc, step, locator);
}

async function runAgentStep(step) {
  const resolved = await inlineLocator(step);
  const receipt = await executeDocxAction({
    doc,
    step,
    locator: resolved,
    toolkit: await toolkit(),
    placementOf,
    directTextReplacement: directSingleTextReplacement,
  });
  return {
    action: step.action,
    resolvedTargets:
      receipt.selectedTargets ??
      [resolved?.nodeId ?? resolved?.address?.nodeId].filter(Boolean),
    receipt,
  };
}

async function runMutationPlan(task, executeStep, preview = false, hooks = {}) {
  const result = await executeDocxMutationPlan({ client, doc, task, executeStep, preview, discard: dispose,
    restart:async()=>{await dispose();await load(task);return{client,doc};},
    setDocument: (next) => { doc = next; }, resetClient: async()=>{await client.dispose();client=clientFor(task);return client;}, userName: author, ...hooks });
  if (!preview) {
    currentPath = task.output;
    currentRevision = result.saved.revision;
    packageInventory = null;
    invalidateReads();
  }
  return result;
}

async function agentApply(task, preview = false) {
  let plan;
  if (task.args.semanticWorkflow) {
    const locator = task.args.workflowSelect ? await inlineLocator({
      action: task.args.workflow, select: task.args.workflowSelect,
    }) : task.args.workflowLocator;
    plan = await semanticWorkflow(doc, task.args.semanticWorkflow, locator, task);
    task = { ...task, args: { ...task.args, steps: plan.steps } };
  }
  const baseline = task.args.readback ? await captureBodyChanges(doc) : null;
  const { saved, steps: executed, postconditions, semanticDelta, readback, workflowVerification } = await runMutationPlan(task, runAgentStep, preview, {
    verifyWorkflow: plan?.verify,
    ...(baseline ? { captureReadback: async (preview, reopened = doc) => bodyReadback(baseline, await captureBodyChanges(reopened), task.args.readback, preview) } : {}),
  });
  const evidence = { ...(readback ? { readback } : {}), ...(workflowVerification ? { workflowVerification } : {}) };
  if (preview) {
    return {
      preview: true,
      applied: false,
      wouldChange: executed.length > 0,
      atomicity: "isolated-SuperDoc-session-discarded",
      steps: executed,
      ...evidence,
      ...(task.args.workflow ? { workflow: task.args.workflow } : {}),
      ...(postconditions ? { postconditions, semanticDelta } : {}),
      nextAction:
        "Review the resolved targets and receipts, then repeat with mode=apply and an operationId.",
    };
  }
  const receipt = {
    status: "ok",
    intent: executed.length === 1 ? executed[0].action : "mutation_plan",
    verificationPassed: true,
    steps: executed,
  };
  return {
    ...saved,
    action: receipt.intent,
    receipt,
    affectedTargets: executed.flatMap((step) => step.resolvedTargets),
    ...evidence,
    ...(task.args.workflow ? { workflow: task.args.workflow } : {}),
    ...(postconditions ? { postconditions, semanticDelta } : {}),
  };
}

function placementOf(args, locator) {
  const placement = args.placement ?? "end";
  if (placement === "start") return { kind: "documentStart" };
  if (placement === "end") return { kind: "documentEnd" };
  requireThat(
    locator?.address,
    "INVALID_ARGUMENT",
    `${placement} requires a block selectorTarget.`,
  );
  return { kind: placement, target: locator.address };
}

async function agentCompare(task) {
  const targetSnapshot = await doc.diff.capture();
  const base = await client.open({ doc: task.original, runtime: "v2" });
  try {
    const compared = await base.diff.compare({ targetSnapshot });
    const hasChanges = compared.summary?.hasChanges ?? false;
    const changedComponents = compared.summary?.changedComponents ?? [];
    return {
      hasChanges,
      changedComponents,
      summary: {
        hasChanges,
        changedComponentCount: changedComponents.length,
      },
      coverage: {
        engine: ENGINE,
        semanticSnapshotCompared: true,
        comparedComponents: Object.keys(
          compared.coverage ?? targetSnapshot.coverage ?? {},
        ),
        notVerified: [
          "visual-layout",
          "content-not-exposed-by-SuperDoc",
          "exact-XML-equivalence",
        ],
      },
      baseFingerprint: compared.baseFingerprint,
      targetFingerprint: compared.targetFingerprint,
    };
  } finally {
    await base.close({ discard: true });
  }
}
const locator = (p, index, start = 0, end = textOf(p).length) => ({
  kind: start === 0 && end === textOf(p).length ? "paragraph" : "text",
  nodeId: p.address.nodeId,
  address: p.address,
  paragraphIndex: index,
  start,
  end,
  text: textOf(p).slice(start, end),
  paragraphText: textOf(p),
});
const page = (items, a, domain) => ({
  items: items.slice(a.offset, a.offset + a.limit),
  total: items.length,
  offset: a.offset,
  nextOffset: a.offset + a.limit < items.length ? a.offset + a.limit : null,
  engine: ENGINE,
  coverage: {
    completeDocument: false,
    windowComplete: a.offset + a.limit >= items.length,
    features: [{ feature: domain, state: "read", count: items.length }],
    notChecked: [
      "content-not-exposed-by-SuperDoc",
      "other-domains",
      "visual-layout",
    ],
  },
});
async function read(a) {
  if (a.mode === "inspect") {
    const info = await doc.info(),
      caps = await doc.capabilities.get();
    return {
      counts: info.counts,
      outline: info.outline,
      sdkCapabilityCount: Object.keys(caps.operations ?? {}).length,
      engine: ENGINE,
    };
  }
  if (a.mode === "find") {
    const found = [];
    for (const [i, p] of (await paragraphs(doc)).entries()) {
      const text = textOf(p),
        pattern = new RegExp(
          a.text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"),
          a.caseSensitive ? "gu" : "giu",
        );
      for (const m of text.matchAll(pattern)) {
        found.push({
          kind: "text",
          text: m[0],
          context: text.slice(
            Math.max(0, m.index - 60),
            m.index + m[0].length + 60,
          ),
          nativeTarget: locator(p, i, m.index, m.index + m[0].length),
        });
        requireThat(
          found.length <= 10000,
          "LIMIT_EXCEEDED",
          "Narrow the text query.",
        );
      }
    }
    return page(found, a, "body-text");
  }
  const map = {
    styles: "styles",
    lists: "lists",
    tables: "tables",
    table_cells: "tables",
    table_layout: "tables",
    sections: "sections",
    page_setup: "sections",
    images: "images",
    image_layout: "images",
    headers_footers: "headers_footers",
    hyperlinks: "hyperlinks",
    bookmarks: "bookmarks",
    notes: "footnotes",
    fields: "fields",
    toc: "toc",
    content_controls: "content_controls",
    review: "tracked_changes",
  };
  const domain =
    a.action === "list"
      ? map[a.tool.slice(5)]
      : (a.docxDomain ?? a.domain ?? "paragraphs");
  let items;
  if (domain === "paragraphs" || domain === "content") {
    const changes = await list(doc, "trackChanges.list");
    items = (await paragraphs(doc)).map((p, i) => ({
      kind: p.node.kind,
      text: textOf(p),
      node: p.node,
      reviewChanges: changes
        .filter((c) => c.navigationTarget?.blockId === p.address.nodeId)
        .flatMap((c) =>
          Object.entries(c.wordRevisionIds ?? {})
            .filter(([, id]) => id)
            .map(([type, id]) => ({
              id,
              type,
              author: c.author,
              text: type === "insert" ? c.insertedText : c.deletedText,
            })),
        ),
      nativeTarget: locator(p, i),
    }));
    if (domain === "content") {
      for (const [index, table] of (await tables(doc)).entries())
        items.push({
          kind: "table",
          node: table.node,
          rows: table.node.table.rows.map((r) =>
            r.cells.map((c) => ({
              paragraphs: c.content.map((n) => ({ text: textOf(n) })),
            })),
          ),
          nativeTarget: {
            kind: "table",
            address: table.address,
            nodeId: table.address.nodeId,
            index,
          },
        });
      for (const h of await headers(doc))
        if (h.resolution.refId)
          items.push({
            kind: "paragraph",
            story: h.kind,
            text: h.text,
            resolution: h.resolution,
          });
      for (const n of await list(doc, "footnotes.list"))
        items.push({ kind: "note", text: n.content, ...n });
      for (const image of await list(doc, "images.list"))
        items.push({
          kind: "image",
          ...image,
          source: { kind: "sdk_image", imageId: image.sdImageId },
          text: image.properties?.description ?? "",
        });
      for (const comment of await list(doc, "comments.list"))
        items.push({ kind: "existing-comment", ...comment });
      for (const f of await list(doc, "fields.list"))
        items.push({ kind: "field", ...f });
    }
    return page(items, a, domain);
  }
  if (domain === "capabilities")
    items = Object.entries((await doc.capabilities.get()).operations ?? {}).map(
      ([operation, support]) => ({ operation, ...support }),
    );
  else if (domain === "styles")
    items = (await doc.styles.getCatalog({ includeUsage: true })).items;
  else if (domain === "tables")
    items = (await tables(doc)).map((i, index) => ({
      ...i,
      kind: "table",
      nativeTarget: {
        kind: "table",
        address: i.address,
        nodeId: i.address.nodeId,
        index,
      },
    }));
  else if (domain === "headers_footers")
    items = (await headers(doc)).map((i) => ({
      ...i,
      ...(i.sdkStory
        ? {
            nativeTarget: {
              kind: "header_footer",
              story: i.sdkStory,
              nodeId: i.id,
              text: i.text,
            },
          }
        : {}),
    }));
  else {
    items = await readDocxDomain(doc, domain);
  }
  return page(items, a, domain);
}
async function edit(task) {
  const steps = [...task.args.steps].sort((a, b) =>
    a.locator.nodeId === b.locator.nodeId
      ? (b.locator.start ?? 0) - (a.locator.start ?? 0)
      : a.locator.nodeId.localeCompare(b.locator.nodeId),
  );
  const seen = new Map();
  const result = await runMutationPlan({ ...task, args: { ...task.args, steps } }, async (step) => {
    const old = seen.get(step.locator.nodeId);
    requireThat(!old || (step.type === "text" && old.type === "text" && step.locator.end <= old.start),
      "TARGET_CONFLICT", "Overlapping batch targets.");
    seen.set(step.locator.nodeId, { type: step.type, start: step.locator.start });
    return mutate(doc, step);
  });
  return { ...result.saved, affectedTargets: steps.map((step) => step.target.id) };
}

installDocxHostProcess({
  invalidateReads,
  prepareValidation:task=>task.operations?.length?Promise.resolve():load(task),
  dispose,
  load,
  create: createSdk,
  edit,
  validate: (task) => validateDocxCandidate(doc, task),
  agentApply,
  agentInspect:(args,task)=>cachedRead('inspect',args,()=>agentInspect(args,task)),
  agentRead:args=>cachedRead('read',args,()=>agentRead(args)),
  agentQuery:args=>cachedRead('query',args,()=>agentQuery(args)),
  agentCompare,
  read:args=>cachedRead('compatibility-read',args,()=>read(args)),
  loads: () => loads,
  cacheStats:()=>({reads:readCache.stats(),snapshots:readSnapshots.stats(),generation:readGeneration,revision:currentRevision}),
});
