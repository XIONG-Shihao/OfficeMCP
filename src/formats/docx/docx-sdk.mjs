import { createDocxEngine as createSuperDocClient } from "#docx-core";
import fs from "node:fs";
import { digest, fileRevision, fingerprint } from "../../infrastructure/files.mjs";
import { requireThat, OfficeError } from "../../infrastructure/errors.mjs";
import {tracePhase} from "../../infrastructure/docx-phase-trace.mjs";
import {installedSdkBinary} from './docx-native-binary.mjs';
export const ENGINE = "SuperDoc SDK 2.2.1 / v2 (SDK-only)";
export const textOf = (n) =>
  (n?.node?.[n.node.kind]?.inlines ?? n?.[n.kind]?.inlines ?? [])
    .map((i) => i.run?.text ?? "")
    .join("");
export async function paragraphs(doc) {
  const model = await doc
      .get({ options: { includeResolved: true } })
      .catch((e) => {
        throw new OfficeError(
          "ENGINE_FAILED",
          "SDK paragraph model: " + e.message,
        );
      }),
    items = [];
  function walk(blocks) {
    for (const node of blocks ?? []) {
      if (["paragraph", "heading", "listItem"].includes(node.kind)) {
        items.push({
          address: {
            kind: "block",
            nodeType: node.kind,
            nodeId: node.paragraphIds?.paraId ?? node.id,
          },
          node,
        });
      } else if (node.kind === "table")
        for (const r of node.table.rows)
          for (const c of r.cells) walk(c.content);
      else if (node.kind === "sdt") walk(node.sdt?.content);
    }
  }
  walk(model.body);
  requireThat(
    items.length <= 10000,
    "LIMIT_EXCEEDED",
    "SDK paragraph inventory exceeds 10000.",
  );
  return items;
}

/** Resolve an existing body container through the same model used by paragraphs. */
export async function descendantBlockIds(doc, address) {
  requireThat(address?.kind === "block" && (!address.story || address.story.storyType === "body"),
    "UNSUPPORTED_FEATURE", "Structural withinTarget requires a body block container.");
  const model = await doc.get({ options: { includeResolved: true } });
  const ids = new Set();
  let found = false, visited = 0;
  const walk = (nodes, inside = false) => {
    for (const node of nodes ?? []) {
      requireThat(++visited <= 20000, "LIMIT_EXCEEDED", "Container inventory exceeds 20000 nodes.");
      const nodeIds = [node.id, node.paragraphIds?.paraId, node.attrs?.sdBlockId].filter(Boolean);
      const selected = inside || nodeIds.includes(address.nodeId);
      if (nodeIds.includes(address.nodeId)) found = true;
      if (selected) for (const id of nodeIds) ids.add(id);
      walk(node.content ?? node.sdt?.content, selected);
      for (const row of node.table?.rows ?? [])
        for (const cell of row.cells ?? []) walk(cell.content, selected);
    }
  };
  walk(model.body);
  requireThat(found, "UNSUPPORTED_FEATURE", "The selected container is not exposed by the SDK body model. Use a scoped text/node query instead.");
  return ids;
}
export async function list(doc, name, options = {}) {
  let owner = doc;
  const keys = name.split(".");
  for (const k of keys.slice(0, -1)) owner = owner[k];
  const method = owner[keys.at(-1)].bind(owner);
  let items = [],
    offset = 0;
  while (true) {
    let r;
    try {
      r = await method({ ...options, offset, limit: 100 });
    } catch (e) {
      throw new OfficeError("ENGINE_FAILED", "SDK " + name + ": " + e.message);
    }
    requireThat(
      Array.isArray(r.items),
      "ENGINE_FAILED",
      "Unexpected SDK list response: " + name,
    );
    items.push(...r.items);
    requireThat(
      items.length <= 10000,
      "LIMIT_EXCEEDED",
      "SDK inventory exceeds 10000: " + name,
    );
    if (items.length >= r.total) return items;
    requireThat(
      r.items.length,
      "ENGINE_FAILED",
      "SDK pagination did not advance.",
    );
    offset += r.items.length;
  }
}
export async function headers(doc) {
  const slots = await list(doc, "headerFooters.list"),
    parts = await list(doc, "headerFooters.parts.list");
  for (const s of slots) {
    s.resolution = await doc.headerFooters.resolve({
      target: {
        kind: "headerFooterSlot",
        section: s.section,
        headerFooterKind: s.kind,
        variant: s.variant,
      },
    });
    if (s.resolution.refId) {
      s.text = await doc.getText({
        in: {
          kind: "story",
          storyType: "headerFooterPart",
          refId: s.resolution.refId,
        },
      });
      s.source = {
        kind: "sdk_story",
        name: parts.find((p) => p.refId === s.resolution.refId)?.partPath,
      };
      s.textPreview = s.text;
      s.contentComplete = true;
      s.sdkStory = {
        kind: "story",
        storyType: "headerFooterPart",
        refId: s.resolution.refId,
      };
    }
  }
  return slots;
}
export function clean(value) {
  if (Array.isArray(value)) return value.map(clean);
  if (value && typeof value === "object")
    return Object.fromEntries(
      Object.entries(value)
        .filter(
          ([k]) =>
            ![
              "handle",
              "evaluatedRevision",
              "revision",
              "raw",
              "selectionTarget",
              "ref",
              "id",
              "sdImageId",
              "address",
              "target",
              "paragraphIds",
              "anchorRef",
            ].includes(k),
        )
        .map(([k, v]) => [k, clean(v)]),
    );
  return value;
}
// The SDK projects a newly created tracked insertion/deletion into comments.list
// as a virtual tc-comment. It is not an authored OOXML comment and may disappear
// after reopening. Ignore only a projection whose full linkage/text/author agrees
// with the independently verified tracked-change record; retain real/malformed comments.
function derivedTrackedComment(comment, changes) {
  const link = comment.trackedChangeLink;
  if (comment.trackedChange !== true || link?.trackedChange !== true ||
      ["text", "content", "body", "commentText", "replies"].some(key => Object.hasOwn(comment, key))) return false;
  const change = changes.find(item => item.id === link.trackedChangeId && item.id === comment.trackedChangeParentId);
  if (!change?.revisionGroupId || !["insertion", "deletion"].includes(change.type)) return false;
  const side = change.type === "insertion" ? "inserted" : "deleted";
  const field = side + "Text", text = change[field];
  return typeof text === "string" && comment.commentId === `tc-comment:${change.revisionGroupId}` &&
    comment.status === "open" && comment.creatorName === change.author &&
    comment.createdTime === Date.parse(change.date) &&
    comment.trackedChangeType === (side === "inserted" ? "insert" : "delete") &&
    comment.trackedChangeSide === side && link.side === side && link.trackedChangeType === change.type &&
    comment.trackedChangeAnchorKey === change.revisionGroupId && link.trackedChangeAnchorKey === change.revisionGroupId &&
    Boolean(change.navigationTarget?.story && comment.trackedChangeStory && link.trackedChangeStory) &&
    fingerprint(comment.trackedChangeStory) === fingerprint(change.navigationTarget.story) &&
    fingerprint(link.trackedChangeStory) === fingerprint(change.navigationTarget.story) &&
    comment[field] === text && link[field] === text &&
    comment.trackedChangeText === text && link.trackedChangeText === text && comment.anchoredText === text;
}
export async function snapshot(doc, includeBookmarks = false) {
  const domains = {};
  let trackedChanges = [];
  for (const name of [
    "images.list",
    "sections.list",
    "footnotes.list",
    "fields.list",
    "hyperlinks.list",
    "contentControls.list",
    "trackChanges.list",
    "comments.list",
    "captions.list",
    "crossRefs.list",
  ]) {
    const items = await list(doc, name);
    if (name === "trackChanges.list") trackedChanges = items;
    domains[name] = clean(name === "comments.list"
      ? items.filter(comment => !derivedTrackedComment(comment, trackedChanges))
      : items);
  }
  if (includeBookmarks)
    domains["bookmarks.list"] = clean(await list(doc, "bookmarks.list"));
  const parts = await list(doc, "headerFooters.parts.list"),
    stories = [];
  for (const part of parts)
    stories.push({
      ...part,
      text: await doc.getText({
        in: { kind: "story", storyType: "headerFooterPart", refId: part.refId },
      }),
    });
  return {
    model: clean(await doc.get({ options: { includeResolved: true } })),
    headers: clean(stories),
    domains,
  };
}
export function success(r) {
  requireThat(
    r?.success !== false && r?.receipt?.success !== false,
    "ENGINE_FAILED",
    "SuperDoc rejected the mutation.",
  );
  return r;
}
export function clientFor(task) {
  const binary=process.env.SUPERDOC_CLI_BIN??installedSdkBinary();
  return createSuperDocClient({
    env: { SUPERDOC_CLI_STATE_DIR: task.scratch, BUN_BE_BUN: undefined, ...(binary?{SUPERDOC_CLI_BIN:binary}:{}) },
    requestTimeoutMs: task.timeoutMs,
    watchdogTimeoutMs: task.timeoutMs,
    startupTimeoutMs: Math.min(10000, task.timeoutMs),
  });
}
export async function saveChecked(client, doc, task, inputRevision, { setDocument, resetClient, userName, sequentialReopen, inspectReopened } = {}) {
  const includeBookmarks = (task.args.steps ?? []).some(
    (s) =>
      s.type === "bookmarks" ||
      ["rename_bookmark", "remove_bookmark"].includes(s.action),
  );
  let expectedHash;
  try {
    // Retain the digest, not the entire model, while the SDK writes/reopens.
    expectedHash = fingerprint(await tracePhase('semanticBefore',()=>snapshot(doc, includeBookmarks)));
  } catch (e) {
    throw new OfficeError("ENGINE_FAILED", "Before save: " + e.message);
  }
  await tracePhase('sdkExport',()=>doc.save({ out: task.output, force: false }));
  const outputRevision = fileRevision(task.output, task.maxFileBytes);
  const releaseOriginal = sequentialReopen ?? Math.max(
    fs.statSync(task.output).size,
    task.input ? fs.statSync(task.input).size : 0,
  ) > 64 * 1024 * 1024;
  if (releaseOriginal) {
    requireThat(typeof setDocument === "function", "INTERNAL_ERROR", "Sequential SDK reopen requires a document owner.");
    await doc.close({ discard: true });
    setDocument(null);
    // Closing one document does not require the SDK process to return its old
    // parse/export heap to the OS. Retire that process before the large reopen.
    if(resetClient)client=await tracePhase('sdkRuntimeReset',resetClient);
  }
  let reopened, readback;
  try {
    reopened = await tracePhase('sdkReopen',()=>client.open({ doc: task.output, runtime: "v2", ...(userName ? { userName } : {}) }));
    const actualHash = fingerprint(await tracePhase('semanticAfter',()=>snapshot(reopened, includeBookmarks)));
    requireThat(
      actualHash === expectedHash,
      "VALIDATION_FAILED",
      "SDK saved/reopened document differs from its pre-save semantic snapshot.",
    );
    readback = await inspectReopened?.(reopened);
    if (releaseOriginal) {
      setDocument(reopened);
      reopened = null; // Ownership transferred only after successful verification.
    }
  } finally {
    if (reopened) await reopened.close({ discard: true });
  }
  return {
    revision: outputRevision,
    ...(readback ? { readback } : {}),
    engine: ENGINE,
    engineEvidence: {
      engine: ENGINE,
      verification: "superdoc-sdk-readback-v1",
      inputRevision,
      outputRevision,
      sdkSnapshotHash: expectedHash,
      includeBookmarks,
      evaluation: false,
      postconditions: [
        "SDK-request-postconditions",
        "saved-document-reopened-with-SDK",
        "SDK-semantic-snapshot-roundtrip",
      ],
      notVerified: [
        ...(!includeBookmarks ? ["untouched-bookmark-metadata"] : []),
        "independent-package-preservation",
        "content-not-exposed-by-SDK",
        "visual-layout",
        "binary-resource-equivalence",
      ],
    },
  };
}
export async function createSdk(task) {
  const client = clientFor(task);
  let doc;
  try {
    doc = await client.open({ runtime: "v2", userName: "OfficeMCP" });
    const spec = task.args.document;
    for (const b of spec.blocks) {
      if (b.kind === "table") {
        requireThat(
          b.rows.every((r) => r.length === b.rows[0].length),
          "INVALID_ARGUMENT",
          "SDK table creation requires a rectangular matrix.",
        );
        const r = success(
          await doc.create.table({
            rows: b.rows.length,
            columns: b.rows[0].length,
            at: { kind: "documentEnd" },
          }),
        );
        const target =
          r.table ?? (await doc.find({ type: "table" })).items.at(-1).address;
        for (let row = 0; row < b.rows.length; row++)
          for (let column = 0; column < b.rows[row].length; column++)
            success(
              await doc.tables.setCellText({
                target,
                rowIndex: row,
                columnIndex: column,
                text: b.rows[row][column],
              }),
            );
      } else if (b.kind === "image") {
        const file = task.assets?.[b.assetId ?? b.path];
        requireThat(
          file,
          "ACCESS_DENIED",
          "Creation image must be an owned local asset.",
        );
        const bytes = fs.readFileSync(file);
        success(
          await doc.create.image({
            src:
              "data:image/" +
              (bytes[0] === 137 ? "png" : "jpeg") +
              ";base64," +
              bytes.toString("base64"),
            alt: b.alt,
            size: { width: (b.width * 4) / 3, height: (b.height * 4) / 3 },
            at: { kind: "documentEnd" },
          }),
        );
      } else {
        const r = success(
          await doc.create.paragraph({
            at: { kind: "documentEnd" },
            text: b.text,
          }),
        );
        if (b.kind === "heading")
          success(
            await doc.styles.paragraph.setStyle({
              target: r.paragraph,
              styleId: "Heading" + (b.level ?? 1),
            }),
          );
        if (b.kind === "bullet")
          success(
            await doc.lists.create({
              mode: "fromParagraphs",
              target: r.paragraph,
              kind: "bullet",
              level: 0,
            }),
          );
        if (b.style && b.text)
          success(
            await doc.format.apply({
              blockId: r.paragraph.nodeId,
              start: 0,
              end: b.text.length,
              inline: {
                ...b.style,
                ...(b.style.color ? { color: "#" + b.style.color } : {}),
              },
              changeMode: "direct",
            }),
          );
      }
    }
    // The SDK blank seed contains one empty paragraph; remove it through the SDK.
    const first = (await paragraphs(doc))[0];
    if (first && textOf(first) === "" && (await paragraphs(doc)).length > 1)
      success(await doc.blocks.delete({ target: first.address }));
    for (const kind of ["header", "footer"])
      if (spec[kind] !== undefined) {
        const part = success(await doc.headerFooters.parts.create({ kind })),
          section = (await list(doc, "sections.list"))[0];
        success(
          await doc.create.paragraph({
            in: {
              kind: "story",
              storyType: "headerFooterPart",
              refId: part.refId,
            },
            at: { kind: "documentEnd" },
            text: spec[kind],
          }),
        );
        success(
          await doc.headerFooters.refs.set({
            target: {
              kind: "headerFooterSlot",
              section: section.address,
              headerFooterKind: kind,
              variant: "default",
            },
            refId: part.refId,
          }),
        );
      }
    const result = await saveChecked(client, doc, task, null, { setDocument: (next) => { doc = next; }, userName: "OfficeMCP" });
    return {
      ...result,
      format: "docx",
      bytes: fs.statSync(task.output).size,
      warnings: [
        ...(spec.title
          ? ["Document title metadata is not set by this SDK creation adapter."]
          : []),
        "Creation and validation use SuperDoc only; no page rendering or independent OOXML validation.",
      ],
    };
  } finally {
    if (doc) await doc.close({ discard: true });
    await client.dispose();
  }
}

export async function tables(doc) {
  const items = [];
  let offset = 0;
  while (true) {
    const r = await doc.find({ type: "table", offset, limit: 100 });
    items.push(...r.items);
    requireThat(
      items.length <= 1000,
      "LIMIT_EXCEEDED",
      "SDK table inventory exceeds 1000.",
    );
    if (items.length >= r.total) return items;
    requireThat(
      r.items.length,
      "ENGINE_FAILED",
      "Table pagination did not advance.",
    );
    offset += r.items.length;
  }
}
