import { readFileSync } from "node:fs";
import path from "node:path";
import { decodePackage, validateXml, xmlPart, relationships } from "./ooxml.mjs";
import { digest } from "../../infrastructure/files.mjs";
import { requireThat } from "../../infrastructure/errors.mjs";

const CONTENT_TYPES =
  "http://schemas.openxmlformats.org/package/2006/content-types";

function contentTypeEvidence(parts) {
  const document = xmlPart(parts, "[Content_Types].xml");
  requireThat(
    document.documentElement.namespaceURI === CONTENT_TYPES &&
      document.documentElement.localName === "Types",
    "INVALID_DOCUMENT",
    "Invalid OOXML content-types root.",
  );
  const defaults = new Map();
  const overrides = new Map();
  for (const node of Array.from(document.documentElement.childNodes).filter(
    (item) => item.nodeType === 1,
  )) {
    if (node.localName === "Default") {
      const extension = node.getAttribute("Extension")?.toLowerCase();
      const type = node.getAttribute("ContentType");
      requireThat(
        extension && type && !defaults.has(extension),
        "INVALID_DOCUMENT",
        "Duplicate or malformed default content type.",
      );
      defaults.set(extension, type);
    } else if (node.localName === "Override") {
      const name = node.getAttribute("PartName")?.replace(/^\//, "");
      const type = node.getAttribute("ContentType");
      requireThat(
        name && type && Object.hasOwn(parts,name) && !overrides.has(name),
        "INVALID_DOCUMENT",
        "Missing, duplicate or malformed content-type override.",
      );
      overrides.set(name, type);
    }
  }
  const uncovered = Object.keys(parts).filter((name) => {
    if (
      name === "[Content_Types].xml" ||
      name.endsWith("/") ||
      name.endsWith(".rels")
    )
      return false;
    const extension = path.posix.extname(name).slice(1).toLowerCase();
    return !overrides.has(name) && !defaults.has(extension);
  });
  requireThat(
    uncovered.length === 0,
    "INVALID_DOCUMENT",
    `OOXML parts lack content types: ${uncovered.slice(0, 5).join(", ")}.`,
  );
  return { defaults: defaults.size, overrides: overrides.size, uncovered };
}

function sourceForRelationshipPart(name) {
  if (name === "_rels/.rels") return "";
  const marker = "/_rels/";
  const at = name.lastIndexOf(marker);
  if (at < 0 || !name.endsWith(".rels")) return null;
  return `${name.slice(0, at)}/${name.slice(at + marker.length, -5)}`;
}

function relationshipEvidence(parts) {
  const all = [];
  for (const name of Object.keys(parts)) {
    const source = sourceForRelationshipPart(name);
    if (source === null) continue;
    const items = relationships(parts, source);
    for (const relationship of items) {
      requireThat(
        /^[a-z][a-z0-9+.-]*:/i.test(relationship.type),
        "INVALID_DOCUMENT",
        "OOXML relationship type must be an absolute URI.",
      );
      requireThat(
        relationship.external || relationship.available,
        "INVALID_DOCUMENT",
        `Missing relationship target: ${relationship.part}.`,
      );
      all.push({ source, ...relationship });
    }
  }
  const referenced = new Set(
    all.filter((item) => !item.external).map((item) => item.part),
  );
  const orphanedMedia = Object.keys(parts).filter(
    (name) => name.startsWith("word/media/") && !referenced.has(name),
  );
  return {
    relationships: all.length,
    externalRelationships: all.filter((item) => item.external).length,
    missingTargets: 0,
    orphanedMedia,
  };
}

function mediaRelationships(parts) {
  const items = new Map();
  for (const name of Object.keys(parts)) {
    const source = sourceForRelationshipPart(name);
    if (source === null) continue;
    for (const relationship of relationships(parts, source)) {
      if (relationship.external || !/\/image$/i.test(relationship.type)) continue;
      items.set(`${source}|${relationship.id}`, {
        source,
        id: relationship.id,
        type: relationship.type,
        part: relationship.part,
      });
    }
  }
  return items;
}

function verifyMediaRelationshipChanges(before, after, mediaChanges) {
  const beforeItems = mediaRelationships(before);
  const afterItems = mediaRelationships(after);
  const allowedSourceParts = new Set(
    mediaChanges
      .filter((change) => ["replace_image", "delete_image"].includes(change.action))
      .map((change) => normalizedPart(change.partName))
      .filter(Boolean),
  );
  const expectedHashes = new Set(
    mediaChanges.map((change) => change.replacementHash).filter(Boolean),
  );
  const removed = [...beforeItems]
    .filter(([key]) => !afterItems.has(key))
    .map(([, value]) => value);
  const added = [...afterItems]
    .filter(([key]) => !beforeItems.has(key))
    .map(([, value]) => value);
  const changed = [...beforeItems]
    .filter(
      ([key, value]) =>
        afterItems.has(key) && afterItems.get(key).part !== value.part,
    )
    .map(([key, beforeValue]) => ({
      before: beforeValue,
      after: afterItems.get(key),
    }));
  const unauthorizedRemoved = removed.filter(
    (item) => !allowedSourceParts.has(normalizedPart(item.part)),
  );
  const unauthorizedChangedSources = changed.filter(
    (item) => !allowedSourceParts.has(normalizedPart(item.before.part)),
  );
  const afterTargetAuthorized = (item) => {
    const bytes = after[item.part];
    return bytes && expectedHashes.has(digest(bytes));
  };
  const unauthorizedAdded = added.filter((item) => !afterTargetAuthorized(item));
  const unauthorizedChangedTargets = changed.filter(
    (item) => !afterTargetAuthorized(item.after),
  );
  const unauthorized = [
    ...unauthorizedRemoved,
    ...unauthorizedChangedSources,
    ...unauthorizedAdded,
    ...unauthorizedChangedTargets,
  ];
  requireThat(
    unauthorized.length === 0,
    "PRESERVATION_FAILED",
    "Image edit changed an unrelated or unauthorized package relationship.",
  );
  return { removed, added, changed, unauthorizedCount: 0 };
}

function binaryEntries(parts) {
  return Object.keys(parts).filter(
    name => !/\.(xml|rels|vml)$/i.test(name) && !name.endsWith("/"),
  );
}
// Exact byte equality is stronger than comparing two SHA digests and avoids
// hashing every large unchanged image a second time during package inspection.
const equalBytes = (left, right) => left.byteLength === right.byteLength &&
  Buffer.from(left.buffer,left.byteOffset,left.byteLength).equals(Buffer.from(right.buffer,right.byteOffset,right.byteLength));

function normalizedPart(name) {
  const value = String(name ?? "").replace(/^\//, "");
  return value.startsWith("media/") ? `word/${value}` : value;
}

function verifyBinaryChanges(before, after, mediaChanges) {
  const changes = mediaChanges ?? [];
  const allowedSourceParts = new Set(
    changes
      .filter((change) => ["replace_image", "delete_image"].includes(change.action))
      .map((change) => normalizedPart(change.partName))
      .filter(Boolean),
  );
  const expectedContentHashes = new Set(
    changes
      .filter((change) => ["insert_image", "replace_image"].includes(change.action))
      .map((change) => change.replacementHash)
      .filter(Boolean),
  );
  const beforeEntries = new Set(
    binaryEntries(before),
  );
  const afterEntries = new Set(
    binaryEntries(after),
  );
  const removedParts = [...beforeEntries].filter(
    name => !afterEntries.has(name),
  );
  const addedParts = [...afterEntries].filter(
    name => !beforeEntries.has(name),
  );
  const changedParts = [...beforeEntries].filter(
    name => afterEntries.has(name) && !equalBytes(before[name],after[name]),
  );
  const unauthorizedSourceChanges = [...removedParts, ...changedParts].filter(
    (name) => !allowedSourceParts.has(normalizedPart(name)),
  );
  requireThat(
    unauthorizedSourceChanges.length === 0,
    "PRESERVATION_FAILED",
    `Engine removed or changed unrelated binary package members: ${unauthorizedSourceChanges.slice(0, 5).join(", ")}. Candidate was not accepted.`,
  );
  const unauthorizedAddedParts = addedParts.filter((name) => {
    const hash = digest(after[name]);
    return !name.startsWith("word/media/") || !expectedContentHashes.has(hash);
  });
  requireThat(
    unauthorizedAddedParts.length === 0,
    "PRESERVATION_FAILED",
    `Media edit added unauthorized binary package members: ${unauthorizedAddedParts.slice(0, 5).join(", ")}.`,
  );
  const unauthorizedReplacementBytes = changedParts.filter(
    (name) => !expectedContentHashes.has(digest(after[name])),
  );
  requireThat(
    unauthorizedReplacementBytes.length === 0,
    "PRESERVATION_FAILED",
    `Media edit wrote unexpected bytes to: ${unauthorizedReplacementBytes.slice(0, 5).join(", ")}.`,
  );
  const suppliedHashes = expectedContentHashes.size ? new Set([...afterEntries].map(name=>digest(after[name]))) : new Set();
  const missingReplacementHashes = [...expectedContentHashes].filter(hash => !suppliedHashes.has(hash));
  requireThat(
    missingReplacementHashes.length === 0,
    "PRESERVATION_FAILED",
    "The accepted package does not contain the exact supplied image bytes.",
  );
  return {
    allowedRemovedParts: [...allowedSourceParts],
    allowedChangedParts: [...allowedSourceParts],
    allowedAddedContentHashes: [...expectedContentHashes],
    actualRemovedParts: removedParts,
    actualChangedParts: changedParts,
    actualAddedParts: addedParts,
    preserveAllOtherBinaryMembers: true,
    unauthorizedSourceChanges,
    unauthorizedAddedParts,
    missingReplacementHashes,
  };
}

function storyParts(parts) {
  return Object.keys(parts)
    .filter((name) =>
      /^word\/(header\d+|footer\d+|footnotes|endnotes|comments)\.xml$/i.test(
        name,
      ),
    )
    .filter((name) => {
      if (/^word\/(header|footer)\d+\.xml$/i.test(name)) return true;
      const document = xmlPart(parts, name);
      if (/word\/comments\.xml$/i.test(name))
        return Array.from(document.getElementsByTagNameNS("*", "comment")).length > 0;
      const localName = /word\/footnotes\.xml$/i.test(name)
        ? "footnote"
        : "endnote";
      return Array.from(document.getElementsByTagNameNS("*", localName)).some(
        (node) =>
          !node.getAttribute("w:type") &&
          !node.getAttribute("type") &&
          Number(node.getAttribute("w:id") || node.getAttribute("id")) >= 0,
      );
    })
    .sort();
}

// Engine round trips may normalize XML. This profile deliberately does not
// claim byte preservation of XML, or visual equivalence.
export function packageEvidence(
  beforeBytes,
  afterBytes,
  limit,
  { mediaChanges = [],nativeInflate=false } = {},
) {
  const beforePackage = decodePackage(beforeBytes, limit,{nativeInflate}),
    afterPackage = decodePackage(afterBytes, limit,{nativeInflate}),
    before = beforePackage.parts,
    after = afterPackage.parts;
  const xml = (n) => /\.(xml|rels|vml)$/i.test(n);
  for (const n of Object.keys(after)) {
    if (!xml(n)) continue;
    // Keep reused metadata DOMs, but do not retain every large document/story
    // tree merely to establish well-formedness during preservation validation.
    if (afterPackage.xml.has(n) || n.endsWith(".rels") || n === "[Content_Types].xml")
      xmlPart(after, n);
    else validateXml(after[n]);
  }
  const contentTypes = contentTypeEvidence(after);
  const beforeRelationshipGraph = relationshipEvidence(before);
  const relationshipGraph = relationshipEvidence(after);
  const operationOwnedOrphanParts = new Set(
    mediaChanges
      .filter((change) => ["replace_image", "delete_image"].includes(change.action))
      .map((change) => normalizedPart(change.partName))
      .filter(Boolean),
  );
  const newlyOrphanedMedia = relationshipGraph.orphanedMedia.filter(
    (name) =>
      !beforeRelationshipGraph.orphanedMedia.includes(name) &&
      !operationOwnedOrphanParts.has(normalizedPart(name)),
  );
  requireThat(
    newlyOrphanedMedia.length === 0,
    "PRESERVATION_FAILED",
    `Engine created orphaned media: ${newlyOrphanedMedia.slice(0, 5).join(", ")}.`,
  );
  const beforeStories = storyParts(before);
  const afterStories = new Set(storyParts(after));
  const removedStories = beforeStories.filter((name) => !afterStories.has(name));
  requireThat(
    removedStories.length === 0,
    "PRESERVATION_FAILED",
    `Engine removed document stories: ${removedStories.join(", ")}.`,
  );
  const binaryChanges = verifyBinaryChanges(before, after, mediaChanges);
  const mediaRelationshipChanges = verifyMediaRelationshipChanges(
    before,
    after,
    mediaChanges,
  );
  const removedParts = Object.keys(before).filter((n) => !(n in after));
  const addedParts = Object.keys(after).filter((n) => !(n in before));
  const changedParts = Object.keys(before).filter(
    (n) => after[n] && !equalBytes(before[n],after[n]),
  );
  beforePackage.filePackage?.assertUnchanged();
  afterPackage.filePackage?.assertUnchanged();
  return {
    inputRevision: beforePackage.filePackage?.revision ?? digest(beforePackage.bytes),
    outputRevision: afterPackage.filePackage?.revision ?? digest(afterPackage.bytes),
    binaryMembersPreserved: binaryEntries(before).length,
    removedParts,
    addedParts,
    changedParts,
    contentTypes,
    relationshipGraph,
    operationOwnedOrphanedMedia: relationshipGraph.orphanedMedia.filter(
      (name) => operationOwnedOrphanParts.has(normalizedPart(name)),
    ),
    newlyOrphanedMedia,
    binaryChanges,
    mediaRelationshipChanges,
    removedStories,
    checks: [
      "output-XML-well-formed",
      "OPC-content-types-complete",
      "OPC-relationship-targets-resolve",
      "document-story-parts-preserved",
      mediaChanges.length
        ? "only-targeted-media-binaries-and-relationships-may-change"
        : "all-input-binary-members-preserved-by-content",
    ],
  };
}

export function validateEngine(task, bytes) {
  let revision = digest(readFileSync(task.original));
  for (const op of task.operations) {
    requireThat(
      op.engineEvidence?.inputRevision === revision,
      "VALIDATION_FAILED",
      "Engine evidence chain is incomplete; do not mix preservation profiles in a candidate.",
    );
    revision = op.engineEvidence.outputRevision;
  }
  requireThat(
    revision === digest(bytes),
    "VALIDATION_FAILED",
    "Candidate does not match the recorded engine output.",
  );
  const evidence = packageEvidence(
    readFileSync(task.original),
    bytes,
    task.maxExpandedBytes,
  );
  return {
    passed: true,
    revision,
    profile: task.operations[0].preservationProfile,
    checks: [
      "recorded-engine-output-hash-chain",
      ...evidence.checks,
      "edit-postconditions-checked-before-staging",
    ],
    evidence,
    notVerified: [
      "visual-layout",
      "XML-semantic-equivalence-of-unedited-content",
      "embedded-object-behavior",
      "cross-application-fidelity",
    ],
    evaluation: task.operations.some((op) => op.engineEvidence.evaluation),
  };
}
