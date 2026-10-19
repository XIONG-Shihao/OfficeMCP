import { clean, textOf } from "./docx-sdk.mjs";
import { readCoverage, storyLabel, storyOf } from "./docx-story.mjs";

export function canonicalItem(item, index, domain) {
  if (item.itemId && item.provenance && Array.isArray(item.path))
    return { ...item, order: index };
  const kind = item.kind ?? item.type ?? domain ?? "item";
  const {
    nativeTarget: nativeTargetValue,
    children: childItems,
    ...detailSource
  } = item;
  const text =
    item.text ??
    item.textPreview ??
    item.commentText ??
    item.content ??
    item.snippet ??
    item.name ??
    "";
  return {
    kind,
    itemId: String(
      item.nodeId ?? item.entityId ?? item.id ?? item.sdImageId ?? `${domain}-${index}`,
    ),
    story:
      item.story?.storyType ??
      item.story?.type ??
      item.target?.story?.storyType ??
      item.address?.story?.storyType ??
      "body",
    order: index,
    path: item.path ?? [storyLabel(item.story), index],
    ...(item.parentId ? { parentId: item.parentId } : {}),
    text: typeof text === "string" ? text : JSON.stringify(text),
    ...(item.headingLevel !== undefined
      ? { headingLevel: item.headingLevel }
      : {}),
    ...(item.tableContext ? { tableContext: item.tableContext } : {}),
    ...(kind === "comment"
      ? { status: item.status ?? (item.resolved ? "resolved" : "open") }
      : item.status
        ? { status: item.status }
        : {}),
    ...(item.author ? { author: item.author } : {}),
    ...((kind === "trackedChange" ||
      ["trackedChanges", "tracked_changes"].includes(domain)) &&
    (item.changeType ?? item.type)
      ? {
          changeType:
            {
              insertion: "insert",
              deletion: "delete",
              formatting: "format",
            }[item.changeType ?? item.type] ?? (item.changeType ?? item.type),
        }
      : {}),
    ...(item.trackedChanges ? { trackedChanges: item.trackedChanges } : {}),
    ...(item.anchoredText ? { anchoredText: item.anchoredText } : {}),
    ...(childItems ? { children: childItems } : {}),
    formattingSummary: clean(
      item.formatting ?? item.properties?.formatting ?? item.attrs ?? {},
    ),
    provenance:
      (kind === "textboxInventory" || item.contentSource === "read-only-ooxml-package")
        ? {
            engine: "read-only OOXML package inventory",
            adapter: "textbox-inventory",
          }
        : {
            engine: "SuperDoc SDK 2.2.1 / v2",
            adapter: domain,
            ...(item.address ? { addressKind: item.address.kind } : {}),
          },
    ...(kind === "table" && (item.node?.table?.rows ?? item.table?.rows)
      ? {
          rows: (item.node?.table?.rows ?? item.table.rows).map((row, rowIndex) =>
            row.cells.map((cell, columnIndex) => ({
              row: rowIndex,
              column: columnIndex,
              text: cell.content.map((block) => textOf(block)).join("\n"),
            })),
          ),
        }
      : {}),
    ...(kind === "image" || domain === "images"
      ? {
          altText: item.properties?.description ?? item.alt ?? "",
          ...(item.properties?.size ?? {}),
          ...(item.properties?.src
            ? { packagePart: item.properties.src }
            : {}),
        }
      : {}),
    ...(kind === "textboxInventory"
      ? {
          partName: item.partName,
          textLength: item.textLength,
          truncated: item.truncated,
          contentSource: item.contentSource,
          mutationTargetAvailable: false,
        }
      : {}),
    details: clean(detailSource),
    ...(nativeTargetValue ? { nativeTarget: nativeTargetValue } : {}),
  };
}

export function paged(items, args, domain, extra = {}) {
  const offset = args.offset ?? 0;
  const limit = args.limit ?? 30;
  const nextOffset = offset + limit < items.length ? offset + limit : null;
  const coverage = readCoverage(args, domain, nextOffset);
  return {
    items: items
      .slice(offset, offset + limit)
      .map((item, index) => canonicalItem(item, offset + index, domain)),
    total: items.length,
    offset,
    nextOffset,
    representation: args.representation ?? "linear",
    ...extra,
    coverage,
    remainingReads: coverage.remainingReads,
  };
}

export function pagedWindow(items, args, domain, total, nextOffset, extra = {}) {
  const offset = args.offset ?? 0;
  const coverage = readCoverage(args, domain, nextOffset);
  return {
    items: items.map((item, index) =>
      canonicalItem(item, offset + index, domain),
    ),
    total,
    offset,
    nextOffset,
    representation: args.representation ?? "domain",
    ...extra,
    coverage,
    remainingReads: coverage.remainingReads,
  };
}

function structuredChildren(node, path, parentId, story) {
  const direct = Array.isArray(node.content) ? node.content : [];
  const table = node.table?.rows?.flatMap((row) =>
    row.cells.flatMap((cell) => cell.content ?? []),
  );
  return [...direct, ...(table ?? [])].map((child, index) =>
    structuredNode(child, [...path, index], parentId, story),
  );
}

export function structuredNode(node, path, parentId, story) {
  const nodeId =
    node.nodeId ??
    node.address?.nodeId ??
    node.paragraphIds?.paraId ??
    node.id ??
    node.attrs?.sdBlockId ??
    path.join(".");
  const children = structuredChildren(node, path, String(nodeId), story);
  return canonicalItem(
    {
      ...node,
      nodeId,
      story,
      path,
      parentId,
      children,
      text: textOf(node),
      nativeTarget: {
        kind: "agent_selector",
        selector: { kind: "nodeId", nodeId },
        address: {
          kind: "block",
          nodeType: node.kind ?? "paragraph",
          nodeId,
          story: storyOf(story),
        },
      },
    },
    path.at(-1) ?? 0,
    "structured",
  );
}
