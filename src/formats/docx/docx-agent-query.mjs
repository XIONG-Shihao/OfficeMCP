import { readDocxDomain } from "./docx-domain-items.mjs";
import { replaceSelection } from "./docx-text-actions.mjs";
import { pagedWindow } from "./docx-canonical.mjs";
import { storyOf } from "./docx-story.mjs";
import { requireThat } from "../../infrastructure/errors.mjs";
import { descendantBlockIds, paragraphs, tables, textOf } from "./docx-sdk.mjs";

function comparable(value, caseSensitive) {
  const text = String(value ?? "");
  return caseSensitive ? text : text.toLocaleLowerCase();
}

function entityField(item, name) {
  if (name === "id") return item.id ?? item.entityId ?? item.noteId;
  if (name === "name") return item.name ?? item.title ?? item.label;
  if (name === "text")
    return item.text ?? item.content ?? item.displayText ?? item.snippet ?? item.excerpt;
  if (name === "tag") return item.tag ?? item.properties?.tag;
  if (name === "author") return item.author ?? item.userName;
  if (name === "changeType")
    return {
      insertion: "insert",
      deletion: "delete",
      formatting: "format",
    }[item.changeType ?? item.type] ?? (item.changeType ?? item.type);
}

async function structuralMatches(doc, select, scope = {}) {
  if (select.type === "entity") {
    // SDK entity inventories have no uniform story/container filter.
    requireThat(
      !scope.withinLocator && !scope.withinTarget && !scope.story,
      "UNSUPPORTED_FEATURE",
      `${select.type} selectors do not support this story or withinTarget scope. Use a scoped text/node query instead.`,
    );
  }
  const selectedStory = scope.story ? storyOf(scope.story)
    : scope.withinLocator?.address?.story ?? storyOf();
  if (["table", "headingPath"].includes(select.type))
    requireThat(selectedStory.storyType === "body", "UNSUPPORTED_FEATURE",
      `${select.type} selectors support the body story only. SuperDoc cannot expose scoped structural models; use a scoped text/node query instead.`);
  const inContainer = async (entries) => {
    if (!scope.withinLocator) return entries;
    const within = scope.withinLocator.address;
    requireThat(within?.kind === "block", "UNSUPPORTED_FEATURE", "Structural withinTarget requires a block container.");
    const ids = await descendantBlockIds(doc, within);
    return entries.filter((entry) => ids.has(entry.address.nodeId));
  };
  if (select.type === "entity") {
    const entries = await readDocxDomain(doc, select.domain);
    const filters = ["id", "name", "text", "tag", "author", "changeType"]
      .filter((name) => select[name] !== undefined);
    return entries
      .map((item, index) => ({ item, index }))
      .filter(({ item }) =>
        filters.every(
          (name) =>
            comparable(entityField(item, name), select.caseSensitive) ===
            comparable(select[name], select.caseSensitive),
        ),
      )
      .map(({ item, index }) => ({
        ...item,
        kind: item.kind ?? select.domain,
        nodeId: item.nativeTarget?.nodeId ?? item.id ?? item.entityId ?? String(index),
        text: entityField(item, "text") ?? entityField(item, "name") ?? "",
        nativeTarget: item.nativeTarget,
      }));
  }
  if (select.type === "headingPath") {
    const stack = [];
    return (await inContainer(await paragraphs(doc)))
      .filter((entry) => entry.node.kind === "heading")
      .flatMap((entry) => {
        const level = Number(
          entry.node.heading?.level ??
            entry.node.attrs?.level ??
            entry.node.level ??
            1,
        );
        stack.length = Math.max(0, level - 1);
        stack[level - 1] = textOf(entry.node);
        const current = stack.slice(0, level);
        const matches =
          current.length === select.path.length &&
          current.every(
            (part, index) =>
              comparable(part, select.caseSensitive) ===
              comparable(select.path[index], select.caseSensitive),
          );
        return matches
          ? [{
              kind: "heading",
              nodeId: entry.address.nodeId,
              text: textOf(entry.node),
              headingPath: current,
              nativeTarget: {
                kind: "agent_selector",
                selector: { kind: "nodeId", nodeId: entry.address.nodeId },
                address: entry.address,
                nodeId: entry.address.nodeId,
              },
            }]
          : [];
      });
  }
  if (select.type === "table") {
    const bodyTables = await tables(doc);
    const entries = await inContainer(bodyTables);
    return entries.flatMap((entry) => {
      const rows = entry.node?.table?.rows ?? [];
      const matrix = rows.map((row) =>
        row.cells.map((cell) => (cell.content ?? []).map(textOf).join("\n")),
      );
      const allText = matrix.flat().join("\n");
      const headerText = (matrix[0] ?? []).join("\n");
      const includes = (actual, expected) =>
        expected === undefined ||
        comparable(actual, select.caseSensitive).includes(
          comparable(expected, select.caseSensitive),
        );
      if (!includes(allText, select.contains) || !includes(headerText, select.header))
        return [];
      return [{
        ...entry,
        kind: "table",
        nodeId: entry.address.nodeId,
        text: allText,
        rows: matrix.map((row, rowIndex) =>
          row.map((text, column) => ({ row: rowIndex, column, text })),
        ),
        nativeTarget: {
          kind: "table",
          address: entry.address,
          nodeId: entry.address.nodeId,
          index: bodyTables.findIndex((item) => item.address.nodeId === entry.address.nodeId),
        },
      }];
    });
  }
  return null;
}

function requiredMatches(items, requireMode, label) {
  if (requireMode === "exactlyOne")
    requireThat(
      items.length === 1,
      "TARGET_CONFLICT",
      `${label} resolved ${items.length} items; exactly one is required.`,
    );
  if (requireMode === "first")
    requireThat(items.length > 0, "NOT_FOUND", `${label} did not match.`);
}

function selectionSegments(target) {
  return (
    target?.segments ??
    (target?.start?.kind === "text" &&
    target.start.blockId === target?.end?.blockId
      ? [
          {
            blockId: target.start.blockId,
            range: { start: target.start.offset, end: target.end.offset },
          },
        ]
      : [])
  );
}

function matchedText(match) {
  return match.matchKind === "text" && match.target?.start?.kind === "text"
    ? (match.blocks ?? [])
        .filter((block) => block.blockId === match.target.start.blockId)
        .map((block) =>
          // Some SDK responses contain the matched range, others the full block.
          // Offsets in the target remain absolute in either representation.
          block.text.slice(
            match.target.start.offset - (block.range && block.text.length === block.range.end - block.range.start ? block.range.start : 0),
            match.target.end.offset - (block.range && block.text.length === block.range.end - block.range.start ? block.range.start : 0),
          ),
        )
        .join("")
    : match.blocks?.map((block) => block.text).join("") ?? match.snippet ?? "";
}

export async function queryAgentDocument(doc, args) {
  const structural = await structuralMatches(doc, args.select, args);
  if (structural) {
    requiredMatches(structural, args.require, "Structural selector");
    const offset = args.offset ?? 0;
    const selected = args.require === "first" ? structural.slice(0, 1) : structural;
    const items = selected.slice(offset, offset + args.limit);
    return pagedWindow(
      items,
      args,
      "query",
      structural.length,
      offset + items.length < selected.length ? offset + items.length : null,
      { matchedTotal: structural.length },
    );
  }
  const result = await doc.query.match({
    select:
      args.select.type === "text"
        ? {
            type: "text",
            pattern: args.select.pattern,
            mode: args.select.mode,
            caseSensitive: args.select.caseSensitive,
            wholeWord: args.select.wholeWord,
          }
        : { type: "node", nodeType: args.select.nodeType },
    in: storyOf(args.story),
    ...(args.withinLocator?.address
      ? { within: args.withinLocator.address }
      : {}),
    // Own cardinality refusals at the product boundary; SDK exactlyOne throws
    // an engine exception before OfficeMCP can return actionable target help.
    require: args.require === 'exactlyOne' ? 'any' : args.require,
    offset: args.offset,
    limit: args.limit,
  });
  if (args.require === 'exactlyOne')
    requireThat(result.total === 1, 'TARGET_CONFLICT',
      `Text/node selector resolved ${result.total} items; exactly one is required.`);
  const items = result.items.map((match) => {
    const text = matchedText(match);
    const runs = (match.blocks ?? []).flatMap((block) =>
      (block.runs ?? [])
        .filter(
          (run) =>
            typeof run.text === "string" && run.text.length > 0 &&
            match.target?.start?.kind === "text" &&
            block.blockId === match.target.start.blockId &&
            run.range.end > match.target.start.offset &&
            run.range.start < match.target.end.offset,
        )
        .map((run) => ({
          text: run.text,
          range: run.range,
          effective: run.styles?.effective,
          direct: run.styles?.direct,
          fontFamily: run.fontFamily,
          fontSizePt: run.fontSizePt,
          color: run.color,
          highlight: run.highlight,
        })),
    );
    const segments = selectionSegments(match.target);
    return {
      kind: match.matchKind ?? match.address?.nodeType ?? "match",
      nodeId: match.id,
      text: text || runs.map((run) => run.text).join(""),
      snippet: match.snippet,
      formatting: runs.length === 1 ? runs[0].effective : { runs },
      runs,
      ...(match.matchKind === "text" && !runs.length ? { formattingUnavailable: true } : {}),
      nativeTarget: {
        kind: "agent_selector",
        selector: { kind: "nodeId", nodeId: match.address?.nodeId },
        ...(match.handle?.ref
          ? { refSelector: { kind: "ref", ref: match.handle.ref } }
          : {}),
        address: match.address,
        selectionTarget: match.target,
        textTarget:
          segments.length > 0 ? { kind: "text", segments } : undefined,
        matchedText: text,
      },
    };
  });
  const offset = args.offset ?? 0;
  return pagedWindow(
    items,
    args,
    "query",
    result.total,
    offset + items.length < result.total ? offset + items.length : null,
    {
      evaluatedRevision: result.evaluatedRevision,
      matchedTotal: result.total,
    },
  );
}

export function queryLocator(match) {
  const segments = selectionSegments(match.target);
  return {
    kind: "agent_selector",
    selector: { kind: "nodeId", nodeId: match.address.nodeId },
    ...(match.handle?.ref
      ? { refSelector: { kind: "ref", ref: match.handle.ref } }
      : {}),
    address: match.address,
    nodeId: match.address.nodeId,
    selectionTarget: match.target,
    textTarget: segments.length > 0 ? { kind: "text", segments } : undefined,
    matchedText: matchedText(match),
  };
}

export async function resolveInlineLocator(doc, step) {
  if (!step.select) return step.targetLocator ?? null;
  const structural = await structuralMatches(doc, step.select, step);
  if (structural) {
    requiredMatches(structural, "exactlyOne", `Inline selector for ${step.action}`);
    return structural[0].nativeTarget;
  }
  const result = await doc.query.match({
    in: storyOf(step.story),
    select:
      step.select.type === "text"
        ? {
            type: "text",
            pattern: step.select.pattern,
            mode: step.select.mode,
            caseSensitive: step.select.caseSensitive,
            wholeWord: step.select.wholeWord,
          }
        : { type: "node", nodeType: step.select.nodeType },
    require: "any",
    offset: 0,
    limit: 2,
  });
  requireThat(
    result.total === 1 && result.items[0]?.address?.nodeId,
    "TARGET_CONFLICT",
    `Inline selector for ${step.action} resolved ${result.total} items; exactly one block is required.`,
  );
  return queryLocator(result.items[0]);
}

export async function directTextReplacement(doc, step, locator) {
  if (
    step.action !== "replace_text" ||
    step.arguments.edits.length !== 1
  )
    return null;
  const edit = step.arguments.edits[0];
  if (!locator?.address) {
    // A model may supply a unique literal without first querying an opaque
    // target. Use the same range-preserving route as a targeted replacement.
    // Leave multi-match/batch behavior to the existing guarded SDK path.
    const selected = await doc.query.match({
      select: { type: 'text', pattern: edit.find, mode: 'contains', caseSensitive: step.arguments.caseSensitive },
      in: storyOf(step.story), require: 'any', offset: 0, limit: 2,
    });
    if (selected.total !== 1 || !selected.items[0]?.address?.nodeId || !selected.items[0]?.target) return null;
    locator = queryLocator(selected.items[0]);
  }
  let target = locator.selectionTarget;
  let text = locator.matchedText;
  const same = step.arguments.caseSensitive
    ? text === edit.find
    : text?.toLocaleLowerCase() === edit.find.toLocaleLowerCase();
  if (!target || !same) {
    const narrowed = await doc.query.match({
      select: {
        type: "text",
        pattern: edit.find,
        mode: "contains",
        caseSensitive: step.arguments.caseSensitive,
      },
      in: locator.address?.story ?? { kind: "story", storyType: "body" },
      within: locator.address,
      require: "any",
      offset: 0,
      limit: 2,
    });
    requireThat(
      narrowed.total === 1 && narrowed.items[0]?.target,
      "TARGET_CONFLICT",
      "The selected block does not contain exactly one requested replacement.",
    );
    target = narrowed.items[0].target;
    text = edit.find;
  }
  const result = await replaceSelection(doc, {
      target,
      expectedText: text,
      text: edit.replace,
      changeMode: step.arguments.changeMode,
    });
  return {
    status: "ok",
    intent: step.action,
    selectedTargets: [locator.address?.nodeId].filter(Boolean),
    executedOperations: [
      {
        operationId: "doc.replace",
        find: text,
        replacement: edit.replace,
        result,
      },
    ],
    verificationPassed: true,
    verification: [{ check: { kind: "sdk-success-receipt" }, passed: true }],
  };
}
