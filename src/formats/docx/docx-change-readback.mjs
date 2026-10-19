import { clean, textOf, tables } from "./docx-sdk.mjs";
import { fingerprint } from "../../infrastructure/files.mjs";
import { requireThat } from "../../infrastructure/errors.mjs";

export const nodeId = node => node.paragraphIds?.paraId ?? node.id;
export const blockLocator = node => ({ kind: "agent_selector", nodeId: nodeId(node),
  selector: { kind: "nodeId", nodeId: nodeId(node) },
  address: { kind: "block", nodeType: node.kind, nodeId: nodeId(node) } });
export async function bodyModel(doc) {
  const body = (await doc.get({ options: { includeResolved: true } })).body;
  requireThat(Array.isArray(body) && body.length <= 20000, "LIMIT_EXCEEDED", "Body model exceeds the 20000 block workflow limit.");
  return body;
}
export const cellText = cell => (cell.content ?? []).map(textOf).join("\n");

export async function captureBodyChanges(doc) {
  const entries = new Map(); let visited = 0;
  const tableEntries = await tables(doc);
  const tableLocators = new Map(tableEntries.map((entry, index) => [entry.address.nodeId,
    { kind: "table", nodeId: entry.address.nodeId, address: entry.address, index }]));
  const walk = (nodes, path = []) => {
    for (const [index, node] of nodes.entries()) {
      requireThat(++visited <= 20000, "LIMIT_EXCEEDED", "Readback inventory exceeds 20000 blocks.");
      const id = nodeId(node) ?? `unaddressable:${[...path, index].join(":")}`;
      if (node.kind === "table") {
        const locator = tableLocators.get(id);
        entries.set(id, { kind: "table", id, hash: fingerprint(clean(node)), locator,
          rows: node.table.rows.map(row => row.cells.map(cell => ({ text: cellText(cell), hash: fingerprint(clean(cell)) }))) });
      } else if (["paragraph", "heading", "listItem"].includes(node.kind)) {
        entries.set(id, { kind: node.kind, id, hash: fingerprint(clean(node)), text: textOf(node), locator: nodeId(node) ? blockLocator(node) : undefined });
      } else if (node.kind === "sdt") walk(node.sdt?.content ?? [], [...path, index]);
    }
  };
  walk(await bodyModel(doc));
  return entries;
}

export function bodyReadback(before, after, options, preview) {
  const maxItems = options.maxItems, maxChars = options.maxChars;
  const items = []; let total = 0, remaining = maxChars, textTruncated = false;
  const text = value => {
    const clipped = value.slice(0, remaining); remaining -= clipped.length;
    if (clipped.length < value.length) textTruncated = true;
    return { text: clipped, truncated: clipped.length < value.length };
  };
  function add(old, current, extra = {}) {
    total++;
    if (items.length >= maxItems) return;
    const change = old && current ? "modified" : old ? "deleted" : "inserted";
    const entry = current ?? old;
    items.push({ kind: entry.kind, change, ...extra,
      ...(old ? { before: text(old.text ?? "") } : {}),
      ...(current ? { after: text(current.text ?? "") } : {}),
      ...(!preview && current?.locator ? { nativeTarget: current.locator,
        targetKind: current.locator.kind === "table" ? "table" : "block" } : {}) });
  }
  for (const id of new Set([...before.keys(), ...after.keys()])) {
    const old = before.get(id), current = after.get(id);
    if (old?.hash === current?.hash) continue;
    if (old?.kind === "table" && current?.kind === "table" && old.rows.length === current.rows.length && old.rows.every((row, i) => row.length === current.rows[i].length)) {
      const previousTotal = total;
      old.rows.forEach((row, r) => row.forEach((cell, c) => {
        if (cell.hash !== current.rows[r][c].hash) add(
          { ...cell, kind: "tableCell" }, { ...current.rows[r][c], kind: "tableCell", locator: current.locator }, { row: r, column: c });
      }));
      if (total > previousTotal) continue;
    }
    const summarize = entry => entry && ({ ...entry, text: entry.text ?? entry.rows.map(row => row.map(c => c.text).join("\t")).join("\n") });
    add(summarize(old), summarize(current));
  }
  return { scope: "body", items, total, truncated: total > items.length || textTruncated,
    textChars: maxChars - remaining, reusableTargets: !preview,
    notChecked: ["formatting", "images", "unsupported-body-objects", "headers-footers", "notes", "review-metadata", "nested-tables", "visual-layout"],
    nextAction: "Use targets only with the returned view. Targets address whole blocks or parent tables, not matched text spans; retain row/column for table cells. Read other scopes explicitly." };
}
