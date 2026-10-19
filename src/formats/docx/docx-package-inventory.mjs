import { openFilePackage } from "../opc/file-package.mjs";
import { complexObjectInventory } from "./docx-complex-inventory.mjs";
import { parseXml } from "../opc/ooxml.mjs";

function elements(document, localName) {
  return document
    ? Array.from(document.getElementsByTagNameNS("*", localName))
    : [];
}

function attribute(node, localName) {
  return Array.from(node.attributes ?? []).find(
    (item) => item.localName === localName,
  )?.value;
}

function documentFor(parts, name) {
  const bytes = parts[name];
  return bytes ? parseXml(Buffer.from(bytes).toString("utf8")) : null;
}

function visibleText(node) {
  return Array.from(node.getElementsByTagNameNS("*", "t"))
    .map((item) => item.textContent ?? "")
    .join("");
}

/** Read-only package inventory. It never supplies mutation addresses. */
export function inspectDocxPackage(file, maxExpandedBytes) {
  const archive = openFilePackage(file, maxExpandedBytes, maxExpandedBytes);
  const parts = archive.parts;
  try {
  const xmlParts = Object.keys(parts).filter(
    (name) => name.startsWith("word/") && /\.xml$/i.test(name),
  );
  const parsed = xmlParts.map((name) => [name, documentFor(parts, name)]);
  const endnotesDocument = documentFor(parts, "word/endnotes.xml");
  const endnotes = endnotesDocument
    ? elements(endnotesDocument, "endnote")
        .map((node) => ({ id: attribute(node, "id"), text: visibleText(node) }))
        .filter((item) => item.id && Number(item.id) >= 0)
    : [];
  const footnotesDocument = documentFor(parts, "word/footnotes.xml");
  const footnotes = footnotesDocument
    ? elements(footnotesDocument, "footnote")
        .map((node) => ({ id: attribute(node, "id"), text: visibleText(node) }))
        .filter((item) => item.id && Number(item.id) >= 0)
    : [];
  const textboxes = parsed.flatMap(([partName, document]) =>
    elements(document, "txbxContent").map((node, index) => {
      const text = visibleText(node);
      return {
        id: `${partName}#textbox-${index + 1}`,
        partName,
        textPreview: text.slice(0, 4000),
        textLength: text.length,
        truncated: text.length > 4000,
        readable: false,
        inventoryReadable: true,
        mutationTargetAvailable: false,
        contentSource: "read-only-ooxml-package-inventory",
        reason:
          "Text is available from bounded read-only OOXML inventory, but SuperDoc does not expose a stable textbox story ID for mutation.",
      };
    }),
  );
  const controls = parsed.flatMap(([partName, document]) =>
    elements(document, "sdt").map((node, index) => {
      const properties = elements(node, "sdtPr")[0];
      return {
        id: attribute(elements(properties, "id")[0] ?? {}, "val") ?? `${partName}#sdt-${index + 1}`,
        partName,
        tag: attribute(elements(properties, "tag")[0] ?? {}, "val"),
        title: attribute(elements(properties, "alias")[0] ?? {}, "val"),
        text: visibleText(node),
        repeating: elements(properties, "repeatingSection").length > 0,
      };
    }),
  );
  const bodyText = parsed
    .flatMap(([, document]) => elements(document, "t"))
    .map((node) => node.textContent ?? "")
    .join("\n");
  const literalPlaceholders = [
    ...new Set(
      bodyText.match(/\[(?:insert|enter|fill|select|choose)[^\]\r\n]{0,120}\]/gi) ?? [],
    ),
  ].slice(0, 500);
  const objectInventory = [
    {
      kind: "chart",
      parts: Object.keys(parts).filter((name) => name.startsWith("word/charts/")),
      count: Object.keys(parts).filter((name) => name.startsWith("word/charts/")).length,
      support: "inventory-only",
    },
    {
      kind: "smartArt",
      parts: Object.keys(parts).filter((name) => name.startsWith("word/diagrams/")),
      count: Object.keys(parts).filter((name) => name.startsWith("word/diagrams/")).length,
      support: "inventory-only",
    },
    {
      kind: "embeddedObject",
      parts: Object.keys(parts).filter((name) => name.startsWith("word/embeddings/")),
      count: Object.keys(parts).filter((name) => name.startsWith("word/embeddings/")).length,
      support: "inventory-only",
    },
    {
      kind: "equation",
      count: parsed.reduce(
        (count, [, document]) => count + elements(document, "oMath").length,
        0,
      ),
      support: "inventory-only",
    },
    {
      kind: "textbox",
      count: textboxes.length,
      parts: [...new Set(textboxes.map((item) => item.partName))],
      support: "inventory-only",
    },
  ];
  return {
    storyInventory: [
      { type: "body", readable: true },
      ...footnotes.map((item) => ({
        type: "footnote",
        id: item.id,
        readable: true,
        story: { type: "footnote", id: item.id },
      })),
      ...endnotes.map((item) => ({
        type: "endnote",
        id: item.id,
        readable: true,
        story: { type: "endnote", id: item.id },
      })),
      ...textboxes.map((item) => ({
        type: "textbox",
        id: item.id,
        readable: false,
        inventoryReadable: true,
        mutationTargetAvailable: false,
        textPreview: item.textPreview,
        textLength: item.textLength,
        truncated: item.truncated,
        contentSource: item.contentSource,
        partName: item.partName,
        reason: item.reason,
      })),
    ],
    objectInventory,
    complexObjects: complexObjectInventory(parsed, Object.keys(parts)),
    textboxInventory: textboxes,
    templateReport: {
      literalPlaceholders,
      emptyContentControls: controls
        .filter((item) => item.text.trim().length === 0)
        .map(({ id, tag, title, partName }) => ({ id, tag, title, partName })),
      repeatingSections: controls
        .filter((item) => item.repeating)
        .map(({ id, tag, title, partName }) => ({ id, tag, title, partName })),
      missingDataCount:
        literalPlaceholders.length +
        controls.filter((item) => item.text.trim().length === 0).length,
    },
  };
  } finally { archive.close(); }
}
