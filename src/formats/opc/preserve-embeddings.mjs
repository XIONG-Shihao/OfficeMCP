import { zipSync } from "fflate";
import { decodePackage, descendants, xmlPart, relationships, NS } from "./ooxml.mjs";
import { workbook } from "./package-readers.mjs";
import { requireThat } from "../../infrastructure/errors.mjs";

function objects(parts) {
  const result = new Map();
  for (const sheet of workbook(parts).sheets) {
    const rels = relationships(parts, sheet.partName);
    for (const node of descendants(
      xmlPart(parts, sheet.partName),
      NS.x,
      "oleObject",
    )) {
      const rel = rels.find((r) => r.id === node.getAttributeNS(NS.r, "id"));
      if (!rel || rel.external) continue;
      const key = JSON.stringify([
        sheet.name,
        node.getAttribute("shapeId"),
        node.getAttribute("progId"),
      ]);
      // OOXML AlternateContent may repeat the same logical object.
      requireThat(
        !result.has(key) || result.get(key) === rel.part,
        "PRESERVATION_FAILED",
        "Ambiguous embedded object identity.",
      );
      result.set(key, rel.part);
    }
  }
  return result;
}
export function preserveWorkbookEmbeddings(beforeBytes, afterBytes, limit) {
  const beforePackage = decodePackage(beforeBytes, limit),
    afterPackage = decodePackage(afterBytes, limit),
    before = beforePackage.parts,
    after = afterPackage.parts;
  const oldObjects = objects(before),
    newObjects = objects(after);
  requireThat(
    oldObjects.size === newObjects.size,
    "PRESERVATION_FAILED",
    "Embedded object inventory changed.",
  );
  for (const [identity, source] of oldObjects) {
    const target = newObjects.get(identity);
    requireThat(
      target && before[source] && after[target],
      "PRESERVATION_FAILED",
      "Cannot bind an original embedded object to its saved worksheet and shape.",
    );
    // Cell edits do not edit embedded documents. Aspose rewrites their OLE
    // containers; retain the original bytes at the corresponding relationship.
    after[target] = before[source];
  }
  afterPackage.bytes = oldObjects.size
      ? Buffer.from(zipSync(after, { level: 6, mtime: new Date(1980, 0, 1) }))
      : afterPackage.bytes;
  return {
    bytes: afterPackage.bytes,
    packageContext: afterPackage,
    restoredEmbeddings: oldObjects.size,
  };
}
