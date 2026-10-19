import { readFileSync, writeFileSync } from "node:fs";
import { XMLSerializer } from "@xmldom/xmldom";
import { zipSync, strToU8 } from "fflate";
import {
  unpack,
  xmlPart,
  descendants,
  children,
  child,
  NS,
  relationships,
} from "../opc/ooxml.mjs";
import { presentation } from "../opc/package-readers.mjs";
import { digest } from "../../infrastructure/files.mjs";
import { requireThat } from "../../infrastructure/errors.mjs";
function apply(parts, a, kind) {
  const deck = presentation(parts);
  let partName, doc;
  if (kind === "pptx_reorder_slides") {
    requireThat(
      a.slides.length === deck.slides.length &&
        new Set(a.slides).size === a.slides.length &&
        a.slides.every((i) => i >= 1 && i <= deck.slides.length),
      "INVALID_ARGUMENT",
      "Provide every slide exactly once; this operation reorders without deleting hidden content.",
    );
    partName = "ppt/presentation.xml";
    doc = xmlPart(parts, partName);
    const list = descendants(doc, NS.p, "sldIdLst")[0],
      ids = children(list, NS.p, "sldId");
    for (const i of a.slides) list.appendChild(ids[i - 1]);
  } else if (["pptx_replace_frame", "pptx_edit_notes"].includes(kind)) {
    requireThat(
      deck.slides[a.slide - 1],
      "TARGET_CONFLICT",
      "Slide does not exist.",
    );
    partName = deck.slides[a.slide - 1].partName;
    if (kind === "pptx_edit_notes") {
      const rel = relationships(parts, partName).find(
        (r) => r.type === NS.r + "/notesSlide",
      );
      requireThat(
        rel?.available,
        "UNSUPPORTED_FEATURE",
        "Slide has no notes part. Add notes during creation or use the licensed notes engine.",
      );
      partName = rel.part;
    }
    doc = xmlPart(parts, partName);
    const shape =
      kind === "pptx_edit_notes"
        ? descendants(doc, NS.p, "sp").find((s) =>
            descendants(s, NS.p, "ph").some(
              (p) => p.getAttribute("type") === "body",
            ),
          )
        : descendants(doc, NS.p, "cNvPr").find(
            (n) => n.getAttribute("id") === a.shapeId,
          )?.parentNode.parentNode;
    const frame = child(shape, NS.p, "txBody");
    requireThat(frame, "TARGET_CONFLICT", "Text frame is unavailable.");
    requireThat(
      !descendants(frame, NS.a, "fld").length,
      "UNSUPPORTED_FEATURE",
      "Text frame contains fields.",
    );
    const ps = children(frame, NS.a, "p"),
      old = ps
        .map((p) =>
          children(p, NS.a)
            .map((n) =>
              n.localName === "br"
                ? "\n"
                : descendants(n, NS.a, "t")
                    .map((t) => t.textContent)
                    .join(""),
            )
            .join(""),
        )
        .join("\n");
    requireThat(
      old === a.expectedText,
      "TARGET_CONFLICT",
      "Text frame differs from expected text; read the slide again.",
    );
    const style = descendants(ps[0], NS.a, "rPr")[0],
      pPr = child(ps[0], NS.a, "pPr"),
      replacement = kind === "pptx_edit_notes" ? a.text : a.replacement;
    requireThat(
      !/[\x00-\x08\x0b\x0c\x0e-\x1f]/.test(replacement),
      "INVALID_ARGUMENT",
      "Text contains invalid XML characters.",
    );
    for (const p of ps) frame.removeChild(p);
    for (const line of replacement.split("\n")) {
      const p = doc.createElementNS(NS.a, "a:p"),
        r = doc.createElementNS(NS.a, "a:r"),
        t = doc.createElementNS(NS.a, "a:t");
      if (pPr) p.appendChild(pPr.cloneNode(true));
      if (style) r.appendChild(style.cloneNode(true));
      t.textContent = line;
      r.appendChild(t);
      p.appendChild(r);
      frame.appendChild(p);
    }
  } else {
    requireThat(
      deck.slides[a.slide - 1],
      "TARGET_CONFLICT",
      "Slide does not exist.",
    );
    partName = deck.slides[a.slide - 1].partName;
    doc = xmlPart(parts, partName);
    const id = descendants(doc, NS.p, "cNvPr").find(
      (n) => n.getAttribute("id") === a.shapeId,
    );
    requireThat(id, "TARGET_CONFLICT", "Shape ID does not exist.");
    const shape = id.parentNode.parentNode;
    let xfrm = descendants(shape, NS.a, "xfrm")[0];
    requireThat(
      xfrm,
      "UNSUPPORTED_FEATURE",
      "Shape inherits its transform; explicit positioning is unavailable for this target.",
    );
    const off = child(xfrm, NS.a, "off"),
      ext = child(xfrm, NS.a, "ext");
    requireThat(
      off && ext,
      "UNSUPPORTED_FEATURE",
      "Shape transform is incomplete.",
    );
    off.setAttribute("x", String(Math.round(a.x * 12700)));
    off.setAttribute("y", String(Math.round(a.y * 12700)));
    ext.setAttribute("cx", String(Math.round(a.width * 12700)));
    ext.setAttribute("cy", String(Math.round(a.height * 12700)));
  }
  return {
    ...parts,
    [partName]: strToU8(new XMLSerializer().serializeToString(doc)),
  };
}
export function editPresentation(task, bytes) {
  if (task.kind !== "pptx_validate") {
    const out = Buffer.from(
      zipSync(
        apply(unpack(bytes, task.maxExpandedBytes), task.args, task.kind),
        { level: 6, mtime: new Date(1980, 0, 1) },
      ),
    );
    writeFileSync(task.output, out, { flag: "wx", mode: 0o600 });
    return {
      revision: digest(out),
      affectedTargets: task.args.slides ?? [task.args.shapeId],
    };
  }
  let expected = unpack(readFileSync(task.original), task.maxExpandedBytes);
  for (const op of task.operations) expected = apply(expected, op, op.tool);
  const actual = unpack(bytes, task.maxExpandedBytes);
  requireThat(
    Object.keys(expected).length === Object.keys(actual).length &&
      Object.keys(expected).every(
        (n) => actual[n] && digest(actual[n]) === digest(expected[n]),
      ),
    "VALIDATION_FAILED",
    "Presentation differs from declared structure operations.",
  );
  return {
    passed: true,
    revision: digest(bytes),
    profile: "pptx.structure.v1",
    checks: ["exact-structure-edit-replay", "unrelated-parts-byte-preserved"],
    notVerified: ["visual-overlap", "off-slide-content", "layout-fidelity"],
  };
}
