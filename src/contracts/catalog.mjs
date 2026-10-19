import { operationRegistry, VERSION, CONTRACT } from "./registry.mjs";
import { operationManifestEntry } from "./operation-spec.mjs";

export const layerInfo = [
  {
    id: "docx",
    label: "DOCX",
    engine:
      "SuperDoc SDK 2.2.1 / v2 plus guarded OOXML adapters; optional Word/LibreOffice layout",
    summary:
      "14 agent-facing DOCX tools backed by SuperDoc v2, durable candidates, semantic/package validation and durable PDF export.",
    limitations:
      "Optional layout providers supply revision-bound page text targets and explicit Word-reference appearance comparisons. Guarded textbox/equation/chart edits and merged/nested cell text are supported within their declared limits. Full table reconstruction, style definitions and some complex tracked ranges remain incomplete.",
    missing: [
      "Exact Microsoft Word layout certification",
      "Reliable tracked edits in all formatted/table-bearing documents",
      "Page targeting of ambiguous text, cross-page ranges and opaque objects",
      "Unrestricted nested-table reconstruction and arbitrary object authoring",
      "Style and numbering-definition editing",
      "Guaranteed fresh page fields across Word-compatible renderers",
      "Cross-version human-readable diff",
    ],
    proposedEngine:
      "SuperDoc remains the only mutation engine. Independent package checks are read-only evidence; LibreOffice is used only for PDF export.",
  },
  {
    id: "xlsx",
    label: "XLSX / CSV",
    engine:
      "Shared preserving OOXML service · Univer OSS 0.25.2 calculation · ExcelJS 4.4.0 creation · optional legacy Aspose.Cells",
    summary:
      "Read XLSX/CSV ranges and business keys. Apply atomic guarded cells through the same preserving service as the editor. XLSX additionally supports formulas, named tables, formatting, qualified row insertion, fill and sorting/filtering; CSV preserves literal fields and encoding.",
    limitations:
      "Opaque content is preserved. Special/shared formulas and unsafe numeric dependencies remain unresolved; protection and unsafe structural references are refused. No external-data refresh or chart/pivot authoring.",
    missing: [
      "Chart/pivot editing",
      "Full comment/table semantic reading",
      "External data refresh",
    ],
    proposedEngine:
      "Ordinary supported XLSX workflows use local open-source calculation and the preserving writer. Aspose remains an explicit optional legacy route.",
  },
  {
    id: "pptx",
    label: "PPTX",
    engine:
      "PptxGenJS 4.0.1 creation · preserving OOXML edits · optional Aspose.Slides 26.8.1",
    summary:
      "Create slides with text, tables, images and notes. Read shapes and notes; replace text, update existing notes, move shapes and reorder slides without Aspose.",
    limitations:
      "Full-frame replacement adopts the first run style. Inherited geometry is refused. Aspose evaluation truncates loaded text; its advanced text edits require a license.",
    missing: [
      "Existing chart/table/image editing",
      "Adding notes to slides with no notes part",
      "Full SmartArt/master semantics",
    ],
    proposedEngine:
      "PptxGenJS for creation; Aspose.Slides for licensed advanced editing.",
  },
  {
    id: "pdf",
    label: "PDF",
    engine:
      "MuPDF.js 1.28.1 · pdf-lib 1.17.1 · optional Tesseract",
    summary:
      "Create Unicode text PDFs; inspect blocks, outlines, annotations, forms and attachment metadata; extract attachments, fill forms, reorder/rotate pages and optionally OCR.",
    limitations:
      "Page removal/duplication is refused when document-level references need remapping. OCR is fallible; text order is not guaranteed visual reading order.",
    missing: [
      "Annotation editing",
      "Semantic table extraction",
      "General PDF text reflow",
      "Reference-aware page deletion/duplication",
    ],
    proposedEngine:
      "MuPDF for inspection, pdf-lib for guarded mutations. Local Tesseract supplies OCR; its rasterization is private to that operation.",
  },
  {
    id: "office",
    label: "Office",
    engine: "Local stdio · typed workflows · SQLite receipts",
    summary:
      "Open or create, read a revision-bound view, edit and save. Extract assets and return native MCP images. Save validates every editing phase automatically.",
    limitations:
      "Save-as never overwrites. Validation reports its scope; it does not prove exact Office appearance. Optional engines must be configured locally.",
    missing: [
      "Broad cross-platform qualification",
      "Representative visual acceptance corpus",
    ],
    proposedEngine:
      "Standalone process for any stdio MCP host; the app consumes its public catalog.",
  },
];

const evaluations = {
  docx: {
    status: "Current DOCX workflow coverage",
    summary:
      "Open/create, canonical reads, semantic query, guarded actions, validation and save/export are available. Run test:response for semantic/package acceptance; application display testing is separate.",
  },
  xlsx: {
    status: "Current workflow coverage",
    summary:
      "Shared open-source calculation and preserving edits cover bounded queries, guarded/keyed batches, qualified budget rows and save/reopen. Current actions and limits are documented in XLSX_MCP.md. Optional legacy Aspose routes are separate.",
  },
  pptx: {
    status: "Current workflow coverage",
    summary:
      "Creation, whole-frame text, existing notes, shape positioning and save pass without Aspose. Licensed advanced Slides edits remain unqualified.",
  },
  pdf: {
    status: "Current workflow coverage",
    summary:
      "Unicode creation/forms, outline-aware full page permutations and independent duplicate rotations are tested. Optional native OCR remains a separate operation.",
  },
  office: {
    status: "Current workflow coverage",
    summary:
      "OpenCode with saved DeepSeek completed a synthetic guarded contract edit and reported its footer. Direct Python also succeeded; no general performance advantage established.",
  },
};

/** Public discovery selects DOCX-agent specs and complete non-DOCX specs once. */
export function catalog(toolset = "public", version = VERSION, contractVersion = toolset === "public" ? CONTRACT + " / docx/6 agent" : CONTRACT) {
  const operations = toolset === "public"
    ? [...operationRegistry("all").values()].filter((spec) => !spec.name.startsWith("docx_"))
        .concat([...operationRegistry("docx-agent").values()])
    : [...operationRegistry(toolset).values()];
  return {
    version,
    contractVersion,
    transport: "stdio",
    status: "Capabilities vary by format; inspect each operation and its functional limits.",
    layers: layerInfo.map((layer) => ({
      ...layer,
      evaluation: evaluations[layer.id],
      operations: operations
        .filter((spec) => spec.name.startsWith(layer.id + "_"))
        .map((spec) => ({ ...operationManifestEntry(spec), ...spec.discovery })),
    })),
  };
}
