import { readFileSync } from "node:fs";
import { unpack, parseXml } from "../formats/opc/ooxml.mjs";
import { editPackage } from "../formats/opc/package-edits.mjs";
import { readPackage } from "../formats/opc/package-readers.mjs";
import { requireThat } from "../infrastructure/errors.mjs";
import { digest } from "../infrastructure/files.mjs";
import { aspose } from "../engines/aspose-adapter.mjs";
import { validateEngine } from "../formats/opc/engine-evidence.mjs";
import { editPdf } from "../formats/pdf/pdf-edits.mjs";
import { createDocument } from "../formats/common/document-create.mjs";
import { readDocument } from "../formats/common/document-read.mjs";
import { readPdf, editPdfDocument } from "../formats/pdf/pdf-document.mjs";
import { editForms } from "../formats/pdf/pdf-forms.mjs";
import { editPresentation } from "../formats/presentation/presentation-edit.mjs";
import { runPdfWorker } from "../formats/pdf/pdf-worker-adapter.mjs";
import { exportDocumentPdf } from "../formats/common/document-export.mjs";
import { extractDocument } from "../formats/common/extract.mjs";
import { runCsvWorkbook } from "../formats/xlsx/csv-worker.mjs";
import { readCsvWorkbook } from "#document-core/xlsx/csv-document.mjs";
import {readNativeRawWorkbook} from "../formats/xlsx/raw-native.mjs";
import {nativeWorkbookBinary} from "../formats/xlsx/native-binary.mjs";
import { runObjectWorker } from "../formats/docx/docx-object-worker.mjs";
import { runXlsxWorkbook } from "../formats/xlsx/xlsx-worker.mjs";

function adapter(id, kinds, run) {
  return Object.freeze({ id, kinds, run });
}

const adapters = [
  adapter("docx.objects", ["docx_opc_agent_apply", "docx_opc_agent_preview", "docx_opc_validate", "docx_opc_agent_read", "docx_opc_agent_query"], runObjectWorker),
  adapter("csv.oss", ["csv_oss_read", "csv_oss_apply", "csv_validate"], runCsvWorkbook),
  adapter("xlsx.oss", ["xlsx_oss_read", "xlsx_oss_apply"], runXlsxWorkbook),
  adapter("common.create", ["office_create"], createDocument),
  adapter("common.export", ["office_export_pdf"], exportDocumentPdf),
  adapter("engine.status", ["office_engine_status"], (task) => aspose(task)),
  adapter("common.extract", ["office_extract"], extractDocument),
  adapter("common.read", ["office_read"], (task, bytes) =>
    task.format === "pdf" ? readPdf(task, bytes) : task.format === "csv" ? runCsvWorkbook(task, bytes) : readDocument(task, bytes)),
  adapter("pdf.forms", ["pdf_fill_form"], editForms),
  adapter("presentation.structure", [
    "pptx_reorder_slides", "pptx_move_shape", "pptx_replace_frame", "pptx_edit_notes",
  ], editPresentation),
  adapter("pdf.document", ["pdf_rearrange_pages"], editPdfDocument),
  adapter("common.created-validation", ["office_validate_created"], async (task, bytes) => {
    requireThat(
      digest(bytes) === digest(readFileSync(task.original)),
      "VALIDATION_FAILED",
      "Created baseline changed.",
    );
    if (task.format === "csv") readCsvWorkbook(bytes);
    else if (task.format !== "pdf") {
      const packageParts = unpack(bytes, task.maxExpandedBytes);
      for (const [name, part] of Object.entries(packageParts))
        if (/\.(xml|rels)$/.test(name))
          parseXml(Buffer.from(part).toString("utf8"));
    } else
      await readPdf(
        { ...task, args: { scope: "overview", offset: 0, limit: 1 } },
        bytes,
      );
    return {
      passed: true,
      revision: digest(bytes),
      profile: "office.creation.v1",
      checks: ["created-baseline-hash", "document-reopened"],
      notVerified: ["visual-layout", "formula-calculation"],
    };
  }),
  adapter("engine.validation", [], validateEngine),
  adapter("aspose", [
    "xlsx_engine_read", "xlsx_formula", "xlsx_worksheet", "xlsx_rows_columns",
    "pptx_engine_read", "xlsx_edit_range", "xlsx_recalculate",
    "pptx_edit_shape_text", "pptx_set_notes",
  ], aspose),
  adapter("pdf.package-edit", ["pdf_edit_pages", "pdf_validate"], editPdf),
  adapter("opc.edit", [
    "xlsx_write_cells", "xlsx_set_formula", "pptx_replace_text", "xlsx_validate", "pptx_validate",
  ], editPackage),
  adapter("opc.read", [
    "xlsx_inspect", "xlsx_read", "xlsx_search", "pptx_inspect", "pptx_read", "pptx_search",
    "office_package_parts", "office_package_extract",
  ], readPackage),
  adapter("pdf.pdfjs", ["pdf_inspect", "pdf_read", "pdf_search", "pdf_outline"], runPdfWorker),
];

const byId = new Map(adapters.map((item) => [item.id, item]));
const byKind = new Map();
for (const item of adapters)
  for (const kind of item.kinds) {
    requireThat(!byKind.has(kind), "INTERNAL_ERROR", `Duplicate worker operation: ${kind}.`);
    byKind.set(kind, item);
  }
const validationProfiles = new Map([
  ["xlsx_validate:xlsx.oss.v1", "xlsx.oss"],
  ["xlsx_validate:xlsx.native.v1", "xlsx.oss"],
  ["pdf_validate:pdf.forms.v1", "pdf.forms"],
  ["pdf_validate:pdf.mupdf.v1", "pdf.document"],
  ["pptx_validate:pptx.structure.v1", "presentation.structure"],
]);

export function formatAdapterId(task) {
  if (["xlsx_validate", "pptx_validate", "pdf_validate"].includes(task.kind)) {
    const operation = task.operations?.[0];
    const profileAdapter = validationProfiles.get(`${task.kind}:${operation?.preservationProfile}`);
    if (profileAdapter) return profileAdapter;
    if (operation?.engineEvidence) return "engine.validation";
  }
  return byKind.get(task.kind)?.id ?? null;
}

export async function runFormatAdapter(task) {
  const selected = byId.get(formatAdapterId(task));
  if (!selected) return { handled: false };
  const indexedXlsx=task.input&&nativeWorkbookBinary()&&(['xlsx_oss_read'].includes(task.kind)||task.kind==='xlsx_oss_apply'&&task.args.preservationProfile==='xlsx.native.v1'||task.kind==='xlsx_validate'&&task.operations?.[0]?.preservationProfile==='xlsx.native.v1'||task.kind==='office_read'&&task.format==='xlsx');
  const bytes = task.input&&!indexedXlsx&&selected.id!=="docx.objects" ? readFileSync(task.input) : undefined;
  return {
    handled: true,
    adapter: selected.id,
    result: indexedXlsx&&task.kind==="office_read" ? await readNativeRawWorkbook(task) : await selected.run(task, bytes),
  };
}

export const formatAdapterIds = Object.freeze(adapters.map(({ id }) => id));
