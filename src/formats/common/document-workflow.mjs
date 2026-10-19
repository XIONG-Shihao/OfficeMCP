import { requireThat } from "../../infrastructure/errors.mjs";
import { applyWorkbook } from '../xlsx/xlsx-service.mjs';
export const asView = (documentId, revision, candidateId) => ({
  documentId,
  ...(candidateId ? { candidateId } : {}),
  revision,
});

export function createWorkflow(service, args, signal) {
  return service.documents.create(args, signal);
}

export async function editWorkflow(service, a, signal) {
  const d = service.doc(a.view.documentId),
    x = a.action;
  if (['xlsx', 'csv'].includes(d.format) && ['write_cells', 'set_formula', 'recalculate'].includes(x.type)) {
    const formats = { '0.00': 'number', '0': 'general', '0.00%': 'percent', '0%': 'percent', '@': 'text', 'yyyy-mm-dd': 'date' };
    const commands = x.type === 'write_cells' ? x.cells.map(cell => {
      const { address, format, ...content } = cell;
      requireThat(!format?.fontSize && format?.wrapText === undefined, 'UNSUPPORTED_FEATURE', 'Use a qualified optional engine for font-size or wrapping changes.');
      const numberFormat = format?.numberFormat ? formats[format.numberFormat] ?? format.numberFormat : undefined;
      return { sheet: x.sheet, cell: address, ...content,
        ...(format ? { bold: format.bold, italic: format.italic, numberFormat } : {}) };
    }) : x.type === 'set_formula' ? [{ sheet: x.sheet, cell: x.address,
      formula: x.formula, expectedFormula: x.expectedFormula }] : [{ type: 'recalculate' }];
    return applyWorkbook(service, { view: a.view, operationId: a.operationId, commands }, signal);
  }
  const base = {
    documentId: a.view.documentId,
    expectedSourceRevision: d.revision,
    operationId: a.operationId,
    background: false,
    ...(a.view.candidateId
      ? {
          candidateId: a.view.candidateId,
          expectedCandidateRevision: a.view.revision,
        }
      : {}),
  };
  requireThat(
    a.view.candidateId || a.view.revision === d.revision,
    "REVISION_CONFLICT",
    "Read the current view again before editing.",
  );
  const mappings = {
    replace_text: ["docx_replace_span", "docx.story-text.v1"],
    table_cell: ["docx_table_cell", "docx.story-text.v1"],
    table_row: ["docx_table_row", "docx.story-text.v1"],
    write_cells: ["xlsx_edit_range", "xlsx.aspose.v1"],
    set_formula: ["xlsx_formula", "xlsx.aspose.v1"],
    worksheet: ["xlsx_worksheet", "xlsx.aspose.v1"],
    rows_columns: ["xlsx_rows_columns", "xlsx.aspose.v1"],
    replace_shape_text: ["pptx_replace_frame", "pptx.structure.v1"],
    set_notes: ["pptx_edit_notes", "pptx.structure.v1"],
    reorder_slides: ["pptx_reorder_slides", "pptx.structure.v1"],
    move_shape: ["pptx_move_shape", "pptx.structure.v1"],
    recalculate: ["xlsx_recalculate", "xlsx.aspose.v1"],
    pages: ["pdf_rearrange_pages", "pdf.mupdf.v1"],
    fill_form: ["pdf_fill_form", "pdf.forms.v1"],
  };
  mappings.format_text = ["docx_format_span", "docx.story-text.v1"];
  mappings.review_changes = ["docx_review_changes", "docx.story-text.v1"];
  const [tool, profile] = mappings[x.type];
  const { type, ...args } = x;
  const result = await service.operation(
    tool,
    {
      ...base,
      ...args,
      preservationProfile: profile,
    },
    signal,
    { transition: true },
  );
  return {
    ...result,
    view: asView(
      d.id ?? a.view.documentId,
      result.data.candidateRevision,
      result.data.candidateId,
    ),
    data: {
      ...result.data,
      summary: {
        action: type,
        targets: result.data.affectedTargets,
        originalUnchanged: true,
      },
      nextAction: "Read the returned view or call office_save.",
    },
  };
}

export function saveWorkflow(service, args, signal, requestName = "office_save") {
  return service.documents.save(args, signal, requestName);
}
