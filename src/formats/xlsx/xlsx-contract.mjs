import { z } from 'zod';
import { resultSchemas } from '../../contracts/results.mjs';

// The request surface also serves CSV. Intrinsic Excel text/grid limits are
// enforced by the XLSX validator, not imposed on CSV's literal fields.
const scalar = z.union([z.string(), z.number().finite(), z.boolean(), z.null()]);
const sheet = z.string().min(1).max(31);
const address = z.string().regex(/^[A-Z]+[1-9]\d*$/);
const range = z.string().regex(/^[A-Z]+[1-9]\d*(?::[A-Z]+[1-9]\d*)?$/);
const format = {
  bold: z.boolean().optional(), italic: z.boolean().optional(),
  underline: z.boolean().optional(), fontFamily: z.string().min(1).max(100).optional(),
  fontSize: z.number().min(6).max(72).optional(),
  fontColor: z.string().regex(/^[0-9A-Fa-f]{6}$/).optional(),
  fillColor: z.string().regex(/^[0-9A-Fa-f]{6}$/).optional(),
  horizontalAlign: z.enum(['left', 'center', 'right']).optional(),
  verticalAlign: z.enum(['top', 'center', 'bottom']).optional(),
  wrapText: z.boolean().optional(), border: z.enum(['none', 'all']).optional(),
  styleId: z.number().int().min(0).optional(),
  numberFormat: z.enum(['general', 'number', 'percent', 'date', 'text', 'currency', 'accounting', 'thousands', 'decimal0', 'decimal1', 'decimal2']).optional(),
};
const cellFields = { value: scalar.optional(), formula: z.string().min(1).max(8192).optional(),
  expectedValue: scalar.optional(), expectedFormula: z.string().max(8192).nullable().optional(),
  expectedRawValue: z.string().nullable().optional(), ...format };
const view = z.object({ documentId: z.string().min(1), candidateId: z.string().optional(), revision: z.string().regex(/^sha256:[a-f0-9]{64}$/) }).strict();
const target = z.object({ documentId: z.string().min(1), candidateId: z.string().optional(), sheet, cell: address, revision: z.string().regex(/^sha256:[a-f0-9]{64}$/) }).strict();
export const workbookCommand = z.union([
  z.object({ type: z.literal('recalculate') }).strict(),
  z.object({ target, ...cellFields }).strict(),
  z.object({ sheet, cell: address, ...cellFields }).strict(),
  z.object({ type: z.literal('setCells'), sheet, cells: z.array(z.object({ cell: address, ...cellFields }).strict()).min(1).max(1000) }).strict(),
  z.object({ type: z.literal('formatRange'), sheet, range, format: z.object(format).strict() }).strict(),
  z.object({ type: z.literal('clearRange'), sheet, range }).strict(),
  z.object({ type: z.literal('clearHyperlinks'), sheet, ranges: z.array(range).min(1).max(10000) }).strict(),
  z.object({ type: z.literal('pasteValues'), sheet, sourceSheet: sheet.optional(), source: range, range }).strict(),
  z.object({ type: z.literal('fillSeries'), sheet, range }).strict(),
  z.object({ type: z.literal('fillRange'), sheet, source: range, range }).strict(),
  z.object({ type: z.literal('fillAcrossSheets'), sheet, sourceSheet: sheet, range, mode: z.enum(['all','contents','formats']) }).strict(),
  z.object({ type: z.literal('pasteFormulas'), sheet, source: range, range }).strict(),
  z.object({ sheet, operation: z.literal('insertRow'), index: z.number().int().min(0).max(1048574), count: z.number().int().min(1).max(1048576).optional() }).strict(),
  z.object({ sheet, operation: z.literal('deleteRow'), index: z.number().int().min(0).max(1048575), count: z.number().int().min(1).max(1048576).optional() }).strict(),
  z.object({ sheet, operation: z.enum(['insertColumn', 'deleteColumn']), index: z.number().int().min(0).max(16383), count: z.number().int().min(1).max(16384).optional() }).strict(),
  z.object({ sheet, operation: z.literal('sortRows'), range, column: z.string().regex(/^[A-Z]+$/), direction: z.enum(['asc', 'desc']) }).strict(),
  z.object({ sheet, operation: z.literal('filterRows'), range, column: z.string().regex(/^[A-Z]+$/), values: z.array(scalar).max(100) }).strict(),
  z.object({ sheet, operation: z.literal('removeFilter') }).strict(),
  z.object({ sheet, operation: z.literal('resizeColumn'), index: z.number().int().min(0).max(16383), width: z.number().min(8).max(100) }).strict(),
  z.object({ sheet, operation: z.literal('resizeRow'), index: z.number().int().min(0).max(1048575), height: z.number().min(8).max(200) }).strict(),
  z.object({ sheet, operation: z.literal('freeze'), rows: z.number().int().min(0).max(100), columns: z.number().int().min(0).max(30) }).strict(),
  z.object({ sheet, operation: z.literal('addSheet') }).strict(),
  z.object({ sheet, operation: z.enum(['renameSheet', 'duplicateSheet']), newName: sheet }).strict(),
  z.object({ sheet, operation: z.enum(['mergeCells', 'unmergeCells']), range }).strict(),
  z.object({ sheet, operation: z.literal('addConditionalRule'), range,
    rule: z.enum(['greaterThan', 'lessThan', 'duplicateValues']), threshold: z.number().finite().optional(),
    fillColor: z.string().regex(/^[0-9A-Fa-f]{6}$/) }).strict(),
  z.object({ sheet, operation: z.literal('createTable'), range, tableName: z.string().min(1).max(255) }).strict(),
  z.object({ type: z.literal('updateRows'), sheet, range, keyColumn: z.string().regex(/^[A-Z]+$/),
    updates: z.array(z.object({ key: scalar, cells: z.array(z.object({ column: z.string().regex(/^[A-Z]+$/), ...cellFields }).strict()).min(1).max(100) }).strict()).min(1).max(100) }).strict(),
]);
export const xlsxTools = {
  xlsx_query: {
    description: 'Read a bounded XLSX or CSV range with reusable sheet/cell targets, literal values, stored raw values, formulas, caches and current calculation status. Set summary:true with sheet/range for native numeric statistics (no cell projection); keys and CSV summaries are not supported. Set search:{text,mode,after?} for bounded XLSX search; mode is display, raw or formula. Continue nextCursor until null; the match limit never drops results. Search cannot combine with keys, summary or offset. Omit sheet for inventory. CSV uses sheet CSV and literal text fields (including empty strings); it has no named tables, formulas or formatting. Example: {view,sheet:"Pricing",range:"A1:F20"}. XLSX uses local open-source calculation; unsupported dependencies remain unresolved.',
    input: z.object({ view, sheet: sheet.optional(), table: z.string().min(1).max(255).optional(), range: range.optional(), summary: z.boolean().optional(),
      search: z.object({text:z.string().min(1).max(32768),mode:z.enum(['display','raw','formula']).default('display'),after:address.optional()}).strict().optional(),
      keys: z.object({ column: z.string().regex(/^[A-Z]+$/), values: z.array(scalar).min(1).max(100) }).strict().optional(),
      offset: z.number().int().min(0).default(0), limit: z.number().int().min(1).max(500).default(100) }),
    output: resultSchemas.generic, risk: 'read', handler: 'xlsx.query',
    execute: (service, name, args, signal) => service.commands.queryWorkbook(name, args, signal),
    engines: ['univer-oss'], presets: ['all', 'recommended'], execution: 'taskEligible',
  },
  xlsx_apply: {
    description: 'Apply one atomic batch to XLSX or CSV through the shared preserving service. Use returned view and a unique operationId; reused IDs recover the same result. XLSX supports cells/formulas, formatting, qualified budget rows, fill and sorting/filtering. Clear Hyperlinks uses canonical disjoint ranges and requires the native XLSX helper; it preserves cell content and formatting. CSV supports literal cells/setCells and updateRows by unique business keys with expected old values; values remain text and formula/format/structure commands are refused. CSV example: {sheet:"CSV",cell:"F2",expectedValue:"125",value:"150"}. Large receipt values may be omitted with explicit lengths/hashes; file contents remain complete. Failures publish nothing. Read the new view, then office_save to a new file with the original extension.',
    input: z.object({ view, operationId: z.string().min(1).max(150), commands: z.array(workbookCommand).min(1).max(1000) }),
    output: resultSchemas.mutation, risk: 'mutate', handler: 'xlsx.apply',
    execute: (service, name, args, signal) => service.commands.applyWorkbook(name, args, signal),
    engines: ['univer-oss'], presets: ['all', 'recommended'], execution: 'taskEligible',
  },
};
