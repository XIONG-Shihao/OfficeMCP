import fs from 'node:fs';
import { readCsvWorkbook, editCsvWorkbook } from '#document-core/xlsx/csv-document.mjs';
import { csvCell, csvCellInfo, csvPosition, csvAddress } from '#document-core/xlsx/csv-workbook.mjs';
const cellPosition = address => { const { row, column } = csvPosition(address); return { row, col: column }; };
import { digest } from '../../infrastructure/files.mjs';
import { OfficeError, requireThat } from '../../infrastructure/errors.mjs';
import { readWindow } from '../common/read-window.mjs';

const PROFILE = 'csv.literals.v1';
const engine = 'Shared preserving CSV literal-table service';
const noCalculation = revision => ({ revision, engine: 'csv-literals', calculated: 0, unresolved: 0, error: 0, circular: 0,
  formulaSupport: false });

function bounds(range) {
  const pieces = String(range).split(':');
  if (pieces.length > 2) throw Error('Invalid CSV range');
  const first = cellPosition(pieces[0]), last = cellPosition(pieces[1] ?? pieces[0]);
  if (first.row > last.row || first.col > last.col) throw Error('Invalid CSV range');
  return { first, last, width: last.col - first.col + 1 };
}

export function readCsvRange(index, args, revision = index.revision) {
  requireThat(!args.table, 'UNSUPPORTED_FEATURE', 'CSV has no named tables; query sheet CSV by range or business key.');
  requireThat(args.sheet === 'CSV', 'TARGET_CONFLICT', 'CSV has one sheet named CSV.');
  const { range = 'A1:J20', offset = 0, limit = 100, keys, query } = args;
  requireThat(Number.isSafeInteger(offset) && offset >= 0 && Number.isInteger(limit) && limit >= 1 && limit <= 500,
    'INVALID_ARGUMENT', 'Invalid CSV read window.');
  const { first, last, width } = bounds(range), missingKeys = [], duplicateKeys = [];
  let selectedRows;
  if (keys) {
    const column = cellPosition(`${keys.column}1`).col;
    requireThat(column >= first.col && column <= last.col, 'INVALID_ARGUMENT', 'Key column is outside the requested range.');
    const wanted = new Map(keys.values.map(value => [value, []]));
    for (let row = first.row; row <= Math.min(last.row, index.rows - 1); row++) {
      const value = csvCell(index, row, column);
      if (wanted.has(value)) wanted.get(value).push(row);
    }
    selectedRows = [...new Set([...wanted.values()].flat())].sort((a, b) => a - b);
    for (const [value, matches] of wanted) {
      if (!matches.length) missingKeys.push(value);
      if (matches.length > 1) duplicateKeys.push(value);
    }
  }
  const count = selectedRows ? selectedRows.length * width : (last.row - first.row + 1) * width;
  requireThat(Number.isSafeInteger(count), 'INVALID_ARGUMENT', 'CSV range exceeds JavaScript integer precision.');
  const itemAtCoordinate = (row, col) => {
    const cell = csvAddress(row, col), info = csvCellInfo(index, row, col);
    return { sheet: 'CSV', cell, address: cell, target: { sheet: 'CSV', cell, revision },
      value: info?.value ?? null, rawValue: info?.rawValue ?? null, dataType: info?.dataType ?? 'blank',
      formula: null, cachedValue: null, calculation: null, expectedValue: info?.value ?? null, expectedFormula: null };
  };
  const itemAt = position => itemAtCoordinate(selectedRows ? selectedRows[Math.floor(position / width)]
    : first.row + Math.floor(position / width), first.col + position % width);
  const items = []; let total = count;
  if (query) {
    const needle = query.toLocaleLowerCase(); total = 0;
    const rowCount = selectedRows?.length ?? Math.max(0, Math.min(last.row, index.rows - 1) - first.row + 1);
    for (let position = 0; position < rowCount; position++) for (let col = first.col; col <= Math.min(last.col, index.columns - 1); col++) {
      const item = itemAtCoordinate(selectedRows ? selectedRows[position] : first.row + position, col);
      if (!String(item.value ?? '').toLocaleLowerCase().includes(needle)) continue;
      if (total >= offset && items.length < limit) items.push(item);
      total++;
    }
  } else for (let position = offset; position < Math.min(count, offset + limit); position++) items.push(itemAt(position));
  return { items, total, offset, nextOffset: offset + items.length < total ? offset + items.length : null,
    ...(keys ? { missingKeys, duplicateKeys } : {}), calculation: noCalculation(revision),
    coverage: { sheet: 'CSV', range, ...(keys ? { keys } : {}), completeDocument: false,
      omissions: ['unrequested-ranges'], notChecked: [],
      format: 'csv', literalValuesOnly: true, note: 'CSV stores text fields, without formula calculation, workbook formatting or multiple sheets.' } };
}

export function readCsvDocument(task, bytes) {
  requireThat(!task.args.table, 'UNSUPPORTED_FEATURE', 'CSV has no named tables; query sheet CSV by range or business key.');
  const index = readCsvWorkbook(bytes), revision = digest(bytes), args = task.args;
  const inventory = { name: 'CSV', rows: index.rows, columns: index.columns, usedCells: index.nonemptyCells };
  if (task.kind === 'office_read') {
    if (!args.sheet || args.scope === 'overview') return { ...readWindow([{ kind: 'sheet', ...inventory }], { offset: 0, limit: 100, ...args }),
      features: [{ feature: 'csv-literal-table', state: 'read' }], notChecked: ['cell-values'],
      coverage: { completeDocument: false, omissions: ['cell-values'], literalValuesOnly: true } };
    const result = readCsvRange(index, args, revision);
    return { ...result, items: result.items.map(item => ({ kind: 'cell', ...item })),
      features: [{ feature: `cells:CSV!${args.range ?? 'A1:J20'}`, count: result.items.length, state: 'read' }],
      notChecked: result.coverage.omissions };
  }
  if (!args.sheet) return { sheets: [inventory], namedRanges: [], tables: [], readOnly: false,
    format: 'csv', delimiter: index.delimiter, encoding: index.encoding, bom: index.bom, calculation: noCalculation(revision),
    limitations: ['CSV stores literal text fields; formulas, formatting and workbook objects require XLSX.'] };
  return readCsvRange(index, args, revision);
}

export async function runCsvWorkbook(task, bytes) {
  try {
    if (task.kind === 'csv_oss_read' || task.kind === 'office_read') return readCsvDocument(task, bytes);
    if (task.kind === 'csv_validate') {
      let expected = fs.readFileSync(task.original);
      for (const operation of task.operations) expected = editCsvWorkbook(expected, operation.commands, digest(expected)).bytes;
      requireThat(digest(expected) === digest(bytes), 'VALIDATION_FAILED', 'CSV differs from its accepted command batch.');
      readCsvWorkbook(bytes);
      return { passed: true, revision: digest(bytes), profile: PROFILE,
        checks: ['guarded-command-replay', 'exact-encoding-and-unmodified-text-preservation', 'csv-reopened'], notVerified: [] };
    }
    const result = editCsvWorkbook(bytes, task.args.commands, digest(bytes));
    fs.writeFileSync(task.output, result.bytes, { flag: 'wx', mode: 0o600 });
    return { revision: digest(result.bytes), affectedTargets: result.affected.slice(0, 100), engine, evaluation: false,
      postconditions: { affectedCellCount: result.affected.length, affectedCellsTruncated: result.affected.length > 100,
        calculation: noCalculation(digest(result.bytes)), literalValuesOnly: true, encoding: result.index.encoding } };
  } catch (error) {
    if (error instanceof OfficeError) throw error;
    const code = /Conflict:|Duplicate key|Missing key/.test(error.message) ? 'TARGET_CONFLICT'
      : /unsupported|does not support|does not store|cannot|CSV.*(?:formula|format)/i.test(error.message) ? 'UNSUPPORTED_FEATURE' : 'INVALID_ARGUMENT';
    throw new OfficeError(code, error.message);
  }
}
