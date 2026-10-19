# XLSX and CSV MCP tools

OfficeMCP is a local stdio tool server. It reads and edits spreadsheet files for
an agent; the independently distributed XLSX editor owns the visible grid
and previews. The primary spreadsheet workflow is:

`office_open` or `office_create` → `xlsx_query` → `xlsx_apply` → `office_save`

An opened source is an immutable revision. `xlsx_apply` stages a private candidate
and returns its new view. `office_save` validates and publishes that view to a
new destination without overwriting the source or an existing output.

## Actions available to an agent

| Tool | Current action |
| --- | --- |
| `office_open` | Open an existing `.xlsx` or `.csv` file inside a configured read root. |
| `office_create` | Create a private XLSX workbook with sheets, values, formulas and optional tables, or a literal CSV table. |
| `xlsx_query` | Inventory sheets, names and tables, or page through a chosen range. Find rows by unique business key. Return revision-bound cell targets and distinguish display values, raw stored values, formulas, caches and calculation status. |
| `xlsx_apply` | Submit an atomic batch against the current view. Set guarded cells or formulas, set multiple cells, format or fill a range, update rows by unique key, recalculate, or apply supported row/sort/filter/size/freeze operations. |
| `office_validate` | Request semantic and package validation of a candidate. Save also validates before publication. |
| `office_save` | Publish to a new path under a configured write root and return the saved path and validation scope. |
| `office_operation_status` | Recover an edit or save receipt by operation ID after a lost reply or reconnect. |

For example, an agent can query the row whose ID is `SKU-42`, read its current
unit price, then submit an `updateRows` command with that ID and the expected old
price. A missing or duplicate ID, changed old value, or stale view stops the
batch without publishing part of it. The agent can query the returned view,
save a new workbook, reopen it and check the result.

`xlsx_query` returns bounded pages (up to 500 cells per response); `xlsx_apply`
accepts batches of up to 1,000 commands. Those are request/response bounds, not
limits on the size of an existing spreadsheet. Reuse the returned view after
each edit and use a distinct `operationId` for each intended mutation. Retrying
the same ID and arguments recovers the same receipt.

## XLSX commands and calculation

The primary `xlsx_apply` schema supports a cell address or a target returned by
`xlsx_query`; `setCells`; `formatRange` with bold, italic and General/Number/
Percent/Date/Text formats; `fillRange`; qualified `insertRow`; `sortRows`;
`filterRows`; `removeFilter`; `resizeColumn`; `resizeRow`; `freeze`; `updateRows` by unique key;
and `recalculate`. Exact value, raw-value and formula guards are available for
cell updates. Sorting and structural changes are refused when affected workbook
features cannot be preserved safely.

`filterRows` with an empty `values` list shows every row and leaves the filter
switched on, as Excel's Clear does. `removeFilter` switches the worksheet's
filter off entirely and restores the row visibility it had before filtering.
A worksheet with a filter refuses row and column insertion and deletion, so
remove the filter first. `removeFilter` works on filters created through this
tool; a filter that arrived with the file is left alone and reported.

Ordinary XLSX edits and calculation use a local open-source Univer core and a
preserving OOXML writer; no Aspose license or hosted conversion is needed. The
service retains untouched package content where a supported edit permits it.
Formula results report calculated, error, circular or unresolved states. An old
cache or an unsupported dependency is not presented as a newly calculated total.

The primary `xlsx_apply` surface does not add, rename or delete worksheets, nor
delete rows or insert/delete columns. The broader `office_edit` contract exposes
some of those operations through an optional Aspose route with its own licensing
and compatibility requirements. Chart/pivot authoring, full comment/table
semantics and external-data refresh are outside the current shared path.

## CSV behavior

CSV uses the same `office_open` → `xlsx_query` → `xlsx_apply` → `office_save`
sequence, with one sheet named `CSV`. Queries return literal text, including
empty strings, leading-zero IDs and text beginning with `=`. Allowed mutations
are explicit cell edits, `setCells`, and `updateRows` with unique keys and
expected old values. Formulas, formatting, recalculation, workbook objects and
structural commands are refused.

CSV publication preserves unchanged fields, quoting, delimiter, line endings
and source encoding. UTF-8 with or without BOM and BOM-tagged UTF-16 are
supported; malformed or unsupported encoding is refused rather than silently
replacing characters. Save to a `.csv` destination. The current OfficeMCP
contract does not advertise `.tsv` as a separate format, though the desktop
editor can open it.

The executable input schemas are in [xlsx-contract.mjs](../src/formats/xlsx/xlsx-contract.mjs)
and the [workflow contract](../src/contracts/workflow.mjs). For server setup and
configured roots, see [Usage](USAGE.md).
