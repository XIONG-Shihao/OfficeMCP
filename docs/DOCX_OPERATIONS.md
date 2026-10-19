# DOCX agent operations

The current preview exposes 14 tools, 79 actions in 12 families, under `docx/6`.
Ordinary text, table, list and formatting changes use frozen SuperDoc SDK 2.2.1.
Declared guarded OOXML adapters cover content the SDK cannot safely expose.
OfficeMCP owns these routes; every compatible MCP client receives the same behavior.
Comments are read-only. Full-app optimization is deferred.

Executable schemas: [docx-agent-tools.json](../contracts/docx-agent-tools.json).
Discovery: `docx_inspect({actions:[...]})` or `actionFamily`; workflow discovery
works without opening a document. Availability is not universal document-shape qualification.

The discovery families are `objects`, `text`, `formatting`, `lists`, `tables`,
`images`, `review`, `layout`, `controls`, `links`, `notes` and `references`.

## Common operations

Open/create, canonical read/query, guarded apply/preview, validation, comparison,
image retrieval, PDF export, save, close and receipt recovery remain available.
`docx_layout` adds optional revision-bound page text and Word-reference comparison.
Save publishes a new file without overwriting the input or an existing destination.
Use `docx_operation_get` after a lost response before submitting another write.

## Chart title and data

Read `docx_read({view,representation:"domain",domain:"objects"})`; query the chart
by its returned item ID, then use its opaque target with `docx_apply`.

- `set_chart_title({expectedText,text})`: existing direct rich-text titles only.
- `set_chart_data({seriesName,expectedValues,values,expectedCategories?,categories?})`:
  exactly one series, existing point count, finite numeric values; both category
  arrays must be supplied together. Literal caches and local simple A1 ranges are
  supported. Referenced charts must have an embedded XLSX, whose literal cells
  agree with the cached source values. Cache and worksheet change atomically;
  unrelated workbook members/cells and existing cell styles remain.

External/formula/protected/merged worksheets, pivot charts, workbooks shared by
charts, overlapping data used by another series, sparse caches, adding/removing
points and computed/missing titles are refused. No chart calculation engine is advertised.

## Textboxes and equations

The objects domain supplies targets for `replace_textbox_text` and
`replace_equation_text`, each with `{expectedText,find,replace}`. The expected
value is the full recorded object text; find must occur once. Textbox replacement
stays inside one paragraph. Equation replacement changes literal symbols/text,
retaining OfficeMath structure; it does not evaluate mathematics or rebuild equations.
Protected review/field/reference content and structural text insertions are refused.

## Complex table cell text

Prefer `update_table_by_key` for a unique ordinary value under unmerged first-row
headings and a unique business key. Unrelated merges and nested content survive.

For merged owners, nested tables or repeated cell text, read domain `table_cells`.
Each item records table index, physical row/cell, logical column, spans and merge
owner. A vertical continuation targets its recorded owner; it never authorizes
writing into an empty continuation instead of the displayed cell.

`replace_table_cell_text({expectedText,find,replace,paragraphIndex?})` uses the
returned opaque target. `expectedText` joins the cell's own paragraphs with newline,
excluding nested tables. Find must select one own paragraph; paragraphIndex
resolves repeated paragraphs. Empty find inserts into one empty paragraph.
OfficeMCP attempts an SDK range first and uses guarded OOXML when that range is
not exposed. Merge properties, nested content and unrelated cells survive.

This is text editing in merged/nested cells, not unrestricted table reconstruction.
Inconsistent merge ownership, protected/object-bearing selected paragraphs and
implicit structural deletion are refused. Use dedicated object actions for objects.

## Rich section replacement

Read domain `section_content` with
`section:{headingPath:["Methodology"],subsections:"preserve"}`. Its indices are
relative to the selected section and can be used for `retain`. Join item texts
with newline for expectedText; table cells use tabs and rows use newlines.

`replace_section` keeps the selected heading and takes either `paragraphs` or
`blocks`. Blocks are paragraph, heading, table, list or `{kind:"retain",index}`.
SDK operations replace ordinary text/tables/lists. A request with retained blocks
uses `replace_section_content`, the guarded package route. Objects and protected
review/reference/control blocks must be explicitly retained, keeping their relative
order. Retained blocks and their related parts survive exactly; outside body blocks
are checked. At most 100 source blocks are accepted. An object-bearing replacement
with new lists needs an existing numbering part; a prior SDK list operation can
create it. Arbitrary object deletion and cross-story moves are not supported.

## Page-based text edits and appearance

Configure optional `engines.word` with an installed Microsoft Word.app on macOS,
or `engines.libreOffice` with its executable. `engines.wordScratch` may name a
private render directory; otherwise OfficeMCP uses its state directory. Word's
local Automation/file access and activation must be available. Rendering opens a
private copy read-only, selects it by exact path, exports PDF and closes only
owned files. Temporary source/PDF copies are removed; user documents are untouched.
Macros, external templates/media/data and executable/fetching fields are refused.

1. Call `docx_layout({view,provider:"word"})` or provider `libreoffice`.
2. Query `{type:"page",layoutId,number,pattern}` with a unique literal phrase.
3. Apply to the returned semantic target, or use that page selector inline.
4. Generate a new layout after mutation; old layouts are stale.

Page numbers alone, repeated text, cross-page ranges and opaque object targeting
are refused. Optional reference comparison requires an operator-supplied Word PDF
bound to the exact document revision. Exact pixel equality is a result for those
PDF artifacts and recorded DPI, not a guarantee of universal Word appearance.
`docx_export` remains configured LibreOffice PDF publication.

## Existing comments and evidence

Read/query existing comments, including status. OfficeMCP never creates, edits,
replies to, deletes, resolves or reopens them. Ordinary changes must preserve them.
Tracked-change decisions are a separate, narrower SDK feature.

### SDK round trip

SDK save/reopen checks establish semantic round-trip behavior. Guarded package
changes reconstruct declared mutations and require every unselected member to be
byte-identical.

### Independent package checks

Independent workflow tests inspect saved values, formatting, merges, retained
objects and embedded workbook data. Neither these checks nor SDK round trips
prove every Word object or every page layout.

### Application GUI checks

Use separate scripted GUI checks for the consuming application's visible edit,
save and refresh behavior. Optional revision-bound layout providers establish
page text and reference appearance for the selected document; they do not
qualify a consuming application's UI.

See the [current capability verification report](verification-docx-capabilities-20261009.md) and
[historical standalone qualification](verification-docx-corpus-20261008.md), whose
results apply to its recorded preview rather than the current expanded version.
