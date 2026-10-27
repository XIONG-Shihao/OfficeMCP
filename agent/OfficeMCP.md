# OfficeMCP

OfficeMCP works on PDF, Word (DOCX), Excel (XLSX) and PowerPoint (PPTX) files with standardized actions.
Every tool returns `status`: `ok`, or `refused` with a `code`, a `message` and usually a `hint`.

## Workflow

1. `office_inspect` the file: note its `revision`, the addresses in its outline, and `cannotKeep`.
2. `office_read` (an address, or pages of the whole file) and `office_find` (text) for what you need.
3. `office_actions` lists the actions for the format, with their fields.
4. `office_edit` with the file, its `expectedRevision` and a list of actions. All apply or none does;
   the file is saved in place (or as a new file with `output.path`), with a receipt per action.
   Use `dryRun` to check first.
5. Read the receipts. An action that names `expectedText` or `expected` refuses when the content differs;
   read the target again instead of guessing. Check the result with `office_read` when it matters.

Paths are absolute or relative to the project folder. `office_create` and `office_convert` write new files
and never overwrite one.

## Content a save would lose

Before writing, OfficeMCP compares what the file had with what the save keeps. If something would be lost
(a chart, an image, tracked changes, ...), the edit is refused with `fidelity_loss` and the list in
`details.lost`. Decide: leave the file, save as a new file, or repeat the edit with `allowLoss: true`;
the result then lists what was lost in `fidelity.lost`. Tell the person what was lost.

## Addresses

- DOCX: body/p[N], body/table[N], body/table[N]/r[R]/c[C], body/table[N]/r[R]/c[C]/p[K]
- XLSX: Sheet, Sheet!B2, Sheet!B2:D9, Sheet!B:D, table[Name]
- PPTX: slide[N], slide[N]/shape[ID], slide[N]/shape[ID]/r[R]/c[C], slide[N]/notes
- PDF: page[N], form/field[Name]

Ordinals count from 1.

## Tools

- `office_actions`: List the actions office_edit and office_create accept, with each action's fields. Filter by format (docx, xlsx, pptx, pdf) and family.
- `office_inspect`: Outline a file: its structure with addresses, its revision (needed by office_edit), and anything its library cannot keep when saving.
- `office_read`: Read the content at an address (body/p[3], Sheet!A1:D20, slide[2]/shape[5], page[4], ...), or page through the whole file without one.
- `office_find`: Find text in a file; returns the addresses where it occurs.
- `office_edit`: Apply actions to a file, all or none, and save it in place (or as a new file with output.path). Refuses if the file changed since expectedRevision, and if saving would lose content unless allowLoss is true. dryRun checks everything without writing.
- `office_create`: Create a new file at a path, applying actions to it. A new PDF is built from compose actions (add_heading, add_paragraph, add_table, add_page_break).
- `office_convert`: Convert a file to a new file, next to it unless output.path says where: to txt from any format, to csv from a workbook (one sheet), to pdf from docx, xlsx or pptx (needs LibreOffice).

## Actions

### DOCX (Word)

Text inside hyperlinks, fields and content controls is not changed; tracked changes are not produced.

- `set_paragraph_text` (text): Replace a paragraph's whole text, keeping its style and its first run's formatting. Fields: `target`, `expectedText`, `text`.
- `replace_text` (text): Replace text wherever it occurs (in one target, or in the whole body and its tables), keeping each run's formatting. Fields: `find`, `replace`, `target`?, `expectedCount`?, `matchCase`?.
- `insert_paragraph` (structure): Insert a paragraph before or after a body paragraph or table. Fields: `target`, `position`?, `text`, `style`?.
- `insert_heading` (structure): Insert a heading (Word's Heading 1-9 styles) before or after a body paragraph or table. Fields: `target`, `position`?, `text`, `level`.
- `delete_paragraph` (structure): Delete a body paragraph that holds only text. Fields: `target`, `expectedText`.
- `set_paragraph_style` (formatting): Apply a paragraph style the document defines. Fields: `target`, `style`.
- `set_cell_text` (tables): Replace a table cell's text (lines become paragraphs), keeping its first run's formatting. Fields: `target`, `expectedText`, `text`.
- `update_table_by_key` (tables): In a table whose first row holds headers, find the one row whose key column has the key, and set one of its cells. Fields: `target`, `keyColumn`, `key`, `column`, `expectedText`?, `text`.
- `insert_table` (tables): Insert a table of text before or after a body paragraph or table. Fields: `target`, `position`?, `rows`, `style`?.
- `append_table_row` (tables): Add a row at the end of a table, formatted like its last row. Fields: `target`, `cells`.
- `delete_table_row` (tables): Delete one row of a table that has no merged cells. Fields: `target`, `expectedCells`?.
- `add_comment` (review): Attach a comment to a paragraph. Fields: `target`, `text`, `author`?.
- `set_properties` (document): Set the document's title, subject, author or keywords. Fields: `title`?, `subject`?, `author`?, `keywords`?.

### XLSX (Excel)

openpyxl does not calculate formulas: edits report recalculationNeeded, and recalculate needs LibreOffice. Rows are not inserted or deleted in the middle of a sheet. Editing a workbook whose content openpyxl cannot keep (for example some charts or form controls) needs allowLoss.

- `set_cells` (cells): Write values or formulas into a block of cells, starting at a cell. Fields: `target`, `values`, `expected`?.
- `clear_range` (cells): Clear the values of a range, keeping its formatting. Fields: `target`, `expectedNonEmpty`?.
- `set_number_format` (formatting): Apply a number format (for example '0.00', '#,##0', 'yyyy-mm-dd') to a range. Fields: `target`, `format`.
- `set_column_width` (formatting): Set the width of one or more columns, in characters. Fields: `target`, `width`.
- `append_rows` (rows): Add rows after the last used row of a sheet, or at the end of a table (which grows to include them). Fields: `target`, `rows`.
- `update_row_by_key` (rows): In a sheet or table whose first row holds headers, find the one row whose key column has the key, and set cells by header. Fields: `target`, `keyColumn`, `key`, `values`, `expected`?.
- `add_sheet` (sheets): Add an empty sheet. Fields: `name`, `after`?.
- `rename_sheet` (sheets): Rename a sheet that no formula or defined name refers to. Fields: `target`, `name`.
- `delete_sheet` (sheets): Delete a sheet that no formula or defined name refers to. Fields: `target`, `expectedRows`?.
- `add_table` (tables): Turn a range whose first row holds headers into an Excel table. Fields: `target`, `name`.

### PPTX (PowerPoint)

Text in SmartArt and charts is not changed; new slides come only from the presentation's layouts.

- `set_shape_text` (text): Replace a shape's text (lines become paragraphs), keeping its first paragraph's formatting. Fields: `target`, `expectedText`, `text`.
- `replace_text` (text): Replace text wherever it occurs (on one slide or shape, or the whole deck), keeping each run's formatting. Fields: `find`, `replace`, `target`?, `expectedCount`?, `matchCase`?.
- `add_slide` (slides): Add a slide from one of the presentation's layouts, with an optional title and bullet points. Fields: `layout`, `title`?, `bullets`?, `position`?.
- `delete_slide` (slides): Delete a slide. Fields: `target`, `expectedTitle`?.
- `move_slide` (slides): Move a slide to another position. Fields: `target`, `position`.
- `set_notes` (slides): Replace a slide's speaker notes. Fields: `target`, `text`, `expectedText`?.
- `add_picture` (objects): Place a picture (PNG, JPEG, GIF) on a slide. Fields: `target`, `image`, `left`, `top`, `width`?.

### PDF

Text in existing pages is not edited; removing pages that outline entries or form fields point to is refused. Converting Office files to PDF needs LibreOffice.

- `add_heading` (compose, create only): Add a heading to a PDF being created. Fields: `text`, `level`.
- `add_paragraph` (compose, create only): Add a paragraph to a PDF being created. Fields: `text`.
- `add_table` (compose, create only): Add a table of text to a PDF being created. Fields: `rows`, `header`?.
- `add_page_break` (compose, create only): Start a new page in a PDF being created. Fields: none.
- `select_pages` (pages): Keep only the listed pages, in the order listed (reorders or removes pages). Fields: `pages`.
- `delete_pages` (pages): Delete pages. Fields: `pages`.
- `rotate_pages` (pages): Rotate pages clockwise. Fields: `pages`, `degrees`.
- `append_pdf` (pages): Append the pages of other PDFs after the last page. Fields: `paths`.
- `fill_form` (forms): Fill form fields by name. Fields: `fields`, `expected`?.
- `set_metadata` (document): Set the PDF's title, author, subject or keywords. Fields: `title`?, `author`?, `subject`?, `keywords`?.

## Refusal codes

- `invalid_request`: The request does not match the tool's or action's schema.
- `unknown_action`: No action of that name exists for the file's format.
- `unsupported_format`: The file's format is not PDF, DOCX, XLSX or PPTX, or the conversion is not offered.
- `path_not_allowed`: The path is outside the folders OfficeMCP may use.
- `not_found`: The file does not exist.
- `unreadable`: The file is not a valid document of its format, or is encrypted.
- `revision_mismatch`: The file changed since the revision the request names.
- `target_not_found`: The address does not name anything in the file.
- `expected_mismatch`: The content at the address is not what the request expected.
- `no_match`: The text to find does not occur.
- `ambiguous_match`: The key or text matches more than once, or not the number of times expected.
- `unsupported_content`: The target holds content this action cannot change safely (fields, objects, merged cells, ...).
- `fidelity_loss`: Saving would lose content the file has; nothing was written. allowLoss saves anyway.
- `engine_unavailable`: The request needs LibreOffice, which was not found.
- `output_exists`: A file already exists at the new path; only office_edit without output.path replaces a file.
- `too_large`: The file or the requested range is larger than the configured limits.
