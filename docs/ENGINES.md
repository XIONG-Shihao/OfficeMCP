# Document engines

Dependencies are pinned in [package.json](../package.json) and
[dependencies.json](../contracts/dependencies.json). Engine and license paths are
operator configuration, never agent tool arguments. The current package is
`0.18.0-preview.1` (core contract `0.13`, DOCX contract `docx/6`). The legacy
render tools remain removed. Optional DOCX layout and reference comparison are
available through `docx_layout` under the newer capability-expansion scope.

| Task | Engine | Boundary |
|---|---|---|
| DOCX ordinary create/read/edit/save | SuperDoc SDK 2.2.1, runtime v2 | Preferred engine; SDK internals stay unchanged |
| DOCX complex content | Guarded OOXML adapters | Declared textbox/equation/chart, merged/nested cell text and object-retaining section changes; no general package repair |
| DOCX package evidence | OOXML/XML and binary inventory | Read-only validation; no repair fallback |
| DOCX PDF export | Optional LibreOffice conversion; MuPDF verifies the PDF | Creates a PDF artifact; no page images or layout mapping; not exact Microsoft Word layout |
| DOCX layout/reference | Optional Microsoft Word on macOS or LibreOffice; MuPDF text/pixels | Revision-bound unique semantic text targets and per-artifact appearance evidence |
| XLSX creation | ExcelJS 4.4.0 | New workbooks, not arbitrary round trips |
| XLSX shared editor/MCP path | Univer OSS 0.25.2 and preserving OOXML reader/writer | Local calculation without a paid engine; unsupported dependencies remain explicitly unresolved; external data is not refreshed |
| XLSX optional legacy operations | Preserving OOXML adapters; optional Aspose.Cells 26.9.0 | Advanced legacy engine operations retain their declared license requirements |
| CSV existing-file work | Shared indexed literal-table reader and preserving writer | Text fields, unique-key guards and atomic edits; no formulas or workbook formatting; UTF-8 and BOM-tagged UTF-16 |
| PPTX creation | PptxGenJS 4.0.1 | New slides with text, tables, images and notes |
| PPTX existing-file work | Preserving OOXML adapters; optional Aspose.Slides 26.8.1 | Advanced Aspose mutations require a license |
| PDF inspect/search/extract | MuPDF 1.28.1 | Text and positions, outlines, annotations and attachments; no public raster output |
| PDF create/forms/pages | pdf-lib 1.17.1 and fontkit 1.1.1 | Guarded operations; no general text reflow |
| Unicode PDF text | MuPDF extraction; pdf-lib/fontkit with bundled Noto Sans SC for creation | Extraction and embedded-font support are separate from visual reading order |
| PDF OCR | Optional Tesseract with a bounded private MuPDF raster intermediate | Returns fallible words, confidence and coordinates; the intermediate is deleted |

The `docx` npm dependency is a test-fixture producer, not a production engine.
Optional binaries are not downloaded by document operations.

## Configure optional local engines

```json
{
  "engines": {
    "libreOffice": "/absolute/path/to/soffice",
    "tesseract": "/absolute/path/to/tesseract"
  }
}
```

Use `office_engine_status` to inspect Aspose installation and license status.
LibreOffice is used only for DOCX PDF export; Tesseract is used only for requested
OCR. Configured executable paths do not by themselves prove that conversion or
recognition can run with the installed dependencies and language packs.

## Install Aspose locally

From `OfficeMCP/`:

```sh
python3 -m venv .engines/aspose
.engines/aspose/bin/python -m pip install -r src/engines/aspose-requirements.txt
```

Windows uses `.engines/aspose/Scripts/python.exe`. Aspose.Cells and
Aspose.Slides are optional commercial engines. On macOS, Slides normally also
requires `mono-libgdiplus`.

```json
{
  "engines": {
    "python": "/absolute/OfficeMCP/.engines/aspose/bin/python",
    "asposeLicense": "/absolute/licenses/Aspose.Total.lic",
    "allowEvaluation": false
  }
}
```

Use `asposeCellsLicense` and `asposeSlidesLicense` to override the shared license.
License contents are never returned to tools.

Without a license, `allowEvaluation: true` permits Cells evaluation and Slides
inspection. Evaluation results are labeled; Cells can add evaluation sheets.
Slides evaluation truncates text on load, so Slides mutations
require a valid license. There is no silent destructive fallback.

## Limits for the supplied large DOCX

The supplied `test_doc/response.docx` (about 234 MB) fits the current default
budgets. Hosts can explicitly retain or override these settings:

```json
{
  "limits": {
    "maxFileBytes": 335544320,
    "maxExpandedBytes": 402653184,
    "maxStateBytes": 2147483648,
    "workerTimeoutMs": 120000
  }
}
```

Larger limits increase disk and memory use. At most four DOCX actors can be
configured. Non-spreadsheet Node workers have a 768 MiB JavaScript heap ceiling;
native engine memory is separate. XLSX/CSV workers use the runtime's default heap
policy. File/state quotas, bounded results and cancellation remain configurable
operational controls; there is no benchmark-cell-count or memory-MB acceptance cap.

## Preservation boundaries

XLSX/PPTX engine round trips report changed, added and removed package parts and
check original binary-resource preservation. XML normalization is disclosed; it
does not prove unchanged layout or semantics.

PDF page edits refuse signed documents and feature-rich cases where deletion or
duplication would require semantic remapping. Complete permutations of original
pages are supported for tested forms and outlines.

DOCX validation distinguishes SuperDoc reopen/semantic results from independent
package checks. The general path compares SDK snapshots. One explicit direct
heading insertion at document end can use a narrower serialized proof: source
domains survive baseline serialization, existing document XML and unrelated
package members stay unchanged, and the new heading matches the requested text,
level and position. Save reuses that proof only for the exact immutable candidate
hash chain. Unsupported cases restart from the original snapshot and use the full
SDK roundtrip. A declared action is
available to request; its presence does not imply every action-specific round-trip
case has passed. Content is targeted through text, node, table and heading queries;
OfficeMCP does not create page maps or claim exact Microsoft Word layout.

Visual inspection belongs to the consuming host and is separate from MCP validation.

## DOCX process resources

The DOCX pool defaults to two active requests and two resident workers. Its
global pending queue is limited to 32 requests, with eight per document. Clean
idle workers expire after 60 seconds; their SDK descendants must exit before a
resident slot is reused. Immutable document/candidate views remain available
for a later worker to reopen.

CLI configuration `limits` accepts `maxDocxActors`, `maxDocxResidentActors`,
`maxDocxQueuedRequests`, `maxDocxProcessBytes` and `docxIdleMs`. Active and
resident limits permit values from one to four; the global queue permits up to
128. The process-memory default is 8 GiB across the owned worker/SDK trees.
This is checked every 500 ms and before accepting worker results. It is a
sampled operating budget, so brief allocation overshoot can occur. Exceeding
it refuses the result and terminates the affected workers; no unverified
candidate is accepted. The Node heap flag alone does not bound SDK memory.

## Verify optional engines

```sh
OFFICEMCP_LIBREOFFICE=/path/to/soffice npm test
OFFICEMCP_TESSERACT=/path/to/tesseract npm test
OFFICEMCP_ASPOSE_PYTHON=/path/to/aspose/python npm test
OFFICEMCP_ASPOSE_PYTHON=/path/to/aspose/python npm run test:engines
```

The LibreOffice checks exercise PDF export; the Tesseract checks exercise text
recognition. Set `OFFICEMCP_ASPOSE_LICENSE` for a licensed sample run. Verification writes
copies and reports under `test_doc/officemcp-results/`; originals must remain
unchanged.

The package is private and `UNLICENSED`. Optional engines retain their own license
terms and are not bundled with the npm package.

### Headless DOCX memory boundary

The native SDK remains unchanged. OfficeMCP copies and hashes session inputs with bounded buffers, reads preservation packages member by member, and validates XML syntax without retaining a whole-document DOM. Semantic checks still use structured XML where needed; media preservation still compares exact bytes. Process observation uses a counters-only module so the coordinator does not load the SDK merely to supervise its worker. These paths are shared by the public CLI and direct service consumers; there is no app-specific faster mode.

For cold native startup, OfficeMCP supplies the installed executable through the SDK's public binary option when it is already executable. Lookup follows the SDK's own dependency tree, preserving its pinned version when a host uses another SDK or an isolated package-manager layout. It neither rewrites SDK code/binary nor changes its permissions. Missing/non-executable cases retain the SDK's ordinary resolution and permission recovery. All clients use this same product path; this is not an app-specific integration mode.

Consuming hosts own scripted GUI checks, separate from MCP validation.
