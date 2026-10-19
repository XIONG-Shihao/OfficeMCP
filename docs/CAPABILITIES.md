# Current capabilities

Version `0.18.0-preview.1` exposes four surfaces:

- `full` combines 65 four-format and DOCX agent tools (63 on modern discovery), with agent schemas for overlapping names;
- `docx-agent` exposes 14 agent-facing tools under contract `docx/6`;
- `recommended`: 12 common four-format workflow tools;
- `all`: 53 complete four-format tools, or 51 on modern discovery where the
  Tasks extension replaces two compatibility job tools.

The app capability catalog contains 58 entries because its DOCX layer uses the 13
agent tools while the other four layers use their complete format/office entries.

The primary MCP revision is `2026-07-28`. Requests are stateless at the transport
level; explicit handles retain application state. Tool schemas use JSON Schema
2020-12 and eligible long calls use `io.modelcontextprotocol/tasks` when the
client declares it.

Product priorities follow the [capability-first principles](CAPABILITY_FIRST.md).
Generated schemas, not this overview, define exact inputs and outputs.

## User-facing scope

| Format | A user can currently ask OfficeMCP to do | Important gaps |
|---|---|---|
| DOCX | Create documents; inventory and read body/stories/tables/images/comments/review content; query semantic targets; edit text, paragraphs, lists, tables, images, controls, sections, headers/footers, links, notes, captions and cross-references; validate, compare, save and export PDF | Optional Word/LibreOffice layout supports uniquely anchored page text edits. Guarded adapters support literal textbox/equation text, chart titles/data, rich sections retaining objects, and merged/nested cell text. Arbitrary object authoring/table reconstruction and universal Word appearance remain unsupported; complex tracked ranges can fail in SuperDoc |
| XLSX | Create workbooks with sheets, tables and formulas; query ranges and business keys; use `xlsx_apply` for guarded cells/formulas, formatting, qualified row insertion, fill, sort/filter, sizing and freeze; calculate locally | Worksheet management and other structural edits require the separate optional Aspose route; chart/pivot editing, full comment/table interpretation and external-data refresh remain incomplete |
| CSV | Create, read, guard and edit literal cells/unique keyed rows through `office_open`, `xlsx_query`, `xlsx_apply` and `office_save`; preserve untouched text and source encoding | One sheet named `CSV`; fields remain text, without formulas, formatting or workbook objects; UTF-8 and BOM-tagged UTF-16 are supported |
| PPTX | Create slides with text/tables/images/notes; read shapes and notes; replace text, update existing notes, move shapes and reorder slides; optionally edit through Aspose | Rich text replacement, complex charts/SmartArt, master semantics, adding missing note parts and general image/table editing remain incomplete |
| PDF | Create Unicode text pages; inspect/search text, outlines, annotations, attachments and forms; extract attachments; fill supported forms; reorder/rotate pages; optionally OCR | No general text reflow, annotation editing or semantic table extraction; referenced-page deletion/duplication is refused |

Optional engines and their fidelity boundaries are listed in [Engines](ENGINES.md).
The spreadsheet tools and CSV differences are described in [XLSX and CSV MCP tools](XLSX_MCP.md).

## What makes the MCP useful to an agent

OfficeMCP adds behavior that direct document libraries do not normally provide as
one coherent workflow:

- bounded canonical reads with explicit coverage;
- revision-bound views and opaque targets;
- semantic queries instead of guessed array positions;
- guarded mutations and atomic multi-action DOCX plans;
- durable operation receipts and recovery after a lost reply;
- validation before no-overwrite publication;
- native image content for extracted images and existing image assets.

These mechanisms are only valuable when they complete a user job with fewer
calls and less document data sent to the model. They are not product value by
themselves.

## Historical DOCX evidence before contract 0.13/docx/5

The earlier eight-job synthetic corpus covers formatted replacement, template
filling, table/caption work, footer editing, comments, images, large targeted
reading and ambiguous-target refusal. The 2026-09-16 run passed its semantic
checks in 299.4 seconds using 117 calls, including 59 repeated calls and four
expected or recovered errors.

That run demonstrates workflow coverage, not efficiency or general superiority.
The contract case alone needed 43 calls.

The 2026-09-17 deterministic run on the supplied 234 MB `response.docx` passed
19 calls covering inventory, targeted and cross-domain reads, guarded text editing,
receipt retrieval, page-12 mapping/insertion/rerender, package validation and
save/reopen. It preserved 379 binary members and left the source hash unchanged.
The calls took 271.2 seconds in total and returned 83,804 JSON bytes. Peak summed
process-tree RSS was about 11.6 GiB (shared pages can be counted more than once).
This remains memory-intensive, and these are local acceptance measurements,
not a controlled performance benchmark. Full agent-driven workflow coverage and
human visual acceptance in the app remain separate work. Small fixtures continue
to provide deterministic feature assertions.

## Next work

1. Execute the realistic DOCX workflows defined in
   [Capability-first principles](CAPABILITY_FIRST.md): orientation, targeted and
   cross-domain reads, text/table/image edits, semantic insertion,
   save/reopen and recovery.
2. Reduce calls, retries, latency and model-visible bytes in those workflows.
3. Keep document display verification in the root application's scripted GUI tests,
   separate from MCP semantic/package validation.
4. Extend workflow coverage for styles/numbering, review markup, templates,
   nested tables and rich section/header/footer composition.
5. After each workflow contract stabilizes, finish its action-specific,
   adversarial, cross-platform and visual release evidence.
6. Continue PPTX rich-content and PDF annotation/table workflows according to
   user frequency and value.

## Evidence boundary

Validation combines SDK readback with relevant package or format-specific checks.
The MCP package does not render document pages or perform visual comparison. An operation
being listed in discovery does not prove every document shape works.

Current reproducible checks are in [Release verification](RELEASE.md) and
[`verification.json`](../verification.json). Local sample outputs under
`test_doc/officemcp-results/` are not distributed.

Consuming applications own their display and GUI acceptance checks.
