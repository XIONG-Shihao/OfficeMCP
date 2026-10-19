# OfficeMCP

OfficeMCP is a private, local-only MCP server for DOCX, XLSX, CSV, PPTX and PDF. It
runs over stdio and keeps document state in explicit document, candidate,
revision, cursor, asset, operation and task handles.

The fresh npm dependency-install gate is owned here:
`RPF_MCP_CONSUMER_TARBALL=/absolute/package.tgz npm run test:consumer`.
It installs into a disposable external project, verifies runtime dependency and
engine identities, and exercises the installed workbook engine. Test fixtures
come from the declared headless engine package. The default install is offline;
set `RPF_MCP_CONSUMER_ONLINE=1` to allow registry resolution. This gate is distinct
from the standalone CLI archive and MCP protocol tests.

The standalone DOCX gates are `npm run test:docx:consumer` (fresh dependency
installation and saved-document jobs), `npm run test:docx:workflows` (the same
fixed public MCP requests), and `npm run test:docx:agents` (opt-in synthetic live
agent jobs). See [release verification](docs/RELEASE.md) for configuration,
credentials, recovery injection and evidence limits.

Development scope: **provide useful MCP document tools.** Independent editor
components own their UI and engines; the main application integrates them and
displays saved outputs. Protocol and headless tests run here; host applications own their GUI tests. See the
authoritative [scope and testing boundary](docs/SCOPE.md). The current MCP contracts
exclude document-rendering tools; the root application displays saved outputs.

The primary wire protocol is MCP `2026-07-28`, including JSON Schema 2020-12 and
the `io.modelcontextprotocol/tasks` extension. Compatibility serving remains for
`2025-11-25` and `2025-06-18` clients.

## Run

Runs on Node `>=24.12.0 <25` or Bun `>=1.3.13 <2`; either one is enough. macOS arm64 is the currently exercised platform.
The suite runs on both: `npm test` (Node) and `npm run test:bun` (Bun).

```sh
cd OfficeMCP
npm ci --ignore-scripts
npm run build
npm test
node src/cli.mjs serve --config /absolute/config.json   # or: bun src/cli.mjs serve --config ...
```

### External engine source checkouts

This repository deliberately excludes editor engines, editor UI and full-app
integration code. Source builds require separate DOCX/XLSX engine checkouts,
a Rust toolchain for the workbook helper, and the platform C compiler for the
macOS DOCX process observer. See [Engine source setup](docs/ENGINE_SOURCES.md)
for the exact source repository, reference revision and commands.

Set `OFFICEMCP_XLSX_ENGINE_SOURCE` and `OFFICEMCP_DOCX_ENGINE_SOURCE` to their
component directories before `npm run build`, `npm test` or `npm pack`.
`prepare-core.mjs` builds and exports their headless runtime into ignored
`.generated/`. An assembled runtime package includes those exports and does not
need the external source checkouts to run. Cloning OfficeMCP alone is not enough
to build it; missing sources produce a setup error rather than a fallback engine.
macOS arm64 is the currently qualified native target.

See [Repository boundary](docs/REPOSITORY.md) for source ownership and exclusions.

The configured roots must already exist. Keep private state outside input and
output roots.

```json
{
  "readRoots": ["/absolute/documents"],
  "writeRoots": {"output": "/absolute/output"},
  "stateDirectory": "/absolute/private-office-state",
  "toolset": "docx-agent",
  "engines": {"libreOffice": "/absolute/path/to/soffice"}
}
```

There is no HTTP listener, account, authentication token or hosted service.

## Choose a tool surface

| Toolset | Purpose |
|---|---|
| `full` | Combined four-format and DOCX agent API; the app default (agent versions of `docx_read` and `docx_inspect`) |
| `docx-agent` | The primary DOCX surface: 14 tools under contract `docx/6` for open/create, inspect/read/query, apply, validate/compare, image retrieval, export/save, close and receipt recovery |
| `recommended` | Twelve common workflow tools, including shared XLSX query/apply |
| `all` | The complete 53-tool four-format API, including lower-level and operator tools |

For focused DOCX work use:

`open/create → inspect → read/query → apply → validate/compare → save/export`

`docx_read` returns canonical document items rather than raw SDK nodes.
`docx_query` returns revision-bound opaque targets. `docx_apply` performs one
guarded action or atomic action plan. Publication never overwrites a file.

Comments are read-only: read or query comments already present in the DOCX. OfficeMCP does not add, edit, reply to, delete, resolve or reopen them. Tracked-change accept/reject remains a separate feature.

`docx_inspect` groups the 79 actions into 12 families. For example,
`docx_inspect({view, actionFamily: "tables"})` returns the table actions and their
argument schemas together. The default response retains the flat action list.

For mixed-format work use:

`office_open/create → office_read → office_edit → office_save`

For XLSX use `office_open → xlsx_query → xlsx_apply → office_save`. The shared
open-source calculation/writer also powers the desktop editor; ordinary workbook
edits need no paid engine. See the current [XLSX and CSV MCP tools](docs/XLSX_MCP.md).

CSV uses the same workflow with one sheet named `CSV`. Fields remain literal text,
including leading-zero IDs and text starting with `=`. Guarded cell and keyed-row
edits preserve untouched quoting, record endings and UTF-8/UTF-16 BOM encoding.
Save with the original `.csv` extension. Formulas, formatting and workbook objects
require XLSX. Reads remain paginated; workbook dimensions are not capped to a
benchmark cell count, and resource measurements do not replace correctness checks.

Reuse returned views and operation IDs. Reads are bounded and expose coverage;
an unread page, story, slide, sheet or domain is never represented as empty.

Save/export also publish a [portable artifact descriptor](docs/ARTIFACTS.md) naming
the exact output bytes and original source snapshot. Hosts can use it for their
own revision checks; the ordinary saved file remains directly usable.

## Current capability boundary

DOCX creation, reading, mutation and serialization use SuperDoc SDK 2.2.1 with
runtime v2. Read-only OOXML checks provide independent package evidence.
Optional LibreOffice is used for PDF export. Image extraction returns original
image bytes; dry-run action preview reports a proposed mutation without publishing.

Core contract `0.13` and DOCX contract `docx/6` remove `office_render`,
`docx_render`, `pdf_render` and `pptx_render`, DOCX page/layout targets and
visual-comparison inputs. Document display and screenshot checks belong to the
root application's scripted GUI tests. They are outside the MCP package.

Action availability, successful SDK round trips and independent package checks
prove different things. None alone certifies a document's visual appearance;
see [DOCX evidence](docs/DOCX_OPERATIONS.md#existing-comments-and-evidence).

XLSX creation uses ExcelJS; PPTX creation uses PptxGenJS; PDF inspection uses MuPDF. Existing-file operations use preserving OOXML/PDF adapters
and optional local Aspose engines where documented. Optional engines are
configured by the operator and are never downloaded by a document call.

The generated schemas are the executable reference:

- [`contracts/docx-agent-tools.json`](contracts/docx-agent-tools.json) — DOCX agent surface.
- [`contracts/full-tools.json`](contracts/full-tools.json) — combined app surface.
- [`contracts/tools.json`](contracts/tools.json) — complete four-format surface.
- [`contracts/catalog.json`](contracts/catalog.json) — capability browser consumed by the app.

## Source layout

Stable package entry points live at `src/cli.mjs`, `src/sdk.mjs` and
`src/storage.mjs`. Internal modules are grouped by responsibility:

- `application/` — coordination, durable operations and tasks;
- `protocol/` — MCP negotiation, stdio and Tasks;
- `contracts/` — schemas, result contracts and catalog generation;
- `formats/` — DOCX, XLSX/PPTX, PDF and OPC adapters;
- `workers/`, `engines/` and `infrastructure/` — isolation, optional engines,
  configuration, files, errors, diagnostics and quotas.

The execution path is deliberately data-driven. An executable operation spec is
the single source for its input schema, concrete output schema, risk/task policy,
engine metadata, callable application handler and discovery metadata. Its
handler calls a named application command interface; dispatch does not perform
another tool-name or handler-string lookup. Worker commands are selected through
explicit format adapters. All idempotent writes use one durable-operation state
machine. DOCX actions use one descriptor registry for argument parsing, target
validation, engine selection and receipt verification.

The public catalog is generated once from those specifications: the DOCX layer
selects the agent contract and the other layers select the complete format tools.

Output discovery is assembled from reusable result profiles rather than
per-operation schema routers. PDF inspection, search and positioned reading share the bundled MuPDF runtime. The measured reduction history is
recorded in `contracts/source-budget.json`; dependency and behavior tests—not a
line-count freeze—guard maintainability.

The source-layout test enforces the dependency direction between these layers.

## Documentation

- [Scope and portable-product testing](docs/SCOPE.md) — ownership, testing policy and current render/SDK inventory.
- [Capability-first principles](docs/CAPABILITY_FIRST.md) — product priorities and realistic DOCX evaluation.
- [Capabilities](docs/CAPABILITIES.md) — current user-facing scope, gaps and next work.
- [DOCX operations](docs/DOCX_OPERATIONS.md) — exact agent workflow and action groups.
- [Usage](docs/USAGE.md) — configuration, lifecycle and host integration.
- [Engines](docs/ENGINES.md) — optional engines and fidelity boundaries.
- [Embedding SDK](docs/SDK.md) — optional JavaScript API and storage adapter.
- [Release verification](docs/RELEASE.md) — reproducible checks and platform scope.
- [Standalone DOCX qualification, 2026-10-08](docs/verification-docx-product-20261008.md) — installed package, saved-document jobs, live models, fixes and evidence limits.

The package is private, unpublished and `UNLICENSED`. Dependency and optional
engine licenses remain separate.

### Host guide generation

The `officemcp-local-preview/authoring` package export exposes the versioned
operation schemas and DOCX action descriptions used to generate host-side help
and validate examples. Integrators should resolve this export instead of importing
format implementation files. It does not change the MCP tool surface or the
frozen optional embedding SDK. Runtime application integration uses the MCP CLI.
