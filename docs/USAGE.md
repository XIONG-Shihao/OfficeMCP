# Usage

Commands in this guide run from `OfficeMCP/`. OfficeMCP is local stdio software;
it has no HTTP listener or account service.

## Install and run

Node `>=24.12.0 <25` is required.

```sh
npm ci --ignore-scripts
npm run build
npm test
npm run doctor
node /absolute/OfficeMCP/src/cli.mjs serve --config /absolute/config.json
```

Only MCP messages use stdout. Diagnostics use stderr. One process owns each state
directory, so separate installations need separate state directories.

All configured roots must already exist. Relative paths resolve from the config
file. The state directory must remain outside every read/write root.

```json
{
  "readRoots": ["/absolute/documents"],
  "writeRoots": {"output": "/absolute/output"},
  "stateDirectory": "/absolute/private-office-state",
  "toolset": "docx-agent",
  "engines": {
    "libreOffice": "/absolute/path/to/soffice",
    "tesseract": "/absolute/path/to/tesseract"
  }
}
```

Optional engine paths may be omitted. See [Engines](ENGINES.md).

An MCP host launches an absolute Node or Bun executable and passes the absolute
CLI and config paths. A host that embeds the runtime in a multi-mode executable
sets `ELECTRON_RUN_AS_NODE=1` or `BUN_BE_BUN=1`; the server passes that on to its
workers.

```json
{
  "mcpServers": {
    "office": {
      "command": "/absolute/path/to/node",
      "args": [
        "/absolute/OfficeMCP/src/cli.mjs",
        "serve",
        "--config",
        "/absolute/config.json"
      ]
    }
  }
}
```

The OfficeMCP-enabled RFP Response Studio build configures this connection using
project-scoped roots and private state. The normal app build does not bundle or
enable OfficeMCP automatically.

## DOCX agent workflow

Set `"toolset": "docx-agent"`. The 14-tool agent-facing DOCX contract
(`docx/6`) is the primary DOCX interface:

1. `docx_open` or `docx_create` returns a revision-bound view.
2. `docx_inspect` with a view returns inventory, outline, domains and a reading plan.
   Without a view, it returns capability help without opening a document.
3. `docx_read` returns bounded canonical items and explicit coverage.
4. `docx_query` resolves stable targets by text, heading path, table content or
   typed document entity.
5. `docx_apply` previews or executes an action, plan or named workflow. Optional
   `readback` returns bounded body text/table changes and revision-bound targets.
6. `docx_validate` and `docx_compare` provide semantic and package evidence.
   Display verification belongs to the root application's scripted GUI tests.
7. `docx_save` publishes a new DOCX; `docx_export` publishes a PDF conversion.
8. `docx_operation_get` recovers an uncertain write or publication result.
   `docx_close` releases the document.

Every edit is bound to the supplied view and an immutable `operationId`. Reuse the
same ID and arguments after a lost reply. Use a new ID for intentionally new work.
Mutation results remain private until publication, and publication never
overwrites an existing file.

Call `docx_inspect` with `actionFamily: "tables"` for the table action names and
schemas, or selected `actions` for only those exact argument schemas. The default
response retains the flat action list and adds compact `actionFamilies` summaries.
Use `workflows: ["update_table_by_key", "replace_section"]` to request workflow
schemas and examples without a live view. Keyed updates require unique headers
and row keys; section replacement requires exact old text and an explicit
subsection policy. See the operations guide for supported document shapes.
The complete executable schemas are in
[`contracts/docx-agent-tools.json`](../contracts/docx-agent-tools.json). See
[DOCX operations](DOCX_OPERATIONS.md) for action groups and workflow examples.

### Reading completely

Follow `nextOffset` with the returned view until it is null. Read every story and
domain required by the user request. Coverage distinguishes read data, detected
but uninterpreted features, and unchecked areas. A document overview is not a
complete read.

### Images, dry runs and display

`docx_image_get` retrieves original PNG/JPEG bytes, and `office_image` returns
an existing image asset. Neither generates a page preview. `docx_apply` with
`mode: "preview"` is a dry-run mutation plan; it does not render the document.
DOCX PDF export remains available through `docx_export` with LibreOffice.

Core contract `0.13` / DOCX `docx/6` remove all four `*_render` tools,
DOCX page/layout selectors and visual-evidence inputs. Select content semantically
and validate its saved package. The root application's scripted GUI checks own
visual display acceptance.

## Four-format workflow

Set `"toolset": "recommended"` for ten common `office_*` tools plus
`xlsx_query` and `xlsx_apply`:

`open, create, read, edit, save, image, extract, engine_status,
operation_status, close`.

The normal sequence is:

1. `office_open` or `office_create` returns a document/candidate view.
2. `office_read` returns bounded items, coverage and the view for the next call.
3. `office_edit` accepts one typed action and returns a new view.
4. Read the new view, then use `office_save` to validate and publish it.

XLSX reads require an explicit sheet/range for useful coverage. PPTX and PDF reads
require a slide/page. Formula expressions, stored caches and numeric lexical text
remain separate. Existing images and PDF attachments can be extracted without
executing their content.

For XLSX and CSV, use the [spreadsheet MCP guide](XLSX_MCP.md) for the current
query/apply commands and format-specific limits.

Set `"toolset": "all"` for the complete 53-tool surface. Use the generated
[`contracts/tools.json`](../contracts/tools.json) as the field-level reference;
duplicating all 51 schemas in prose is deliberately avoided. This surface contains
lower-level format tools, validation/publication primitives, package/asset tools,
state maintenance and compatibility job tools.

Configurations using `"docx"` or `"docx-advanced"` are normalized to
`"docx-agent"`. New configuration should use the current name directly.

## Tasks, cancellation and recovery

On MCP `2026-07-28`, eligible long calls may return an
`io.modelcontextprotocol/tasks` task when the client declares the extension.
Use `tasks/get`, `tasks/update` and `tasks/cancel`. Without that capability, the
same tool executes synchronously.

Compatibility clients can still use background job handles. Publication is short
and synchronous; cancellation never rolls back a file already published.
Operation recovery is separate from task polling: use the relevant operation
status tool with the original `operationId` before repeating uncertain work.

## Limits and lifecycle

Defaults are 320 MiB per source/candidate, 384 MiB expanded archive size, two
workers, 120 seconds per worker, 2 GiB private state, 64 open documents and a
200,000-byte JSON response budget. At most 10,000 non-expired operation receipts
are retained by default (`maxActiveOperations`); permanent expired-ID tombstones
do not consume that active capacity. DOCX actors can be configured from one to four.
Non-spreadsheet workers have a 768 MiB JavaScript heap limit; native engine memory
is separate. XLSX/CSV workers use the runtime's default heap policy, with no
application-imposed cell-count or heap-MB limit. Configurable file/state quotas,
bounded responses, cancellation and worker deadlines still apply. Performance
measurements are diagnostic; successful work requires correct, preserved files.

Spreadsheet mutation receipts bound their evidence to the configured response
budget before accepting the candidate. `omittedFields` reports lengths and hashes
for large values, and `evidence` reports omitted targets. This does not truncate
file contents or weaken expected-value guards. Read the returned view for further
evidence. Reduce read windows when responses are too large; a single unusually
large field may require an operator-configured larger response budget.

The supplied 234 MB `test_doc/response.docx` fits these defaults. Run
`npm run test:response` for local semantic retrieval/edit/save acceptance. It does
not render pages or test page placement. Memory must still be measured: worker
heap flags do not cap SDK subprocesses or native engine memory.


Handles expire after seven idle days. Receipt bodies expire after 30 days, while
operation-ID tombstones remain reserved. `office_cleanup` removes eligible
private artifacts only; it never deletes source or published files.

Workers use process boundaries, quotas, cancellation and deadlines. This is not
an OS filesystem/network sandbox or a hard native-memory boundary. Keep the state
directory private because it contains snapshots and operation evidence.

## Verify and package

```sh
npm run build
npm test
npm run test:package
```

`test:package` creates a tarball, installs it offline outside the repository, and
tests the installed CLI and public SDK. Its npm cache must already contain locked
dependencies. Optional native checks and the agent benchmark are documented in
[Release verification](RELEASE.md).

Generated evaluation files belong under `test_doc/officemcp-results/`. Tests must
not modify the supplied originals.
