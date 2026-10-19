# Release verification

Product acceptance follows [SCOPE.md](SCOPE.md): independently install OfficeMCP,
exercise its public MCP tools, and inspect the saved documents. Agent tests check
whether models can select those tools correctly. The integrated root app separately
owns workspace publication and viewer acceptance. Prefer scripted GUI tests for
repeatability; user-authorized Computer Use is also permitted.

Version 0.18.0-preview.1; core contract 0.13 and agent DOCX contract docx/6;
private-state schema 3. The package is local stdio only. The optional Node SDK and
SQLite object store are local APIs and do not start a listener.

The primary MCP revision is `2026-07-28`. It uses stateless request metadata,
JSON Schema 2020-12, explicit application handles and the
`io.modelcontextprotocol/tasks` extension. Compatibility remains for
`2025-11-25` and `2025-06-18` clients.

## Reproduce the standalone checks

First complete [external engine source setup](ENGINE_SOURCES.md), including the
two engine-source environment variables. Run from the OfficeMCP checkout with
Node `>=24.12.0 <25`:

```sh
npm ci --ignore-scripts
npm run build
npm test
npm run test:package
npm run test:docx:consumer
```

`test:package` creates repeated tarballs, checks reproducibility, extracts the
standalone package outside the repository and runs `npm ci --omit=dev --offline`
against its shipped root shrinkwrap. It verifies installed runtime versions,
then exercises the stdio executable and frozen SDK exports. Its npm cache must
already contain the locked packages. This does not qualify an arbitrary
consumer's dependency installation: dependency-level overrides are not inherited
from the package. The integrated app bundles the same locked production tree.

`test:docx:consumer` owns the separate DOCX dependency-consumer gate. It packs the
current product, installs it in a fresh temporary npm project without overrides or
lifecycle scripts, then drives the installed CLI over modern and legacy stdio MCP.
The runner imports no app code. It checks formatted clause replacement, a table
cell selected by key/header, a real text content control, existing-comment reading and unchanged anchoring/status, lost-response Save receipt recovery after CLI restart, ambiguity and
stale-target refusals, blank-document creation, and a unique unscoped formatted
replacement and reuse of a paragraph-domain read target. Eight workflow variants
and blank creation run through raw modern/legacy clients and the official
modern MCP client. The report binds the installed production source/contracts/core
identity and the resolved SDK version/native executable hash. Saved-package checks are
independent of OfficeMCP validators; successful jobs also reopen their outputs.

Reports and outputs go to a unique `artifacts/docx-consumer/<timestamp>/` directory.
Set `OFFICEMCP_DOCX_OUTPUT` to select another directory,
`OFFICEMCP_DOCX_TARBALL` to qualify a specific tarball, `npm_config_cache` for an
existing npm cache, and `OFFICEMCP_DOCX_ONLINE=1` to allow registry access. Offline
installation is the default. Missing cached dependencies are installation failures,
not skipped tests. `OFFICEMCP_DOCX_KEEP_CONSUMER=1` retains the isolated installation
for subsequent agent tests; otherwise it is removed after verification.

`npm run test:docx:workflows` runs the same fixed requests against the source CLI.
Use `OFFICEMCP_DOCX_SOURCE=/absolute/package` for an installed package. These are
synthetic shape-specific checks, not complete OOXML/style fidelity or rendering
qualification. The oracle checks text, bold/italic, table values, text-control
identity/value, stories, binary hashes, image placement/size/alt/reference and
comments. It excludes only reserved
separator footnotes/endnotes with IDs -1/0, not user-authored notes.

## Portable actual-agent checks

The product-owned `test:docx:agents` runner uses one generic Chat Completions tool
loop. It obtains schemas from public MCP discovery and has no private OfficeMCP
execution path. The caller supplies provider settings; it does not read app
configuration or saved credentials:

```sh
OFFICEMCP_AGENT_LIVE=1 \
OFFICEMCP_AGENT_API_KEY="$PROVIDER_API_KEY" \
OFFICEMCP_AGENT_BASE_URL=https://your-provider.example/v1 \
OFFICEMCP_AGENT_MODELS=model-a,model-b \
OFFICEMCP_DOCX_SOURCE=/absolute/installed/officemcp-local-preview \
npm run test:docx:agents
```

Only generated synthetic document content is sent. `OFFICEMCP_AGENT_JOBS` selects
comma-separated job IDs; otherwise all six jobs run. Provider-specific request
options can be supplied as JSON through `OFFICEMCP_AGENT_REQUEST_OPTIONS`; this
changes the test client request, never OfficeMCP behavior. The Save-recovery case
withholds a completed response from the model, restarts the CLI and requires receipt
recovery. Reports retain failures, tool arguments/results, provider and MCP time,
exact request repetition and model-visible text bytes. One run per job/model is a
usability sample, not a general reliability rate or proof for other harnesses.

The currently exercised platform is macOS arm64. These checks do not
establish other-platform support or universal Office fidelity.

## Optional native checks

```sh
OFFICEMCP_LIBREOFFICE=/path/to/soffice OFFICEMCP_TESSERACT=/path/to/tesseract npm test
OFFICEMCP_ASPOSE_PYTHON=/path/to/aspose/python npm test
```

Aspose license and evaluation behavior is documented in [Engines](ENGINES.md).

## Local sample checks

`npm run test:response` exercises a caller-supplied `test_doc/response.docx`.
It checks retrieval, guarded editing, recovery, preservation and save/reopen;
it does not verify page appearance. Inputs and generated reports are local and
ignored by Git. No test may modify supplied originals.

## Evidence boundary

Tests cover source immutability, revision conflicts, durable operation IDs,
process-kill recovery, bounded workers, validation and no-overwrite publication.
Workers have quotas and deadlines, not an OS sandbox or hard native-memory limit.
MCP validation does not include document rendering or visual comparison.
Visual display checks use the root application's scripted GUI tests.

Current product scope and gaps are in [Capabilities](CAPABILITIES.md). Versioned
machine evidence is in [`verification.json`](../verification.json). Older entries
inside that file are historical evidence for their recorded version, not current
capability claims.

## Real-document DOCX qualification

Contract `docx/5` removed comment writers; `docx/6` keeps comments read-only. Removed actions are rejected before mutation; existing comments remain readable and must survive ordinary edits and saves. Do not run historical comment-write scenarios as current release gates.

Prepare a reusable manifest locally (the first path is the supplied private bid; extraction adds no task text):

```sh
python3 scripts/test/docx-corpus-prepare.py /absolute/response.docx /absolute/superdoc /absolute/new-fixtures
OFFICEMCP_DOCX_SOURCE=/absolute/installed/package \
OFFICEMCP_DOCX_CORPUS=/absolute/new-fixtures/corpus.json \
OFFICEMCP_DOCX_OUTPUT=/absolute/new-evidence \
node scripts/test/docx-corpus-workflows.mjs
```

Ordinary applicable jobs run through raw modern/legacy and official modern clients. The 234 MB stress input runs once through the official modern client. The independent Python oracle streams binary hashing and checks ordered text, run formatting, tables/merges, controls, stories, drawings, comments and all binary members. It does not certify arbitrary objects or Microsoft Word layout. A merged price owner can be targeted explicitly through domain `table_cells` and `replace_table_cell_text`. Keyed workflows retain their narrower unmerged key/value contract; safe table-scoped SDK text replacement remains available.

For live agents, `scripts/test/docx-corpus-agents.mjs` reuses the generic public-MCP Chat Completions loop. Supply the same provider environment as `test:docx:agents`, plus the corpus and installed-package paths above. Its default applicable job selection covers clause, price cell, control, existing comments, and restarted Save recovery. Private input requires `OFFICEMCP_CORPUS_REAL_EXCERPTS_ALLOWED=1`; the recorded authorization covers only `deepseek-flash` and `deepseek-v4-pro`. Every model-visible tool reply is limited to 32 KiB and accumulated new results to 256 KiB. Oversized or binary responses are withheld with an explicit client error, never silently truncated. Credentials stay with the caller.


## Expanded DOCX qualification

`test:docx:consumer` now runs both the ordinary jobs and
`scripts/test/docx-expanded-workflows.mjs` against a fresh installed tarball.
The expansion runner covers chart title/data, textbox/equation text, merged/nested
cells, SDK rich sections and object-retaining sections across three public clients.
It independently checks saved packages and restarts the CLI to recover each save
receipt. To include the Word page job on macOS:

```sh
OFFICEMCP_WORD='/Applications/Microsoft Word.app' \
OFFICEMCP_WORD_SCRATCH=/absolute/private-word-render-directory \
OFFICEMCP_DOCX_OUTPUT=/absolute/new-evidence \
npm run test:docx:consumer
```

Word must have access to that private directory and local Automation access.
Run Word qualifications sequentially when sharing one scratch directory; another
owner returns `BUSY`. A correctly typed synthetic fixture and saved complex-object
outputs are also checked by `test/docx-object-word-reference.test.mjs`. Reference
pixels demonstrate changes for those fixtures; they do not certify universal Word
appearance. Comments remain read-only. Performance acceptance and full-app
optimization are separate work.
