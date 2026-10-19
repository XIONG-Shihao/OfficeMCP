# Runtime consolidation

The current package is `0.17.0-preview.2`, with core contract `0.13`, DOCX contract
`docx/5`, and private state schema `3`. OfficeMCP provides semantic document tools:
create, inspect, read, query, edit, compare, validate, extract and publish. Public
rendering tools, page-layout selectors, visual comparison and visual-validation
arguments have been removed. This is an intentional breaking surface change.

The earlier consolidation shared lifecycle and format mechanisms while preserving
then-existing tools. Grouped action discovery added `actionFamily` and family
summaries to `docx_inspect`. Those semantic improvements remain; the subsequently
removed rendering feature is no longer part of the current architecture.
See [scope and ownership](SCOPE.md) for the complete boundary.

## Current responsibilities

- `application/document-runtime.mjs` and its `document-sessions.mjs` base own
  document creation, template candidates, target/read handles, session cleanup,
  semantic validation records, save orchestration, and candidate/artifact
  publication. Save calls internal validation and publication APIs instead of
  constructing public-tool calls.
- Each executable operation declaration owns its input/output schema, policy,
  callable application command and discovery metadata. Catalog generation consumes
  these declarations once. Result schemas are reusable envelopes. Format and
  native worker commands use exact dispatch tables.
- `application/artifact-runtime.mjs` shares allocation, quota checks, worker output
  registration, exact embedded-image responses and failure cleanup for extraction.
  Returning an existing document image is distinct from generating a page preview.
- Native and agent DOCX adapters share a mutation-plan boundary, domain inventories,
  text/story/review/table primitives and save/reopen verification. Their guards and
  ordering rules remain explicit. Similar names are not made aliases when
  occurrence or paragraph-spacing behavior differs. `docx_apply(mode="preview")`
  remains an isolated mutation dry run; it does not produce a visual preview.
- `formats/common/document-export.mjs` provides optional DOCX-to-PDF conversion
  through LibreOffice and checks the resulting PDF with MuPDF. It creates no page
  images, layout maps or comparison artifacts.
- `formats/pdf/pdf-ocr.mjs` provides optional text recognition through Tesseract.
  Its bounded private raster intermediate is deleted after use; tools return
  recognized words, confidence and coordinates.

Format facades do not own coordinator writes or transactions. Persistence remains
SQLite in the coordinator; the document runtime owns lifecycle decisions.
Independent package-preservation checks remain separate from SDK readback.
Operation-scoped package contexts reuse ZIP members, metadata XML and relationships;
there is no cross-revision package cache.

The consuming host owns document display and GUI acceptance; no application
renderer or desktop build is distributed in this repository.

## Publication and state transition

Candidate save requires semantic validation bound to the exact candidate revision.
Both save and PDF export copy to an exclusive temporary file beside the destination,
verify its revision and sync it before journaling and atomic no-overwrite linking.
This works when private state and outputs are on different filesystems. Journal
failures and proven destination collisions clean up their staging files; uncertain
post-link outcomes retain recovery evidence. Cancellation is checked before accepting
results or publishing. Durable operation receipts preserve retry/recovery identity.

State migration `3` retires old layout records and removes visual publication
requirements. Mixed structural/visual validations are retained as retired records
with `passed: false` and their prior outcome recorded for audit. They cannot
authorize a new publication; save obtains fresh semantic validation. Existing
structural-only validations remain usable. Migration is transactional and records
a checksum and retirement counts. Legacy private layout files are reclaimed by
normal artifact cleanup, not by unlinking paths inside the database migration.
See [the state transition](SCOPE.md#contract-and-state-transition) before rolling
back to an older package.

## Acceptance and measurement

The pre-consolidation suite's 109 passes and five optional-engine skips are
historical evidence. `test/fixtures/operation-contracts.json` captures the old
per-preset schema/policy hashes. Compatibility tests distinguish retained semantic
operations from explicitly removed rendering operations and their affected schemas.
`state-0.12.0.json` contains receipts, records and file bytes captured by executing
the archived 0.12 release, rather than regenerating expected data from current code.

Current regressions cover semantic read/edit/save behavior, route parity, failed
mutation plans, exact asset extraction, conversion/OCR boundaries, retired visual
records, receipt replay and publication recovery. Route parity includes formatted
replacement, insertion, cells/rows, review, stories and bounded read windows through
save/reopen. SIGKILL tests exercise candidate and export publication.

Benchmark reports distinguish repeated tool names from identical requests.
`repeatedToolCalls` is retained only as a documented legacy alias for
`repeatedToolNames`. `repeatedRequests` compares exact arguments after sorting object
keys, retaining revisions and pagination. A repeated request may be an intentional
retry; neither statistic measures wasted work. Older corpus reports have no
request-level evidence and must not be relabeled as though they do.

Large-document revision checks use bounded hashing buffers. Save verification
retains the semantic digest rather than the whole expected model. Above 64 MiB,
the original SDK document closes before reopening; the verified handle is then
adopted by its owner. Raster-buffer/crop improvements measured before rendering was
removed are historical only; those preview/comparison implementations and their
renderer-only tests are no longer in OfficeMCP. Retained benchmark and diagnostic
scripts are development tooling. No general agent-efficiency or Microsoft Word
fidelity claim follows from a local run.

## Earlier-stage verification (2026-09-17; historical)

- Full suite with LibreOffice enabled: 124 passed, zero failed, three optional
  OCR/Aspose integration checks skipped. Real visual mismatch rejection and PDF
  export ran; these were not mocked in the engine integration tests.
- Standalone build and root app build passed. Installed-package checks passed,
  including identical repeated packs, offline installation, four-format workflows,
  DOCX publication and the installed agent contract.
- The 234,376,124-byte `response.docx` passed all 19 calls, including page-12
  insertion, independent validation, save and reopen. Source bytes were unchanged;
  379 binary members were preserved. The package report discloses eight removed
  XML parts and three changed parts; this is not byte-for-byte package identity.
- The recorded calls totaled 257,151 ms, including 164,586 ms in rendering. This
  is local acceptance evidence, not a controlled speed comparison. Process-tree
  RSS could not be measured because process inspection was denied.
- Evidence is in `test_doc/officemcp-results/consolidation-verification.json` and
  `response-workflows-consolidation.json` beside it. The prior large-file report is
  retained as `response-workflows-before-consolidation.json` in that directory.
  At that stage, integrated SuperDoc visual review remained pending; no Computer Use was used.
