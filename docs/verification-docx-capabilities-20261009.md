# DOCX capability verification — 2026-10-09

The current local preview implements guarded support for chart titles/data,
textbox/equation text, page-based text edits, rich section replacement and
merged/nested table-cell text. SDK APIs are preferred where they expose a safe
range or structural operation; unsupported object shapes use declared OOXML
adapters. The implementation belongs to standalone OfficeMCP. No app execution
adapter, harness/model-specific route or SDK-internal patch is required.

## Product identity and acceptance

| Gate | Result |
|---|---|
| Package / DOCX contract | `0.18.0-preview.1` / `docx/6`; 14 tools, 79 actions, 12 families |
| Independently installed production identity | `51c5c898fd6bf6edadc14248ca3229b60fb12169da1b93695be9ac611698ec37` |
| Tested tarball SHA-256 | `4a16bf2ea4cfebe5d8c2c25e81182dae72734f89eb1995cee0af35fce1a3c78b` |
| Frozen SDK / native binary | SuperDoc 2.2.1 / `376beb10ba4c28d5ab7b77e0a9330e4dcd57341ab7b60af02364bd7ab7b245b9` |
| Fresh offline npm install | Pass; no consumer overrides or lifecycle scripts |
| Ordinary public-client workflows / blank creates | 24 / 3 pass |
| Expanded public-client workflows | 24 pass; each save independently checked and recovered after CLI restart |
| Unrelated public-document jobs | 36 pass on the installed package |
| Full 234 MB bid, explicit merged-owner edit | 1 pass, independently checked |
| Relevant source regression suite | 55 pass, no skips |
| Actual Word object/page qualification | 2 pass with Word 16.113.3 |
| Modern official-client facade regression / generated guide | Pass / pass |

The three clients are raw modern `2026-07-28`, raw legacy `2025-11-25` and the
official modern MCP client. All use public discovery/calls and ordinary CLI state.
Tripwire app/publication adapter settings are deliberately supplied; importing the
private app adapter would fail the run. Production identity matches the retained
fresh installation. Later documentation updates do not claim identical tarball
documentation bytes. Qualification is macOS arm64; cross-platform builds/testing
are outside this work.

## What is implemented

| Operation | Route and verified behavior | Guarded limits |
|---|---|---|
| Change a chart title | Existing direct rich-text title through OOXML; retained title formatting and unrelated parts; visible Word change | Missing/computed titles refused |
| Change chart data | Unique series values/categories; cache and embedded literal XLSX cells updated together; unrelated workbook cells/styles/members retained; Word chart visibly changes | Existing point count only; formulas, external sources, pivot charts, shared workbook/ranges and sparse caches refused |
| Edit textbox text/equation symbols | Literal ranges retaining text runs or OfficeMath structure; preview, batch rollback, save/reopen and visible Word result | Textbox edits stay in one paragraph; protected content and mathematical reconstruction/evaluation refused |
| Page-based text editing | Word/LibreOffice layout → unique page phrase → actual SDK range; page 2 changed while page 1 stayed; old layouts/repeated phrases refused | No page-number-only, ambiguous, cross-page or opaque-object targeting; no universal Word appearance guarantee |
| Replace a section with tables/lists/objects | SDK handles ordinary structured blocks; explicit retain indices select guarded OOXML for exact object/review-block retention; outside content preserved | Objects/protected blocks cannot be implicitly deleted; retained blocks keep relative order; 100 source blocks; new lists on package route require an existing numbering part |
| Edit merged/nested cells | Explicit cell/merge-owner identities; SDK range preflight then guarded fallback; horizontal, vertical, legacy horizontal and nested cases; repeated paragraphs and empty-cell insertion | Text editing preserves topology; inconsistent/nonempty continuations and protected/object-bearing selected paragraphs refused; body tables only |

`docx_read` domains `objects`, `table_cells` and scoped `section_content` expose
usable identities and limits. Discovery explains the preferred keyed route and the
explicit merged-owner route. Inapplicable entity filters or section scopes refuse
instead of silently selecting unrelated content. Object mutations reject missing
OPC content-type declarations, protected/signed packages and comment-owned parts.
Comments remain strictly read-only.

The bid's price cell is a vertical merge owner spanning several business rows.
The keyed workflow deliberately requires an unmerged selected value, so it refuses
that shape. The new explicit cell action targets the shared owner and changes
`10.50` to `10.55`, preserving the other cells, merge structure, formatting and
binary members under the independent oracle. Editing it is an edit of the shared
price owner, not a value belonging exclusively to one row.

The four unrelated public inputs are the UK Model Services Contract template,
Australian Board of Taxation annual report, Alameda procurement RFP and service
agreement template. They provide ordinary document regression coverage, not
chart/equation coverage. Two are operational templates, not signed contracts.
Complex-object tests use correctly typed synthetic fixtures accepted by Word.
Word PDF comparisons establish results for those artifacts on this environment;
complete font, arbitrary-object and cross-machine equivalence are not established.

## Reuse and retained evidence

[Release instructions](RELEASE.md#expanded-docx-qualification) describe the permanent
fresh-install runner. `docx-expanded-workflows.mjs` runs the new jobs through an
installed package; `test:docx:consumer` runs ordinary and expanded jobs together.
The production tarball does not import the source-checkout test fixtures/harnesses.

The aggregate is accepted-qualification-20261009.json (`../../.build/runs/docx-capability-expansion/accepted-qualification-20261009.json`).
Detailed results live under `.build/runs/docx-capability-expansion/`:

- `consumer-reviewed-20261009/`: exact installed identity, ordinary and expanded public-client traces and independently checked outputs.
- `public-corpus-reviewed-20261009/`: 36 independent public-document jobs.
- `stress-package-cell-reviewed-20261009/`: complete bid and explicit merged-owner saved-output evidence.
- `regression-reviewed-20261009.log`, `word-reviewed-20261009.log`, `public-facade-reviewed-20261009.log`, `guide-reviewed-20261009.log`.

Earlier failed runs remain evidence. The initial synthetic object fixture omitted
a binary content type and an explicit textbox shape definition; Word verification
caught these fixture defects. They were corrected, and missing content types now
refuse mutation. One overlapping Word test run correctly returned BUSY; accepted
Word runs are sequential. SDK section tests also caught inherited heading formatting
at empty-section boundaries; the producer now explicitly produces body paragraphs
there and verifies list membership/type through SDK list APIs.

No new live-model qualification or headless timing/memory acceptance is claimed
for this version. Historical model/benchmark records apply to their original
version. Performance work remains the final stage; full-app optimization is deferred.
