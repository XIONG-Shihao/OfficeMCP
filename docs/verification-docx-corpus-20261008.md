# Standalone DOCX corpus qualification — 2026-10-08

Current policy: **existing comments are read-only**. OfficeMCP can inventory, read and query them; it cannot create, edit, reply to, delete, resolve or reopen comments. Contract `docx/5` intentionally removes `add_comments`, `resolve_comments` and `reply_to_comment`. Tracked-change accept/reject remains a separate feature.

This report supersedes the comment-write scope in [the earlier synthetic qualification](verification-docx-product-20261008.md). Earlier preview-1 reports and failed runs remain historical evidence. They do not authorize comment features in the current product.

## Verified product and results

| Gate | Result |
| --- | --- |
| Fresh offline npm dependency install; no consumer overrides or lifecycle scripts | Pass, package `0.17.0-preview.2` |
| Installed production source/contracts/core identity | `606da6584dd51a38f4066cb35cb5ee958a200119a17546a114fcf85a700e1a26` |
| Native SDK | SuperDoc 2.2.1, unchanged binary `376beb10ba4c28d5ab7b77e0a9330e4dcd57341ab7b60af02364bd7ab7b245b9` |
| Installed synthetic workflows / blank creates | 24 / 3 pass across three public MCP clients |
| Ordinary corpus fixed jobs | 48/48 pass across raw modern, raw legacy, official modern clients |
| Full 234 MB stress fixed jobs | 4/4 pass through the official modern client |
| Live ordinary / full stress model jobs | 12/12 and 6/6 pass |
| Relevant regression suite | 63 pass; one optional LibreOffice conversion skipped |
| Final read-only comment tests | 3 pass, including generic and legacy reads |
| Independent oracle, exposure budget and stored-evidence acceptance tests | 5 pass |

Removed comment actions and aliases fail before any candidate mutation, including when they appear after a valid edit in a batch. Discovery exposes 73 actions in 11 families. Existing comments remain readable through `docx_read`, comment queries, legacy `docx_read_details` and generic `office_read`. An ordinary unrelated text edit/save preserves the authored existing comment and its open status. Generated contracts, the app's catalog snapshot and the generated agent guide are synchronized.

The tested tarball SHA-256 is `77eb89b3af0ab8c14a0e6a183b1c495e7f16e6c4b6ad2bf4b185f7ebac0e29af`. Production source/contracts/core match the retained independent installation. Documentation was clarified afterward; the report identifies the tested tarball rather than claiming that documentation bytes are identical to every later packaging.

## Documents and supported jobs

The private real source is the unchanged supplied `OfficeMCP/test_doc/response.docx`: 234,376,124 bytes, 6,784 paragraphs, 142 tables and 379 media members. It contains no content controls or comments. Two ordinary excerpts retain original business and pricing content, relationships and supporting parts; they add no synthetic task text. **Both excerpts come from that same one bid**, so these are not three independent real documents.

Supplemental public fixtures cover an NDA, a tagged content control, and a review memo with four existing comments and tracked changes. Their upstream fixtures are synthetic; they are not customer-document evidence. Features absent from the bid are tested on these applicable supplemental inputs.

Fixed jobs cover precise clause replacement with original run formatting, a price selected by business key/heading, unchanged copies, restarted Save receipt recovery, existing-comment reading, and a tagged text control. The merged price table safely refuses the by-key workflow, then succeeds through an ordinary table-scoped exact text query and guarded `replace_text`. Both cell identity and its unchanged neighbors/merges are independently checked. Original inputs remain unchanged.

The saved-file oracle is separate from OfficeMCP and the SDK. It checks ordered text and effective run-format spans, table rows/grid/merges, control tag/alias/value, header/footer/note text, drawing locations/size/alt/binary references, existing comment anchors/status, and every binary-member hash. All 379 media members survive successful stress jobs. Adversarial tests prove that lost bold formatting, changed comment text/anchors and altered media are rejected. This is **not** full style, arbitrary-object or Microsoft Word layout qualification.

## Actual agents

One generic public MCP + Chat Completions harness ran `deepseek-flash` and `deepseek-v4-pro`, with thinking disabled, once per applicable job/model. It imports no app adapter and supplies credentials through the caller. The user authorized bounded private-bid excerpts for these two models. Each model-visible tool reply stays within 32 KiB and accumulated new tool output within 256 KiB. Oversized or binary replies produce an explicit client error and are withheld entirely; nothing is silently truncated. Four replies across the completed matrix were withheld by this budget and agents recovered with narrower calls.

Every successful save has a public receipt matching its independently hashed output. Save-recovery jobs discard the successful response, restart the CLI and require the same operation's durable receipt. Existing-comment agent jobs require the actual comment reads, matching status, reported comment text and unchanged saved-package evidence. No DOCX binary or image payload was sent to the models.

Observed job wall time includes provider planning and MCP calls. These single-run observations describe usability; they are **not** SDK-relative timing or memory benchmarks.

| Model | Input / job | Calls | Error replies | Job seconds | New tool output KiB |
| --- | --- | ---: | ---: | ---: | ---: |
| deepseek-flash | bid-business / text-edit | 8 | 0 | 13.2 | 13.6 |
| deepseek-flash | bid-business / save-recovery | 6 | 0 | 10.1 | 9.0 |
| deepseek-flash | bid-pricing / table-edit | 19 | 3 | 35.7 | 64.1 |
| deepseek-flash | public-nda / text-edit | 7 | 0 | 11.3 | 11.3 |
| deepseek-flash | public-controls / template-control | 8 | 0 | 11.4 | 11.6 |
| deepseek-flash | public-comments / comment-read | 6 | 0 | 10.1 | 10.8 |
| deepseek-v4-pro | bid-business / text-edit | 7 | 1 | 24.4 | 11.0 |
| deepseek-v4-pro | bid-business / save-recovery | 8 | 1 | 22.7 | 15.5 |
| deepseek-v4-pro | bid-pricing / table-edit | 16 | 4 | 58.3 | 36.8 |
| deepseek-v4-pro | public-nda / text-edit | 7 | 1 | 23.4 | 10.8 |
| deepseek-v4-pro | public-controls / template-control | 6 | 0 | 20.5 | 9.7 |
| deepseek-v4-pro | public-comments / comment-read | 6 | 1 | 18.3 | 11.1 |
| deepseek-flash | bid-stress / text-edit | 5 | 0 | 37.5 | 9.7 |
| deepseek-flash | bid-stress / save-recovery | 8 | 1 | 40.8 | 33.2 |
| deepseek-flash | bid-stress / table-edit | 15 | 3 | 72.4 | 46.2 |
| deepseek-v4-pro | bid-stress / text-edit | 5 | 0 | 44.6 | 12.0 |
| deepseek-v4-pro | bid-stress / save-recovery | 8 | 1 | 55.1 | 19.3 |
| deepseek-v4-pro | bid-stress / table-edit | 18 | 4 | 165.8 | 190.0 |

Both models completed the ordinary merged-price job after several safe refusals. The stress table job used 15–18 calls; the slower v4-pro run took 165.8 seconds and exposed 190.0 KiB. This leaves a clear guidance improvement opportunity: teach the supported table-scoped text fallback sooner. Passing these jobs does not imply that every model attempt was efficient or error-free.

## Reuse and evidence

[Release instructions](RELEASE.md#real-document-docx-qualification) document `docx-corpus-prepare.py`, `docx-corpus-workflows.mjs` and `docx-corpus-agents.mjs`. The prepared manifest binds sources and exact applicable jobs. Use fresh evidence directories, an independently installed package and caller-owned model settings. These runners live in the OfficeMCP source checkout; the production npm runtime does not import them.

`docx-corpus-report.mjs consumer.json corpus-report.json ...` is the permanent fail-closed stored-evidence gate. It checks completion, matching production/SDK identities, every independent saved-file verdict and model byte limits. Negative tests prove that a stored pass flag cannot override a failed/missing independent verdict. The ordinary agent run began before an explicit `evidence.passed` assertion was added inside the agent runner; every stored verdict is true and the permanent gate accepts all 12. Stress runs include the assertion. No false verdict was accepted.

Raw source identities, source hashes, fixed and model traces, receipt evidence and outputs are under `.build/runs/docx-product-corpus/`. The aggregate is final-qualification.json (`../../.build/runs/docx-product-corpus/final-qualification.json`); the independent stored-report gate is accepted-corpus.json (`../../.build/runs/docx-product-corpus/accepted-corpus.json`). Initial comment-write failures on the real bid are archived and superseded by the user's read-only scope, rather than concealed as passing comment-write support.

Remaining qualification limits: macOS arm64 only; one real source and two derived excerpts; supplemental synthetic feature shapes; one agent harness, two models and single observations; no full-app integration, layout/pixel or new baseline-performance acceptance. Integration optimization remains deferred.
