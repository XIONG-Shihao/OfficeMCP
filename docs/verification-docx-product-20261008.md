# Standalone DOCX product qualification — 2026-10-08

Historical preview-1 synthetic evidence. The current `docx/5` product makes comments read-only; [current corpus qualification](verification-docx-corpus-20261008.md) supersedes this report’s comment-write scope.

The authorized sequence is complete in the scope below: reusable installation
test → fixed saved-document jobs → actual model-driven jobs. OfficeMCP remains a
private preview; this is useful qualification evidence, not a general DOCX or
cross-platform fidelity claim.

## Permanent commands

Run from the OfficeMCP source checkout:

```sh
npm run test:docx:consumer
npm run test:docx:workflows
npm run test:docx:agents
```

[Release instructions](RELEASE.md) document offline caches, package/source
selection, output directories, opt-in provider credentials and model settings.
The consumer gate builds a tarball and installs it outside the repository as an
ordinary npm dependency. No lifecycle scripts, consumer overrides, app adapter,
Electron host or root-app code is required. The agent runner accepts caller-owned
credentials/configuration and discovers schemas through the public MCP connection.
Only synthetic task documents are used.

## Final results

| Gate | Result |
| --- | --- |
| Fresh offline npm dependency installation | Pass |
| Installed production source/contracts/core identity | Matches the current source |
| Installed SuperDoc SDK/native identity | SDK 2.2.1; native binary matches source |
| Eight fixed workflow variants × three clients | 24/24 pass |
| Blank create/save through each client | 3/3 pass |
| DOCX/protocol/resource regressions | 59 pass; one optional LibreOffice conversion skipped |
| Additional safe-refusal oracle regressions | Two pass |
| Six synthetic jobs × two live models | 12/12 independently accepted |
| Fresh safe-refusal rerun with both models | 2/2 pass |

Clients: independent raw `2026-07-28`, independent raw `2025-11-25`, and official
modern MCP client. This permanent gate does not replace the earlier six-profile
matrix's separate optional-Tasks/older-client evidence.

Live models: `deepseek-flash` and `deepseek-v4-pro`, with thinking disabled, through
one generic Chat Completions tool loop. One observation per job/model is not a
population success rate or evidence for Codex/OpenCode/other agent harnesses.

The fixed jobs cover formatted clause replacement, table-by-key update, a real
text content control, comment addition/resolution, Save recovery after restart,
ambiguous refusal, unique unscoped replacement and a paragraph-domain read target.
The first text job also refuses an old target; the table job refuses a wrong
existing cell value. Every successful job saves and reopens its output. The Save
case intentionally discards a completed response, restarts the CLI, recovers its
durable receipt, and verifies that retries produce the same exact output.

The independent saved-package oracle checks all top-level paragraph text and
bold/italic characters, table-cell values, control tag/type/value, header/footer
and note paragraphs, all binary hashes, body-image placement/size/alt/reference
bytes, and comment anchoring/resolution. Reserved separator notes with IDs -1/0
are excluded. These checks do not establish arbitrary style/layout fidelity.

## Product fixes found by the tests

- Missing or ambiguous exactly-one text/node selections now return actionable
  `TARGET_CONFLICT` errors instead of a generic SDK engine exception.
- A unique unscoped text replacement uses the same range-preserving route as a
  targeted replacement. The first live matrix caught a changed number losing bold
  formatting on the former SDK action route. Multi-match/batch behavior retains
  its existing guarded path; this does not qualify every rich replacement shape.
- Paragraph-domain reads now supply a block target that agent actions accept,
  matching the target behavior of linear/query reads. Legacy span identity is kept.
- Public tool descriptions explain table selection and unchanged-copy creation
  before Save. These descriptions apply to every MCP consumer.
- A macOS process-counter snapshot race caused intermittent failures between
  operations. The product retries missing live-root counters at most three times,
  with 25 ms pauses. Persistently missing data, permission errors and real budget
  violations still fail closed. Tests include real SDK over-budget refusal,
  orphaned owned-group accounting and full worker cleanup.

SDK internals/native executable, MCP input/output shapes and app integration were
not redesigned. The optional publication-descriptor addition from the preceding
work is now explicitly checked in the frozen-contract compatibility test, while
all older fields retain their historical hashes. The batch preservation assertion
now checks Word text across XML runs rather than requiring contiguous XML bytes.

## Live observations

Each cell below is **tool calls / agent-loop seconds / tool error replies**.
Timing starts after initial MCP connection/discovery and excludes installation and
fixture preparation. It includes model/provider work and is one observation, not
a latency target.

| Job | DeepSeek Flash | DeepSeek V4 Pro |
| --- | --- | --- |
| Formatted clause | 12 / 17.6 / 0 | 8 / 34.3 / 0 |
| Table by key | 11 / 18.8 / 1 | 8 / 28.0 / 3 |
| Template control | 10 / 14.9 / 0 | 5 / 15.2 / 0 |
| Comment review | 10 / 15.5 / 0 | 6 / 15.1 / 0 |
| Save recovery | 4 / 6.9 / 0 | 4 / 12.5 / 0 |
| Ambiguous refusal | 3 / 6.9 / 1 | 3 / 12.4 / 0 |

The ambiguity error is an expected safe refusal. Table errors were recovered;
successful saved outcomes do not imply error-free or minimal-call agent journeys.
Measured MCP tool time was 0.44–1.23 seconds per job; most agent-loop time occurred
outside those calls. Raw reports distinguish provider time, MCP time, input JSON,
model-visible response text bytes and usage. Text bytes are not tokens, native
image bytes or total transport bytes. No new SDK-relative memory/timing gate was
assessed in this workflow qualification.

## Evidence and acceptance correction

Final qualification (`../../.build/runs/docx-product/final-qualification.json`)
binds source identity, native SDK identity, the tarball hash and per-case results.
The final tarball SHA-256 is
`00b3c4832710d5b9992cdb0c747a3542b0e988969b3b7d750294469f7255293a`.

- Installed consumer (`../../.build/runs/docx-product/release-verified/consumer.json`)
- Fixed workflows (`../../.build/runs/docx-product/release-verified/installed-workflows/workflows.json`)
- Regressions (`../../.build/runs/docx-product/regressions-observation-final.log`)
- Additional oracle checks (`../../.build/runs/docx-product/oracle-final-complete.log`)
- Raw final model matrix (`../../.build/runs/docx-product/agents-release-verified/agents.json`)
- Fresh refusal checks (`../../.build/runs/docx-product/agents-refusal-recheck/agents.json`)

The raw final model report retains one failed classification: the model had read
two distinct matching paragraphs and correctly refused without editing or saving,
but the version-2 oracle accepted preflight evidence only from a query. Version 3
also accepts two distinct paragraph IDs from canonical reads, rejects repeated
reads of one item and rejects accepted edits. All twelve outcomes were rechecked
against retained inputs/outputs; both refusal jobs were rerun. Earlier failed
installation, test and agent runs remain in the evidence directory.

## Remaining scope

Next, apply these gates to representative real documents and appropriate stress
jobs. Improve excessive reads and table-workflow discovery using retained traces.
Other operating systems, other agent harnesses, rich style/numbering definitions,
complex nested tables/review ranges, repeated reliability samples and performance
tails remain separately unqualified. Rendering and integrated-app optimization
remain separate work.
