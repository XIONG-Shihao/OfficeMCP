# Capability-first product principles

OfficeMCP supplies portable MCP document tools to agents. Independent public-MCP
clients and isolated package consumers own product qualification. The integrated
app is one consumer and owns its workspace publication and viewer acceptance.
Follow [SCOPE.md](SCOPE.md) for this boundary and metric-specific cost attribution.
Prefer scripted GUI tests for repeatability; user-authorized Computer Use is also
permitted. The first question for every proposed change is:

> What useful document job can an agent complete for a user after this change
> that it could not complete reliably before?

Qualification, security and conformance still matter, but they are release gates
around useful capabilities. They are not substitutes for those capabilities.

## Priority order

Work in this order:

1. **User outcome coverage.** Cover frequent, valuable document jobs from start
   to finish.
2. **Agent usability.** Make the job discoverable and achievable with a short,
   obvious tool sequence and compact results.
3. **Workflow correctness.** Confirm the requested change happened, unrelated
   content survived, and the output is usable.
4. **Portable product and consumer performance.** Measure fixed public-MCP
   workflows first, then each supported host separately. Report SDK work, product
   orchestration, transport, app publication and viewer costs separately. Use
   metric-specific budgets; do not select a faster path by host/model identity.
5. **Qualification completeness.** After the operation and response contract are
   stable, add the full action-specific evidence and release qualification.

Data-loss prevention, no-overwrite publication, explicit failure, bounded
resource use and no silent success are exceptions: they remain non-negotiable
while capabilities are being built because their absence directly harms users.
Internal labels such as experimental or qualified do not belong in normal agent
responses. Functional facts that change how an agent must work do belong there.
For example, an agent must know that DOCX mutation targets identify semantic
content, not editor page numbers; it does not need an internal test-maturity label.

## Think in user jobs, not SDK methods

A tool is only useful when it contributes to a complete user job. An SDK method
is an implementation option, not a reason to expose an action. Prefer a small
number of capable operations over one tool for every SDK method.

Describe every capability with five concrete facts:

1. **User request:** the sentence a person would actually say.
2. **Agent route:** the exact tools and actions the agent should call.
3. **Targeting rule:** how the agent identifies the correct document object.
4. **Returned evidence:** the small result that tells the agent what changed and
   what to do next.
5. **Acceptance:** observable facts in the saved document, including preservation
   of unrelated content.

Do not describe a feature merely as “image support” or “table support.” Describe
what the agent can do:

| User job | Agent route | Useful returned evidence | Acceptance |
|---|---|---|---|
| Understand the whole document | `docx_open` → `docx_inspect` → paginated `docx_read` across the required stories and domains | Outline, inventory, bounded items, coverage and the next unread range | The agent can answer from body text, tables, headers, footers, images and review content without treating unread areas as empty |
| Change one clause without losing formatting | `docx_query` for the paragraph/span → `docx_apply` with `replace_text` or `rewrite_block` → `docx_validate` → `docx_save` | Resolved target count, new revision, postcondition and output path | The intended occurrence changed, its required formatting remains, unrelated text and objects remain, and a new file opens |
| Update a table by meaning | `docx_query` with table headers/content → `docx_apply` with `update_table` → read the affected table → validate/save | Table target, changed cells and compact readback | The intended row/cell changed; headers, other rows and surrounding content remain |
| Add or change an image | Resolve a paragraph or existing image → `docx_apply` with `insert_image`, `replace_image`, `resize_image`, layout or alternative-text action → retrieve original image bytes when needed | Image target, dimensions, placement, media identity and independent package evidence | Correct image bytes and metadata are present, its anchor is correct, and surrounding layout is acceptable |
| Edit content outside the body | Read/query the relevant header or footer story → materialize inherited content if necessary → apply → validate/save | Story and section identity plus changed target | The intended section/story changed and the document body did not |
| Review a document | Read existing comments, or query and accept/reject a precise tracked change → read review domains → save | Existing comment text/status and tracked-change identity/state | Comments are preserved; a chosen tracked change has the requested state after reopen |
| Deliver a user artifact | Validate/compare; `docx_save` or `docx_export` | Output path, validation summary and package evidence | The user receives a new DOCX/PDF that opens and contains the requested result |

### “Insert an image on page X”

DOCX stores flowing content, not stable page objects. Page numbers belong to the
layout provider and can change after insertion. An optional provider can bind a
page's unique literal text to a semantic mutation target through `docx_layout`.
Regenerate that layout after mutation. A page number alone cannot identify an
image insertion target; resolve the paragraph, table or image semantically.
Display acceptance in a consuming application belongs to its separate scripted
GUI tests and is not a condition supplied through MCP validation arguments.

## DOCX evaluation contract

DOCX evaluation has two complementary fixture classes:

- `test_doc/response.docx` is the required realistic fixture. It is approximately
  234 MB, with an approximately 8 MB main document XML and 379 media files. Use
  it to expose slow reads, excessive payloads, lost content, weak targeting,
  pagination problems and memory pressure.
- Small generated fixtures isolate exact semantics such as comments, tracked
  changes, controls and a particular table shape. They make failures explainable.
  They must not replace the realistic fixture.

A feature is not considered good for agents merely because its direct action test
passes. It must also participate in at least one end-to-end user job. Conversely,
`response.docx` alone cannot prove every feature because it does not contain every
Word construct in a controlled form.

### Required `response.docx` workflows

Run each workflow from a fresh open and publish outputs only under
`test_doc/officemcp-results/`:

| Workflow | What the agent must do | What the test measures |
|---|---|---|
| Fast orientation | Identify language, title/project, major sections, table count, image count, headers/footers and a reading plan without reading the whole payload | Time to first useful answer, calls, bytes returned, inventory correctness |
| Complete targeted retrieval | Find a unique named section and return its surrounding paragraphs and related table without scanning every item through the model | Target accuracy, tool calls, transferred bytes, coverage honesty |
| Cross-domain read | Answer a question that requires body text plus a table or footer and cite the returned targets | Missed stories/domains, pagination correctness, answer completeness |
| Guarded text edit | Change one unique literal in a copied candidate while preserving its run formatting and every unrelated package member | Target precision, edit time, semantic readback, preservation |
| Table edit | Change one table cell selected by header/content rather than array position | Query quality, table readback, preservation of adjacent cells and formatting |
| Image workflow | Inspect one existing image, retrieve its bytes, update its alternative text or size, then validate the package | Image identity, byte preservation, package evidence, response size |
| Insert near known content | Insert a test image before/after a stable paragraph target and verify semantic placement after reopen | Calls, target stability, image metadata, surrounding-content preservation |
| Save and reopen | Validate and publish a new filename, reopen it, and repeat the relevant targeted reads | Publication time, output usability, source immutability, result persistence |
| Recovery | Recover the result of an intentionally uncertain long mutation by operation ID instead of repeating it | Duplicate prevention, useful recovery response, total recovery time |

Do not mutate the supplied `response.docx`. Every mutation starts from a private
candidate or copy and uses a unique output name.

### Measurements

Record these per workflow, not only for the whole suite:

- successful user outcome: yes/no, with the failed acceptance fact;
- elapsed time to open, first useful read, target resolution, mutation,
  validation and publication;
- peak resident memory and output file size;
- tool calls, repeated calls, invalid calls and recovered errors;
- JSON bytes returned to the model and native artifact/image bytes returned
  separately;
- number of pages/items inspected versus the number actually needed;
- preservation checks for text, styles, tables, relationships and media;
- whether the agent could determine the next call from each response without
  reading internal documentation.

Establish a measured baseline before choosing hard performance limits. After a
baseline exists, fail continuous tests on material regressions instead of using
arbitrary thresholds. A faster result that misses a footer or corrupts an image
is a failure, not an optimization.

## Capability coverage, not action counts

Track coverage by user job. Action counts are implementation detail and can
reward fragmentation. The primary dashboard should answer:

- What can a user ask for today?
- Which exact tool route should an agent use?
- Does it work on a realistic document?
- How many calls, retries, seconds and model-visible bytes does it cost?
- What functional gap prevents completion when it fails?

Prioritize new work roughly by:

`frequency × user value × agent difficulty removed ÷ implementation cost`

Use failure severity as a release decision, not as the definition of product
value. Frequent workflows such as complete reading, precise text edits, table
updates, image placement, header/footer work and reliable save/export outrank
rare SDK actions with no demonstrated user job.

## Implementation loop

Keep four evidence levels separate: action availability, SDK round-trip results,
independent package invariants and separate application GUI acceptance. A schema or
shared validation path is not action-specific proof. Report exactly which
fixture, semantic assertions, package checks and, separately, application views were exercised;
retain prior measurements as historical evidence when the implementation changes.

Grouped discovery is an interface improvement, not a performance result by
itself. `docx_inspect({view, actionFamily: "images"})` supplies the image action
schemas in one bounded request. Prove reduced calls, repeated requests or
returned bytes with comparable live runs before claiming the agent workflow is
easier or faster.

For each capability slice:

1. write one real user request and its saved-document acceptance facts;
2. attempt it with the current MCP and record the exact tool trace;
3. identify whether the blocker is missing data, targeting, mutation, feedback,
   layout evidence or performance;
4. improve the smallest coherent group of tools/actions that completes the job;
5. run the workflow on `test_doc/response.docx` when the fixture contains the
   needed feature, plus a focused fixture for exact assertions;
6. reduce calls, retries and response noise after correctness is established;
7. freeze the agent-facing contract only after the workflow is good to use;
8. complete action-specific qualification, adversarial and release checks before
   claiming production readiness.

Reject work that only increases exposed action count, adds an abstraction without
unlocking a user job, or tests only for a success boolean. An operation is done
when an agent can discover it, call it correctly, observe what happened, recover
from an actionable failure, and deliver a document that satisfies the user's
request.
