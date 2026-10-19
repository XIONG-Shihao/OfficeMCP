# Portable OfficeMCP scope

OfficeMCP provides local stdio MCP tools for DOCX, XLSX, CSV, PPTX and PDF.
It owns tool discovery, document operations, revision/candidate handles,
validation, durable receipts, quotas and safe publication. Its CLI and optional
JavaScript SDK work without a desktop application, app account or app database.

Required headless document engines are built from external sources and included
in assembled runtime packages; their sources stay outside this repository. Editor UI,
standalone editor shells, desktop app packaging, workspace authorization,
agent harnesses, app settings and viewer integration do not belong here.
Hosts configure ordinary read/write roots and private state through the public
configuration described in [Usage](USAGE.md). Hosts own their additional access
controls and how they display or adopt saved documents.

Prioritize useful document workflows, correct edits, preservation, bounded
responses and reliable recovery. Comments remain read-only. DOCX page-based
operations use the explicit optional layout provider; provider-specific results
must not imply universal Word layout compatibility. SDK internals are not patched.

Qualification runs through public stdio clients and independent saved-document
checks, including an installation outside this source tree. Live-model tests
use caller-supplied credentials and explicit opt-in. A successful synthetic test
does not prove arbitrary-model reliability or document appearance. GUI checks
and combined app journeys belong to the consuming application repository.

macOS arm64 is the currently exercised native target. Other platforms require
separate qualification. See [Release verification](RELEASE.md) and
[Repository boundary](REPOSITORY.md).

## Contract and state transition

The current package is `0.18.0-preview.1`, core contract `0.13`, DOCX contract
`docx/6`, with private-state schema 3. Removed tools and invalid arguments fail
before document mutation. Comments remain read-only.

Supported state upgrades run transactionally. Older executables may reject a
newer state schema; preserve a private-state backup before attempting a downgrade.
Revision-bound receipts cannot authorize a different document candidate. Save
requires validation for the exact candidate being published. Migration and
cleanup must preserve source documents and published outputs.
