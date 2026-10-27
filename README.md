# OfficeMCP

An MCP server that gives language models standardized, predictable operations on PDF, Word (DOCX),
Excel (XLSX) and PowerPoint (PPTX) files. It contains no document engine: published Python
libraries read and write the files (`python-docx`, `openpyxl`, `python-pptx`, `pypdf`, `reportlab`),
and OfficeMCP is the operation catalog and the MCP layer over them. The design and its reasons are
in the Agent Platform repository's `docs/officemcp-simplification.md`.

## Run it

```sh
uv sync --frozen                                   # Python 3.12 and the locked libraries
uv run python -m officemcp serve --config config.json
```

The server speaks MCP over stdio. The configuration says where it may read and write:

```json
{
  "roots": ["/path/to/project"],
  "engines": {"libreOffice": "/Applications/LibreOffice.app/Contents/MacOS/soffice"},
  "limits": {"maxBytes": 209715200, "maxCells": 5000}
}
```

OfficeMCP reads and writes files inside `roots`; a relative path starts at the first root. It is a
tool like a script the model runs: it has no sandbox of its own, and the host that starts it decides
what the process may touch. Temporary files go to the system's temporary folder.
`engines.libreOffice` is optional: without it OfficeMCP looks in the usual install locations, and
requests that need LibreOffice (recalculating formulas, converting to PDF) are refused with
`engine_unavailable` when it is not found.

## Tools

| Tool | Does |
|---|---|
| `office_actions` | Lists the actions, by format and family, with their fields |
| `office_inspect` | Outlines a file with addresses, gives its revision, and lists what its library cannot keep |
| `office_read` | Reads the content at an address, or pages through the file |
| `office_find` | Finds text; returns addresses |
| `office_edit` | Applies actions, all or none, and saves the file in place, or as a new file with `output.path`; `dryRun` checks without writing |
| `office_create` | Creates a file at a new path and applies actions to it (a new PDF from compose actions) |
| `office_convert` | Converts to txt, to csv (a workbook sheet), or to pdf (through LibreOffice), into a new file |

Every result has `status`: `ok`, `refused` with a stable `code`, a message and a hint, or `error`
for an unexpected failure. The actions, addresses and refusal codes are listed in the generated
agent guide, [agent/OfficeMCP.md](agent/OfficeMCP.md); the generated
[contracts/catalog.json](contracts/catalog.json) carries every schema for a host's settings page.

## What an edit promises

- **Addresses and expected values.** Actions take addresses that `inspect`, `read` and `find`
  return (`body/p[3]`, `Budget!B2:D9`, `slide[2]/shape[5]`, `page[4]`), and the content they
  expect there. Different content refuses with `expected_mismatch`.
- **Revisions.** `office_edit` takes the file's revision (the SHA-256 of its bytes) and refuses if
  the file changed (`revision_mismatch`), checked again just before the file is replaced.
- **All or none.** A batch either applies completely or writes nothing. The file is replaced in one
  rename and keeps its permissions; new files (`output.path`, `office_create`, `office_convert`)
  never overwrite an existing one.
- **The fidelity check.** Before writing, OfficeMCP compares what the input and the output contain:
  parts by content type, binary parts by hash, and features a library can drop (charts, drawings,
  tracked changes, extensions, form controls, PDF forms, outlines, attachments, signatures, ...).
  Anything lost that no action declared refuses the edit with `fidelity_loss` and the list. The
  model decides: with `allowLoss: true` the edit is saved and the result lists what was lost
  (`fidelity.lost`). `office_inspect` reports the same check for an unchanged save (`cannotKeep`),
  so a model knows before editing.
  Bookkeeping that Office rebuilds or only records how a file was made is not counted (the
  calculation chain, shared-string tables, theme family names, default slicer and timeline styles,
  the folder a workbook was last saved in).

## Develop

```sh
uv run python -m unittest                          # the tests, on files they generate
uv run python -m officemcp contracts               # rewrite agent/ and contracts/ after changing an action
```

An action is declared once, in its format's module under `src/officemcp/formats/`, with the
`@action` decorator: its name, family, summary, fields and the function that carries it out. The
tools, the guide and the catalog follow from that. `tests/test_contracts.py` fails until the
generated files are rewritten. From the app repository, `node scripts/test/run-parallel.mjs --suite
officemcp` runs the tests in an environment made with the app's pinned uv and Python.

## Not covered

- LibreOffice paths (recalculating, converting to PDF) are tested only where LibreOffice is
  installed; elsewhere their tests skip and only the refusal is tested.
- Tests use files the libraries generate; real-world files are checked by hand. On the XLSX
  editor's 29 test workbooks, 21 are kept whole by `openpyxl`; the others hold dynamic arrays,
  images in cells or comment formatting that `openpyxl` drops, so editing them needs `allowLoss`.
- The Full App's integration is tested from source (`tests/officemcp-installed.mjs` in the app
  repository); a built Full App package has not been verified.
