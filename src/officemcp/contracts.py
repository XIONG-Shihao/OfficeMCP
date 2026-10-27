"""Files generated from the catalog: the settings page's catalog and the agent guide.

`python -m officemcp contracts` writes them; `--check` fails when they are out of date.
"""

import json
from importlib.metadata import version
from pathlib import Path

from . import catalog
from .errors import CODES
from .formats import LIBRARIES
from .server import VERSION
from .tools import RESULT, TOOLS

ROOT = Path(__file__).resolve().parents[2]
CATALOG = ROOT / "contracts" / "catalog.json"
GUIDE = ROOT / "agent" / "OfficeMCP.md"
RECEIPT = {"type": "object", "properties": {"changed": {"type": "array", "items": {"type": "string"}, "description": "Addresses the action changed."}}}

ADDRESSES = {
    "docx": "body/p[N], body/table[N], body/table[N]/r[R]/c[C], body/table[N]/r[R]/c[C]/p[K]",
    "xlsx": "Sheet, Sheet!B2, Sheet!B2:D9, Sheet!B:D, table[Name]",
    "pptx": "slide[N], slide[N]/shape[ID], slide[N]/shape[ID]/r[R]/c[C], slide[N]/notes",
    "pdf": "page[N], form/field[Name]",
}
LIMITATIONS = {
    "docx": "Text inside hyperlinks, fields and content controls is not changed; tracked changes are not produced.",
    "xlsx": "openpyxl does not calculate formulas: edits report recalculationNeeded, and recalculate needs LibreOffice. Rows are not inserted or deleted in the middle of a sheet. Editing a workbook whose content openpyxl cannot keep (for example some charts or form controls) needs allowLoss.",
    "pptx": "Text in SmartArt and charts is not changed; new slides come only from the presentation's layouts.",
    "pdf": "Text in existing pages is not edited; removing pages that outline entries or form fields point to is refused. Converting Office files to PDF needs LibreOffice.",
}


def library(format: str) -> str:
    name, _ = LIBRARIES[format]
    return f"{name} {version(name)}"


def catalog_json() -> dict:
    layers = [{
        "id": "office", "label": "Tools", "engine": "OfficeMCP " + VERSION,
        "summary": "Seven tools, the same for every format: discover actions, inspect, read, find, edit, create, convert.",
        "limitations": "Edits save in place after a revision check; new files never overwrite.",
        "missing": [], "proposedEngine": "", "evaluation": {"status": "Implemented", "summary": "Tested on generated documents of every format."},
        "operations": [{
            "name": t["name"], "description": t["description"], "inputSchema": t["inputSchema"], "outputSchema": RESULT,
            "annotations": {"readOnlyHint": t["readOnly"]}, "resultType": "result", "wireTypes": ["json"], "status": "Implemented",
        } for t in TOOLS],
    }]
    for format in ("docx", "xlsx", "pptx", "pdf"):
        found = catalog.actions(format)
        layers.append({
            "id": format, "label": format.upper(), "engine": library(format) + (" and reportlab " + version("reportlab") if format == "pdf" else ""),
            "summary": f"{len(found)} actions. Addresses: {ADDRESSES[format]}.",
            "limitations": LIMITATIONS[format], "missing": [], "proposedEngine": "",
            "evaluation": {"status": "Implemented", "summary": "Each action is tested on a generated file, with its refusals."},
            "operations": [{
                "name": f"{format}_{a.name}", "description": a.summary, "inputSchema": a.schema(), "outputSchema": RECEIPT,
                "annotations": {"readOnlyHint": False}, "resultType": "receipt", "wireTypes": ["json"], "status": "Implemented",
                "preset": a.family, "contractVersion": a.tier,
            } for a in found],
        })
    return {"version": VERSION, "contractVersion": "office/1", "transport": "stdio",
            "status": "Actions over published libraries; edits report what a save would lose.", "layers": layers}


def guide_markdown() -> str:
    lines = [
        "# OfficeMCP", "",
        "OfficeMCP works on PDF, Word (DOCX), Excel (XLSX) and PowerPoint (PPTX) files with standardized actions.",
        "Every tool returns `status`: `ok`, or `refused` with a `code`, a `message` and usually a `hint`.", "",
        "## Workflow", "",
        "1. `office_inspect` the file: note its `revision`, the addresses in its outline, and `cannotKeep`.",
        "2. `office_read` (an address, or pages of the whole file) and `office_find` (text) for what you need.",
        "3. `office_actions` lists the actions for the format, with their fields.",
        "4. `office_edit` with the file, its `expectedRevision` and a list of actions. All apply or none does;",
        "   the file is saved in place (or as a new file with `output.path`), with a receipt per action.",
        "   Use `dryRun` to check first.",
        "5. Read the receipts. An action that names `expectedText` or `expected` refuses when the content differs;",
        "   read the target again instead of guessing. Check the result with `office_read` when it matters.", "",
        "Paths are absolute or relative to the project folder. `office_create` and `office_convert` write new files",
        "and never overwrite one.", "",
        "## Content a save would lose", "",
        "Before writing, OfficeMCP compares what the file had with what the save keeps. If something would be lost",
        "(a chart, an image, tracked changes, ...), the edit is refused with `fidelity_loss` and the list in",
        "`details.lost`. Decide: leave the file, save as a new file, or repeat the edit with `allowLoss: true`;",
        "the result then lists what was lost in `fidelity.lost`. Tell the person what was lost.", "",
        "## Addresses", "",
        *[f"- {f.upper()}: {a}" for f, a in ADDRESSES.items()],
        "", "Ordinals count from 1.", "",
        "## Tools", "",
        *[f"- `{t['name']}`: {t['description']}" for t in TOOLS],
        "", "## Actions", "",
    ]
    for format in ("docx", "xlsx", "pptx", "pdf"):
        lines += [f"### {catalog.FORMAT_NAMES[format]}", "", LIMITATIONS[format], ""]
        for a in catalog.actions(format):
            fields = ", ".join(f"`{k}`" + ("" if k in a.required else "?") for k in a.params)
            lines.append(f"- `{a.name}` ({a.family}{', create only' if a.creates else ''}): {a.summary} Fields: {fields or 'none'}.")
        lines.append("")
    lines += ["## Refusal codes", "", *[f"- `{code}`: {meaning}" for code, meaning in CODES.items()], ""]
    return "\n".join(lines)


def write(check: bool = False) -> list[str]:
    """Write the generated files, or with check=True list those that differ from what would be written."""
    wanted = {CATALOG: json.dumps(catalog_json(), indent=2, ensure_ascii=False) + "\n", GUIDE: guide_markdown()}
    stale = [str(p.relative_to(ROOT)) for p, text in wanted.items() if not p.exists() or p.read_text(encoding="utf-8") != text]
    if not check:
        for path, text in wanted.items():
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(text, encoding="utf-8")
    return stale
