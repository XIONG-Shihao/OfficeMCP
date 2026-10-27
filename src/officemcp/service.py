"""The seven tools, as plain functions over a configuration. The MCP server only exposes these.

Every result has `status`: `ok`, or `refused` with a stable `code`, a message and a hint.
"""

import csv
import io
import subprocess
from pathlib import Path

from . import catalog
from .errors import Refusal
from .fidelity import inventory, losses
from .files import Config, Scratch, format_of, replace, revision, write_new
from .formats import HANDLERS

FORMATS = ["docx", "xlsx", "pptx", "pdf"]
CONVERSIONS = {  # target -> sources OfficeMCP converts itself; PDF comes from LibreOffice
    "txt": {"docx", "xlsx", "pptx", "pdf"},
    "csv": {"xlsx"},
    "pdf": {"docx", "xlsx", "pptx"},
}


def guarded(run):
    """Turn refusals (and unexpected library errors) into results."""
    def call(*args, **kwargs):
        try:
            return {"status": "ok", **run(*args, **kwargs)}
        except Refusal as refusal:
            return refusal.result()
        except Exception as error:  # a library failure or a bug: reported, never hidden
            return {"status": "error", "message": f"{type(error).__name__}: {error}"}
    call.__name__ = run.__name__
    call.__doc__ = run.__doc__
    return call


class Service:
    def __init__(self, config: Config):
        self.config = config

    # -- discovery ------------------------------------------------------------------------
    @guarded
    def actions(self, format: str | None = None, family: str | None = None) -> dict:
        if format is not None and format not in FORMATS:
            raise Refusal("unsupported_format", f"Formats are {', '.join(FORMATS)}")
        found = catalog.actions(format, family)
        return {"actions": [a.entry() for a in found],
                "families": sorted({(a.format, a.family) for a in found}),
                "note": "Use an action in office_edit (or office_create) as {\"action\": name, ...its fields}."}

    # -- reading --------------------------------------------------------------------------
    def _open(self, path: str):
        source = self.config.input(path)
        kind = format_of(source)
        handler = HANDLERS[kind]
        try:
            return source, kind, handler, handler.open(source)
        except Refusal:
            raise
        except Exception as error:
            raise Refusal("unreadable", f"{source.name} could not be opened as {kind}: {error}") from None

    @guarded
    def inspect(self, path: str) -> dict:
        source, kind, handler, document = self._open(path)
        result = {"path": str(source), "format": kind, "revision": revision(source), "bytes": source.stat().st_size,
                  "outline": handler.outline(document)}
        # What a save through the library would lose, found by saving the file unchanged.
        if kind != "pdf":
            with Scratch(self.config) as scratch:
                probe = scratch.file("probe" + handler.extension)
                handler.save(document, probe)
                result["cannotKeep"] = losses(inventory(source, kind), inventory(probe, kind), catalog.Edit(self.config).allowed)
        else:
            result["cannotKeep"] = []
        return result

    @guarded
    def read(self, path: str, address: str | None = None, offset: int = 0, limit: int = 100) -> dict:
        source, kind, handler, document = self._open(path)
        limit = max(1, min(int(limit), 1000))
        if kind == "xlsx":
            body = handler.read(document, address, max(0, int(offset)), limit, self.config.max_cells)
        else:
            body = handler.read(document, address, max(0, int(offset)), limit)
        return {"path": str(source), "format": kind, "revision": revision(source), **body}

    @guarded
    def find(self, path: str, text: str, matchCase: bool = False, limit: int = 100) -> dict:
        if not isinstance(text, str) or not text:
            raise Refusal("invalid_request", "text is required")
        source, kind, handler, document = self._open(path)
        found = handler.find(document, text, matchCase, max(1, min(int(limit), 1000)))
        return {"path": str(source), "format": kind, "revision": revision(source), "matches": found}

    # -- changing -------------------------------------------------------------------------
    def _apply(self, kind, document, actions, creating=False):
        if not isinstance(actions, list) or not actions:
            raise Refusal("invalid_request", "actions must be a non-empty list")
        edit = catalog.Edit(self.config)
        receipts = []
        for index, args in enumerate(actions):
            if not isinstance(args, dict) or not isinstance(args.get("action"), str):
                raise Refusal("invalid_request", f"actions[{index}] needs an action name")
            definition = catalog.lookup(kind, args["action"])
            if definition is None:
                raise Refusal("unknown_action", f"{kind} has no action '{args['action']}'", hint="office_actions lists them", index=index)
            if definition.creates and not creating:
                raise Refusal("invalid_request", f"{args['action']} builds a new PDF; use office_create", index=index)
            problem = validate(definition.schema(), args)
            if problem:
                raise Refusal("invalid_request", f"actions[{index}] ({args['action']}): {problem}", index=index)
            try:
                receipt = definition.run(document, args, edit)
            except Refusal as refusal:
                refusal.details.setdefault("index", index)
                refusal.details.setdefault("action", args["action"])
                raise
            receipts.append({"index": index, "action": args["action"], **receipt})
        return edit, receipts

    def _finish(self, kind, handler, document, edit, receipts, write, dry_run=False, recalculate=False,
                source=None, allow_loss=False):
        """Save, check what the source had against what the save keeps, then write (or only report)."""
        with Scratch(self.config) as scratch:
            draft = scratch.file("draft" + handler.extension)
            handler.save(document, draft)
            if recalculate:
                draft = self._libreoffice(draft, "xlsx", scratch)
            lost = losses(inventory(source, kind), inventory(draft, kind), edit.allowed) if source else []
            if lost and not allow_loss:
                raise Refusal("fidelity_loss", "Saving would lose content the file has; nothing was written",
                              hint="Pass allowLoss: true to save without it, or leave the file as it is", lost=lost)
            result = {"receipts": receipts, "fidelity": {"checked": source is not None, "lost": lost}}
            if edit.warnings:
                result["warnings"] = edit.warnings
            if kind == "xlsx":
                result["recalculationNeeded"] = edit.recalculation_needed and not recalculate
            if dry_run:
                return {**result, "dryRun": True}
            target = write(draft)
            return {**result, "path": str(target), "revision": revision(target)}

    @guarded
    def edit(self, path: str, expectedRevision: str, actions: list, output: dict | None = None,
             dryRun: bool = False, recalculate: bool = False, allowLoss: bool = False) -> dict:
        source, kind, handler, document = self._open(path)
        current = revision(source)
        if expectedRevision != current:
            raise Refusal("revision_mismatch", "The file changed since that revision", hint="Inspect or read it again", revision=current)
        if recalculate and kind != "xlsx":
            raise Refusal("invalid_request", "recalculate applies to workbooks")
        if recalculate and not self.config.libreoffice:
            raise Refusal("engine_unavailable", "Recalculating needs LibreOffice, which was not found")
        target = self.config.new_file(output["path"], handler.extension) if output else None
        edit, receipts = self._apply(kind, document, actions)

        def write(draft):
            if target is None:
                replace(draft, source, current)
                return source
            write_new(draft, target)
            return target
        return {"sourceRevision": current, **self._finish(kind, handler, document, edit, receipts, write, dryRun,
                                                          recalculate, source, allowLoss)}

    @guarded
    def create(self, format: str, path: str, actions: list | None = None) -> dict:
        if format not in FORMATS:
            raise Refusal("unsupported_format", f"Formats are {', '.join(FORMATS)}")
        handler = HANDLERS[format]
        target = self.config.new_file(path, handler.extension)
        document = handler.blank()
        if format == "pdf" and not actions:
            raise Refusal("invalid_request", "A new PDF needs compose actions (add_heading, add_paragraph, add_table, add_page_break)")
        edit, receipts = self._apply(format, document, actions, creating=True) if actions else (catalog.Edit(self.config), [])
        return self._finish(format, handler, document, edit, receipts, lambda draft: (write_new(draft, target), target)[1])

    @guarded
    def convert(self, path: str, to: str, output: dict | None = None, sheet: str | None = None) -> dict:
        source, kind, handler, document = self._open(path)
        if to not in CONVERSIONS or kind not in CONVERSIONS[to]:
            raise Refusal("unsupported_format", f"{kind} does not convert to {to}",
                          hint="txt from any format, csv from xlsx, pdf from docx, xlsx or pptx (LibreOffice)")
        target = self.config.new_file(output["path"], f".{to}") if output else self.config.beside(source, f"{source.stem}.{to}")
        with Scratch(self.config) as scratch:
            draft = scratch.file("converted." + to)
            if to == "txt":
                draft.write_text(handler.text(document), encoding="utf-8")
            elif to == "csv":
                workbook = document.workbook
                if sheet and sheet not in workbook.sheetnames:
                    raise Refusal("target_not_found", f"No sheet '{sheet}'")
                ws = workbook[sheet] if sheet else workbook.worksheets[0]
                values = document.values[ws.title] if document.values is not None else ws
                buffer = io.StringIO()
                writer = csv.writer(buffer)
                for row in values.iter_rows(values_only=True):
                    writer.writerow(["" if v is None else v for v in row])
                draft.write_text(buffer.getvalue(), encoding="utf-8")
            else:
                if not self.config.libreoffice:
                    raise Refusal("engine_unavailable", "Converting to PDF needs LibreOffice, which was not found")
                draft = self._libreoffice(source, "pdf", scratch)
            write_new(draft, target)
        return {"path": str(target), "revision": revision(target), "from": kind, "to": to}

    def _libreoffice(self, source: Path, to: str, scratch: Scratch) -> Path:
        out = scratch.path / f"lo-{to}"
        out.mkdir()
        profile = scratch.path / "lo-profile"
        run = subprocess.run([str(self.config.libreoffice), f"-env:UserInstallation=file://{profile}", "--headless",
                              "--convert-to", to, "--outdir", str(out), str(source)], capture_output=True, text=True, timeout=300)
        produced = list(out.iterdir())
        if run.returncode != 0 or len(produced) != 1:
            raise Refusal("engine_unavailable", f"LibreOffice did not convert the file: {run.stderr.strip()[:500]}")
        return produced[0]


def validate(schema: dict, value) -> str | None:
    """A small JSON-schema check for action arguments (types, enums, required and unknown fields)."""
    from jsonschema import Draft202012Validator

    errors = sorted(Draft202012Validator(schema).iter_errors(value), key=lambda e: list(e.path))
    if not errors:
        return None
    first = errors[0]
    where = "/".join(str(p) for p in first.path)
    return f"{where + ': ' if where else ''}{first.message}"
