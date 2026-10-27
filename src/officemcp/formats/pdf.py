"""PDF files through pypdf (reading and changing existing files) and reportlab (creating new ones).

Addresses: `page[N]` and `form/field[Name]`. Actions on an existing PDF change pages, forms and
metadata; the compose actions (`add_heading`, `add_paragraph`, `add_table`, `add_page_break`)
build a new PDF with office_create.
"""

from pathlib import Path

from pypdf import PdfReader, PdfWriter
from pypdf.generic import ArrayObject, NameObject, NumberObject

from ..catalog import action, address, text
from ..errors import Refusal
from .common import excerpt, pick, steps

PAGES = {"type": "array", "minItems": 1, "maxItems": 10_000, "items": {"type": "integer", "minimum": 1}, "description": "Page numbers, from 1."}


class Pdf:
    """An existing PDF opened for reading and changing, or a new one being composed."""

    def __init__(self, path: Path | None = None):
        self.path = path
        self.reader = PdfReader(path) if path else None
        self.writer = PdfWriter(clone_from=path) if path else None
        self.blocks = None if path else []  # compose mode

    def require_existing(self, name: str):
        if self.writer is None:
            raise Refusal("invalid_request", f"{name} changes an existing PDF; a new PDF is built with the compose actions")

    def require_compose(self, name: str):
        if self.blocks is None:
            raise Refusal("invalid_request", f"{name} builds a new PDF with office_create; it does not change an existing one")


def page_numbers(numbers, count, name):
    for n in numbers:
        if not 1 <= n <= count:
            raise Refusal("target_not_found", f"{name}: there is no page {n} (there are {count})")
    return numbers


def outline_pages(reader: PdfReader) -> list[tuple[str, int]]:
    found = []

    def walk(items):
        for item in items:
            if isinstance(item, list):
                walk(item)
            else:
                try:
                    found.append((str(item.title), reader.get_destination_page_number(item) + 1))
                except Exception:
                    found.append((str(getattr(item, "title", "")), 0))

    try:
        walk(reader.outline)
    except Exception:
        pass
    return found


def field_pages(reader: PdfReader) -> dict[str, int]:
    pages = {}
    for n, page in enumerate(reader.pages, 1):
        for annot in page.get("/Annots") or []:
            annot = annot.get_object()
            name = annot.get("/T") or (annot.get("/Parent") or {}).get("/T") if annot.get("/Subtype") == "/Widget" else None
            if name:
                pages[str(name)] = n
    return pages


def reorder(pdf: Pdf, order: list[int], name: str, edit) -> None:
    """Keep only the listed pages, in that order, refusing to strand outline entries or form fields."""
    count = len(pdf.reader.pages)
    page_numbers(order, count, name)
    if len(set(order)) != len(order):
        raise Refusal("invalid_request", f"{name}: a page may appear once")
    removed = set(range(1, count + 1)) - set(order)
    stranded = [title for title, page in outline_pages(pdf.reader) if page in removed]
    fields = [f for f, page in field_pages(pdf.reader).items() if page in removed]
    if stranded or fields:
        raise Refusal("unsupported_content", f"{name}: removing those pages would strand outline entries or form fields",
                      outlineEntries=stranded[:20], formFields=fields[:20])
    edit.allow("pdf:pages", len(removed))
    for n in removed:
        edit.allow("pdf:annotations", len(pdf.reader.pages[n - 1].get("/Annots") or []))
    root = pdf.writer.root_object["/Pages"]
    refs = [pdf.writer.pages[n - 1].indirect_reference for n in order]
    root[NameObject("/Kids")] = ArrayObject(refs)
    root[NameObject("/Count")] = NumberObject(len(refs))
    for ref in refs:
        ref.get_object()[NameObject("/Parent")] = root.indirect_reference


class Handler:
    format = "pdf"
    extension = ".pdf"

    def open(self, path):
        try:
            reader = PdfReader(path)
            if reader.is_encrypted:
                raise Refusal("unreadable", "The PDF is encrypted")
            return Pdf(path)
        except Refusal:
            raise
        except Exception as error:
            raise Refusal("unreadable", f"Not a PDF pypdf can open: {error}") from None

    def blank(self):
        return Pdf()

    def save(self, pdf, path):
        if pdf.blocks is not None:
            compose(pdf.blocks, path)
        else:
            with open(path, "wb") as handle:
                pdf.writer.write(handle)

    def outline(self, pdf, limit=200):
        reader = pdf.reader
        if reader is None:
            return {"pages": 0, "composing": len(pdf.blocks)}
        meta = reader.metadata or {}
        fields = reader.get_fields() or {}
        return {
            "pages": len(reader.pages),
            "pageList": [{"address": f"page[{n}]", "size": [float(p.mediabox.width), float(p.mediabox.height)], "rotation": p.rotation}
                         for n, p in enumerate(reader.pages[:limit], 1)],
            "metadata": {k.lstrip("/"): str(v) for k, v in meta.items()},
            "outline": [{"title": t, "page": p} for t, p in outline_pages(reader)][:limit],
            "formFields": [{"address": f"form/field[{name}]", "type": str(f.get("/FT", "")), "value": str(f.get("/V", "")) if f.get("/V") is not None else None}
                           for name, f in list(fields.items())[:limit]],
            "attachments": sorted((reader.attachments or {}).keys()),
            "truncated": len(reader.pages) > limit,
        }

    def read(self, pdf, where, offset, limit):
        reader = pdf.reader
        if not where:
            pages = reader.pages[offset: offset + limit]
            return {"pages": [{"address": f"page[{n}]", "text": p.extract_text() or ""} for n, p in enumerate(pages, offset + 1)],
                    "next": offset + limit if offset + limit < len(reader.pages) else None}
        path = steps(where)
        if path[0][0] == "page" and len(path) == 1:
            page = pick(reader.pages, path[0][1], where, "page")
            return {"address": where, "kind": "page", "text": page.extract_text() or "", "rotation": page.rotation}
        if path[0][0] == "form" and len(path) == 2 and path[1][0] == "field":
            fields = reader.get_fields() or {}
            name = str(path[1][1])
            if name not in fields:
                raise Refusal("target_not_found", f"{where}: no form field '{name}'", fields=sorted(fields)[:50])
            f = fields[name]
            return {"address": where, "kind": "field", "type": str(f.get("/FT", "")), "value": str(f.get("/V")) if f.get("/V") is not None else None,
                    "options": [str(o) for o in f.get("/Opt", [])] or None}
        raise Refusal("target_not_found", f"{where} is not a PDF address", hint="page[N] or form/field[Name]")

    def find(self, pdf, needle, match_case, limit):
        fold = (lambda s: s) if match_case else str.lower
        found = []
        for n, page in enumerate(pdf.reader.pages, 1):
            content = page.extract_text() or ""
            if fold(needle) in fold(content):
                i = fold(content).find(fold(needle))
                found.append({"address": f"page[{n}]", "text": excerpt(content[max(0, i - 80): i + 120]), "count": fold(content).count(fold(needle))})
                if len(found) >= limit:
                    break
        return found

    def text(self, pdf):
        return "\n\n".join(f"# Page {n}\n{p.extract_text() or ''}" for n, p in enumerate(pdf.reader.pages, 1))


# ---- creating a PDF ------------------------------------------------------------------------

def compose(blocks, path):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    from xml.sax.saxutils import escape

    styles = getSampleStyleSheet()
    # Text beyond Latin-1 (Chinese, for example) needs a font that has it; reportlab ships CID fonts.
    wide = any(ord(ch) > 255 for block in blocks for ch in str(block))
    if wide:
        pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
        for style in styles.byName.values():
            style.fontName = "STSong-Light"
    story = []
    for block in blocks:
        kind = block["kind"]
        if kind == "heading":
            story.append(Paragraph(escape(block["text"]), styles[f"Heading{min(block['level'], 6)}"]))
        elif kind == "paragraph":
            story.append(Paragraph(escape(block["text"]).replace("\n", "<br/>"), styles["BodyText"]))
            story.append(Spacer(1, 4))
        elif kind == "table":
            table = Table([[Paragraph(escape(c), styles["BodyText"]) for c in row] for row in block["rows"]], repeatRows=1 if block.get("header") else 0)
            table.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.grey), ("VALIGN", (0, 0), (-1, -1), "TOP")] +
                                      ([("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke)] if block.get("header") else [])))
            story.append(table)
            story.append(Spacer(1, 8))
        elif kind == "break":
            story.append(PageBreak())
    SimpleDocTemplate(str(path), pagesize=A4).build(story or [Spacer(1, 1)])


@action("pdf", "add_heading", "compose", "Add a heading to a PDF being created.",
        {"text": text("The heading.", 2000), "level": {"type": "integer", "minimum": 1, "maximum": 6}}, ["text", "level"], creates=True)
def add_heading(pdf, args, edit):
    pdf.require_compose("add_heading")
    pdf.blocks.append({"kind": "heading", "text": args["text"], "level": args["level"]})
    return {"changed": [f"block[{len(pdf.blocks)}]"]}


@action("pdf", "add_paragraph", "compose", "Add a paragraph to a PDF being created.",
        {"text": text("The paragraph; newlines break lines.")}, ["text"], creates=True)
def add_paragraph(pdf, args, edit):
    pdf.require_compose("add_paragraph")
    pdf.blocks.append({"kind": "paragraph", "text": args["text"]})
    return {"changed": [f"block[{len(pdf.blocks)}]"]}


@action("pdf", "add_table", "compose", "Add a table of text to a PDF being created.",
        {"rows": {"type": "array", "minItems": 1, "maxItems": 2000, "items": {"type": "array", "minItems": 1, "maxItems": 30, "items": {"type": "string", "maxLength": 5000}}},
         "header": {"type": "boolean", "default": True, "description": "The first row is a header, repeated on each page."}},
        ["rows"], creates=True)
def add_table(pdf, args, edit):
    pdf.require_compose("add_table")
    if len({len(r) for r in args["rows"]}) != 1:
        raise Refusal("invalid_request", "Every row needs the same number of cells")
    pdf.blocks.append({"kind": "table", "rows": args["rows"], "header": args.get("header", True)})
    return {"changed": [f"block[{len(pdf.blocks)}]"]}


@action("pdf", "add_page_break", "compose", "Start a new page in a PDF being created.", {}, [], creates=True)
def add_page_break(pdf, args, edit):
    pdf.require_compose("add_page_break")
    pdf.blocks.append({"kind": "break"})
    return {"changed": [f"block[{len(pdf.blocks)}]"]}


# ---- changing a PDF ------------------------------------------------------------------------

@action("pdf", "select_pages", "pages", "Keep only the listed pages, in the order listed (reorders or removes pages).",
        {"pages": PAGES}, ["pages"])
def select_pages(pdf, args, edit):
    pdf.require_existing("select_pages")
    reorder(pdf, args["pages"], "select_pages", edit)
    return {"changed": ["pages"], "pages": len(args["pages"])}


@action("pdf", "delete_pages", "pages", "Delete pages.", {"pages": PAGES}, ["pages"])
def delete_pages(pdf, args, edit):
    pdf.require_existing("delete_pages")
    count = len(pdf.reader.pages)
    page_numbers(args["pages"], count, "delete_pages")
    keep = [n for n in range(1, count + 1) if n not in set(args["pages"])]
    if not keep:
        raise Refusal("invalid_request", "A PDF keeps at least one page")
    reorder(pdf, keep, "delete_pages", edit)
    return {"changed": [f"page[{n}]" for n in args["pages"]], "pages": len(keep)}


@action("pdf", "rotate_pages", "pages", "Rotate pages clockwise.",
        {"pages": PAGES, "degrees": {"type": "integer", "enum": [90, 180, 270]}}, ["pages", "degrees"])
def rotate_pages(pdf, args, edit):
    pdf.require_existing("rotate_pages")
    page_numbers(args["pages"], len(pdf.writer.pages), "rotate_pages")
    for n in args["pages"]:
        pdf.writer.pages[n - 1].rotate(args["degrees"])
    return {"changed": [f"page[{n}]" for n in args["pages"]]}


@action("pdf", "append_pdf", "pages", "Append the pages of other PDFs after the last page.",
        {"paths": {"type": "array", "minItems": 1, "maxItems": 50, "items": address("A PDF inside a folder OfficeMCP may read.")}},
        ["paths"], tier="recipe")
def append_pdf(pdf, args, edit):
    pdf.require_existing("append_pdf")
    added = 0
    for item in args["paths"]:
        source = edit.config.input(item)
        if source.suffix.lower() != ".pdf":
            raise Refusal("unsupported_format", f"{item} is not a PDF")
        other = PdfReader(source)
        if other.is_encrypted:
            raise Refusal("unreadable", f"{item} is encrypted")
        for page in other.pages:
            pdf.writer.add_page(page)
            added += 1
    return {"changed": ["pages"], "pagesAdded": added}


@action("pdf", "fill_form", "forms", "Fill form fields by name.",
        {"fields": {"type": "object", "minProperties": 1, "additionalProperties": {"type": "string", "maxLength": 10_000}},
         "expected": {"type": "object", "additionalProperties": {"type": ["string", "null"]}, "description": "Optional: current values by name."}},
        ["fields"])
def fill_form(pdf, args, edit):
    pdf.require_existing("fill_form")
    current = pdf.reader.get_fields() or {}
    unknown = [name for name in args["fields"] if name not in current]
    if unknown:
        raise Refusal("target_not_found", f"No form fields named {unknown}", fields=sorted(current)[:50])
    for name, value in (args.get("expected") or {}).items():
        actual = current.get(name, {}).get("/V")
        if (str(actual) if actual is not None else None) != value:
            raise Refusal("expected_mismatch", f"form/field[{name}] holds {actual!r}, not {value!r}")
    pages = field_pages(pdf.reader)
    for n in sorted({pages[name] for name in args["fields"] if name in pages}):
        pdf.writer.update_page_form_field_values(pdf.writer.pages[n - 1], {k: v for k, v in args["fields"].items() if pages.get(k) == n}, auto_regenerate=False)
    pdf.writer.set_need_appearances_writer(True)
    return {"changed": [f"form/field[{name}]" for name in args["fields"]]}


@action("pdf", "set_metadata", "document", "Set the PDF's title, author, subject or keywords.",
        {k: {"type": "string", "maxLength": 1000} for k in ("title", "author", "subject", "keywords")}, [])
def set_metadata(pdf, args, edit):
    values = {f"/{k.capitalize()}": v for k, v in args.items() if k in ("title", "author", "subject", "keywords")}
    if not values:
        raise Refusal("invalid_request", "Name at least one of title, author, subject, keywords")
    if pdf.writer is None:
        raise Refusal("invalid_request", "set_metadata changes an existing PDF")
    pdf.writer.add_metadata(values)
    return {"changed": [f"metadata/{k.strip('/').lower()}" for k in values]}
