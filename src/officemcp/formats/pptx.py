"""PowerPoint presentations through python-pptx.

Addresses: `slide[N]`, `slide[N]/shape[ID]` (the shape's ID, as office_inspect lists it),
`slide[N]/shape[ID]/r[R]/c[C]` (a table cell) and `slide[N]/notes`.
"""

import hashlib
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches

from ..catalog import EXPECTED_TEXT, action, address, text
from ..errors import Refusal
from .common import clone, excerpt, expect, pick, replace_in_runs, steps

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def slides(presentation):
    return list(presentation.slides)


def shape_by_id(slide, shape_id, where):
    for shape in slide.shapes:
        if str(shape.shape_id) == str(shape_id):
            return shape
    raise Refusal("target_not_found", f"{where}: no shape with ID {shape_id} on this slide", hint="office_inspect lists shape IDs")


def resolve(presentation, where: str):
    path = steps(where)
    if not path or path[0][0] != "slide":
        raise Refusal("target_not_found", f"{where}: PPTX addresses start with slide[N]")
    slide = pick(slides(presentation), path[0][1], where, "slide")
    if len(path) == 1:
        return slide
    if path[1][0] == "notes" and len(path) == 2:
        return ("notes", slide)
    if path[1][0] != "shape":
        raise Refusal("target_not_found", f"{where}: expected shape[ID] or notes after slide[N]")
    shape = shape_by_id(slide, path[1][1], where)
    if len(path) == 2:
        return shape
    if len(path) == 4 and path[2][0] == "r" and path[3][0] == "c" and getattr(shape, "has_table", False):
        rows = list(shape.table.rows)
        row = pick(rows, path[2][1], where, "row")
        return pick(list(row.cells), path[3][1], where, "cell")
    raise Refusal("target_not_found", f"{where} is not a PPTX address")


def title_of(slide) -> str:
    title = slide.shapes.title
    return title.text_frame.text if title is not None and title.has_text_frame else ""


def kind(shape) -> str:
    if shape.is_placeholder:
        return f"placeholder:{shape.placeholder_format.type}".lower().replace(" ", "_").replace("(", "").replace(")", "")
    if getattr(shape, "has_table", False):
        return "table"
    if getattr(shape, "has_chart", False):
        return "chart"
    return str(shape.shape_type).lower() if shape.shape_type is not None else "shape"


def frame_text(frame) -> str:
    return frame.text


def set_frame(frame, value: str, where: str) -> None:
    """Replace a text frame's text, keeping the first paragraph's and first run's formatting."""
    for element in frame._txBody.iter():
        tag = element.tag
        if tag in (f"{A}fld",):
            raise Refusal("unsupported_content", f"{where} holds a field (slide number, date) that a text change would lose",
                          hint="Use replace_text instead")
    paragraphs = frame.paragraphs
    first = paragraphs[0]
    template = clone(first._p)
    for p in paragraphs[1:]:
        p._p.getparent().remove(p._p)
    lines = value.split("\n")
    fill(first, lines[0])
    anchor = first._p
    for line in lines[1:]:
        new = clone(template)
        anchor.addnext(new)
        anchor = new
        from pptx.text.text import _Paragraph
        fill(_Paragraph(new, first._parent), line)


def fill(paragraph, value: str) -> None:
    runs = paragraph.runs
    for br in paragraph._p.findall(f"{A}br"):
        paragraph._p.remove(br)
    if not runs:
        paragraph.add_run().text = value
        return
    runs[0].text = value
    for run in runs[1:]:
        run._r.getparent().remove(run._r)


def text_shapes(slide):
    for shape in slide.shapes:
        if shape.has_text_frame:
            yield shape


class Handler:
    format = "pptx"
    extension = ".pptx"

    def open(self, path):
        try:
            return Presentation(str(path))
        except Exception as error:
            raise Refusal("unreadable", f"Not a presentation python-pptx can open: {error}") from None

    def blank(self):
        return Presentation()

    def save(self, presentation, path):
        presentation.save(str(path))

    def outline(self, presentation, limit=200):
        out = []
        for n, slide in enumerate(slides(presentation)[:limit], 1):
            shapes = []
            for shape in slide.shapes:
                item = {"address": f"slide[{n}]/shape[{shape.shape_id}]", "name": shape.name, "kind": kind(shape)}
                if shape.has_text_frame:
                    item["text"] = excerpt(shape.text_frame.text)
                if getattr(shape, "has_table", False):
                    item["rows"], item["columns"] = len(shape.table.rows), len(shape.table.columns)
                shapes.append(item)
            entry = {"address": f"slide[{n}]", "layout": slide.slide_layout.name, "title": title_of(slide), "shapes": shapes}
            if slide.has_notes_slide and slide.notes_slide.notes_text_frame is not None:
                entry["notes"] = excerpt(slide.notes_slide.notes_text_frame.text)
            out.append(entry)
        return {"slides": len(presentation.slides), "slideSize": [presentation.slide_width, presentation.slide_height],
                "layouts": [layout.name for layout in presentation.slide_layouts], "outline": out,
                "truncated": len(presentation.slides) > limit}

    def read(self, presentation, where, offset, limit):
        if not where:
            outline = self.outline(presentation, limit=100_000)
            return {"slides": outline["outline"][offset: offset + limit], "next": offset + limit if offset + limit < outline["slides"] else None}
        target = resolve(presentation, where)
        if isinstance(target, tuple):
            slide = target[1]
            return {"address": where, "kind": "notes", "text": slide.notes_slide.notes_text_frame.text if slide.has_notes_slide else ""}
        if hasattr(target, "shapes"):
            n = steps(where)[0][1]
            return {"address": where, "kind": "slide", "layout": target.slide_layout.name,
                    "shapes": [{"address": f"slide[{n}]/shape[{s.shape_id}]", "name": s.name, "kind": kind(s),
                                **({"paragraphs": [p.text for p in s.text_frame.paragraphs]} if s.has_text_frame else {})} for s in target.shapes]}
        if getattr(target, "has_table", False):
            return {"address": where, "kind": "table",
                    "rows": [[{"address": f"{where}/r[{i}]/c[{j}]", "text": c.text} for j, c in enumerate(row.cells, 1)] for i, row in enumerate(target.table.rows, 1)]}
        if hasattr(target, "text_frame") and getattr(target, "has_text_frame", True):
            return {"address": where, "kind": "shape" if hasattr(target, "shape_id") else "cell", "text": target.text_frame.text,
                    "paragraphs": [{"text": p.text, "level": p.level} for p in target.text_frame.paragraphs]}
        return {"address": where, "kind": kind(target), "text": None}

    def find(self, presentation, needle, match_case, limit):
        fold = (lambda s: s) if match_case else str.lower
        found = []
        for n, slide in enumerate(slides(presentation), 1):
            for shape in slide.shapes:
                if shape.has_text_frame and fold(needle) in fold(shape.text_frame.text):
                    found.append({"address": f"slide[{n}]/shape[{shape.shape_id}]", "text": excerpt(shape.text_frame.text)})
                if getattr(shape, "has_table", False):
                    for i, row in enumerate(shape.table.rows, 1):
                        for j, cell in enumerate(row.cells, 1):
                            if fold(needle) in fold(cell.text):
                                found.append({"address": f"slide[{n}]/shape[{shape.shape_id}]/r[{i}]/c[{j}]", "text": excerpt(cell.text)})
            if slide.has_notes_slide and fold(needle) in fold(slide.notes_slide.notes_text_frame.text or ""):
                found.append({"address": f"slide[{n}]/notes", "text": excerpt(slide.notes_slide.notes_text_frame.text)})
            if len(found) >= limit:
                break
        return found[:limit]

    def text(self, presentation):
        lines = []
        for n, slide in enumerate(slides(presentation), 1):
            lines.append(f"# Slide {n}")
            lines.extend(s.text_frame.text for s in text_shapes(slide))
        return "\n".join(lines)


# ---- actions -------------------------------------------------------------------------------

SHAPE = address("A shape with text: slide[N]/shape[ID].")
SLIDE = address("A slide: slide[N].")


@action("pptx", "set_shape_text", "text", "Replace a shape's text (lines become paragraphs), keeping its first paragraph's formatting.",
        {"target": address("A shape with text, slide[N]/shape[ID], or a table cell, slide[N]/shape[ID]/r[R]/c[C]."),
         "expectedText": EXPECTED_TEXT, "text": text("The new text; newlines start new paragraphs.")},
        ["target", "expectedText", "text"])
def set_shape_text(presentation, args, edit):
    target = resolve(presentation, args["target"])
    frame = getattr(target, "text_frame", None)
    if frame is None or (hasattr(target, "has_text_frame") and not target.has_text_frame):
        raise Refusal("unsupported_content", f"{args['target']} holds no text")
    expect(frame.text, args["expectedText"], args["target"])
    set_frame(frame, args["text"], args["target"])
    return {"changed": [args["target"]]}


@action("pptx", "replace_text", "text", "Replace text wherever it occurs (on one slide or shape, or the whole deck), keeping each run's formatting.",
        {"find": text("The text to find.", 10_000), "replace": text("The replacement."),
         "target": address("Optional: limit to slide[N] or slide[N]/shape[ID]."),
         "expectedCount": {"type": "integer", "minimum": 1}, "matchCase": {"type": "boolean", "default": True}},
        ["find", "replace"], tier="recipe")
def replace_text(presentation, args, edit):
    frames = []
    scope = resolve(presentation, args["target"]) if args.get("target") else None
    for n, slide in enumerate(slides(presentation), 1):
        if scope is not None and hasattr(scope, "shapes") and scope is not slide:
            continue
        for shape in slide.shapes:
            if scope is not None and not hasattr(scope, "shapes") and shape is not scope and getattr(shape, "shape_id", None) != getattr(scope, "shape_id", -1):
                continue
            if shape.has_text_frame:
                frames.append((f"slide[{n}]/shape[{shape.shape_id}]", shape.text_frame))
            if getattr(shape, "has_table", False):
                for i, row in enumerate(shape.table.rows, 1):
                    for j, cell in enumerate(row.cells, 1):
                        frames.append((f"slide[{n}]/shape[{shape.shape_id}]/r[{i}]/c[{j}]", cell.text_frame))
    match_case = args.get("matchCase", True)
    fold = (lambda s: s) if match_case else str.lower
    total = sum(fold(f.text).count(fold(args["find"])) for _, f in frames)
    if not total:
        raise Refusal("no_match", f"'{args['find']}' does not occur")
    if args.get("expectedCount") and total != args["expectedCount"]:
        raise Refusal("ambiguous_match", f"'{args['find']}' occurs {total} times, not {args['expectedCount']}", occurrences=total)
    changed, done = [], 0
    for where, frame in frames:
        count = sum(replace_in_runs(p.runs, args["find"], args["replace"], match_case) for p in frame.paragraphs)
        if count:
            changed.append(where)
            done += count
    if done != total:
        raise Refusal("unsupported_content", f"Some occurrences span line breaks or fields this action does not change ({done} of {total})")
    return {"changed": changed, "replaced": total}


@action("pptx", "add_slide", "slides", "Add a slide from one of the presentation's layouts, with an optional title and bullet points.",
        {"layout": {"type": ["string", "integer"], "description": "A layout name (office_inspect lists them) or its position, from 1."},
         "title": text("Optional title.", 2000),
         "bullets": {"type": "array", "maxItems": 100, "items": {"type": "string", "maxLength": 5000}, "description": "Optional body lines."},
         "position": {"type": "integer", "minimum": 1, "description": "Optional: where the slide goes, from 1; last otherwise."}},
        ["layout"], tier="recipe")
def add_slide(presentation, args, edit):
    layouts = list(presentation.slide_layouts)
    if isinstance(args["layout"], int):
        layout = pick(layouts, args["layout"], "layout", "layout")
    else:
        matches = [l for l in layouts if l.name == args["layout"]]
        if not matches:
            raise Refusal("target_not_found", f"No layout named '{args['layout']}'", layouts=[l.name for l in layouts])
        layout = matches[0]
    slide = presentation.slides.add_slide(layout)
    if args.get("title") is not None:
        if slide.shapes.title is None:
            raise Refusal("unsupported_content", f"Layout '{layout.name}' has no title placeholder")
        slide.shapes.title.text_frame.text = args["title"]
    if args.get("bullets"):
        body = next((s for s in slide.placeholders if s.placeholder_format.idx != 0 and s.has_text_frame), None)
        if body is None:
            raise Refusal("unsupported_content", f"Layout '{layout.name}' has no body placeholder for bullets")
        body.text_frame.text = args["bullets"][0]
        for line in args["bullets"][1:]:
            body.text_frame.add_paragraph().text = line
    count = len(presentation.slides)
    position = args.get("position", count)
    if not 1 <= position <= count:
        raise Refusal("invalid_request", f"position must be between 1 and {count}")
    ids = presentation.slides._sldIdLst
    element = ids[-1]
    ids.remove(element)
    ids.insert(position - 1, element)
    return {"changed": [f"slide[{position}]"], "inserted": "slide"}


@action("pptx", "delete_slide", "slides", "Delete a slide.",
        {"target": SLIDE, "expectedTitle": text("Optional: the slide's current title.", 2000)}, ["target"])
def delete_slide(presentation, args, edit):
    slide = resolve(presentation, args["target"])
    if not hasattr(slide, "shapes"):
        raise Refusal("target_not_found", f"{args['target']} is not a slide")
    if args.get("expectedTitle") is not None:
        expect(title_of(slide), args["expectedTitle"], args["target"])
    if len(presentation.slides) == 1:
        raise Refusal("unsupported_content", "A presentation keeps at least one slide")
    # What disappears with the slide: its part, its notes, and media no other slide uses.
    edit.allow("part:application/vnd.openxmlformats-officedocument.presentationml.slide+xml")
    edit.allow("feature:PowerPoint slides")
    if slide.has_notes_slide:
        edit.allow("part:application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml")
    for rel in slide.part.rels.values():
        if not rel.is_external:
            target = rel.target_part
            edit.allow(f"part:{target.content_type}")
            if not target.partname.endswith(".xml"):
                edit.allow(f"binary:{hashlib.sha256(target.blob).hexdigest()}")
    for element in slide._element.iter():
        from ..fidelity import FEATURES
        label = FEATURES.get(element.tag)
        if label:
            edit.allow(f"feature:{label}")
    ids = presentation.slides._sldIdLst
    for entry in list(ids):
        if presentation.part.related_part(entry.rId) is slide.part:
            presentation.part.drop_rel(entry.rId)
            ids.remove(entry)
            break
    return {"changed": [args["target"]], "deleted": "slide"}


@action("pptx", "move_slide", "slides", "Move a slide to another position.",
        {"target": SLIDE, "position": {"type": "integer", "minimum": 1}}, ["target", "position"])
def move_slide(presentation, args, edit):
    path = steps(args["target"])
    ids = presentation.slides._sldIdLst
    entry = pick(list(ids), path[0][1], args["target"], "slide")
    if not 1 <= args["position"] <= len(ids):
        raise Refusal("invalid_request", f"position must be between 1 and {len(ids)}")
    ids.remove(entry)
    ids.insert(args["position"] - 1, entry)
    return {"changed": [f"slide[{args['position']}]"]}


@action("pptx", "set_notes", "slides", "Replace a slide's speaker notes.",
        {"target": SLIDE, "text": text("The notes."), "expectedText": text("Optional: the current notes.")}, ["target", "text"])
def set_notes(presentation, args, edit):
    slide = resolve(presentation, args["target"])
    if not hasattr(slide, "shapes"):
        raise Refusal("target_not_found", f"{args['target']} is not a slide")
    frame = slide.notes_slide.notes_text_frame
    expect(frame.text, args.get("expectedText"), args["target"] + "/notes")
    frame.text = args["text"]
    return {"changed": [args["target"] + "/notes"]}


@action("pptx", "add_picture", "objects", "Place a picture (PNG, JPEG, GIF) on a slide.",
        {"target": SLIDE, "image": address("The picture's path, inside a folder OfficeMCP may read."),
         "left": {"type": "number", "minimum": 0, "description": "Inches from the left edge."},
         "top": {"type": "number", "minimum": 0, "description": "Inches from the top edge."},
         "width": {"type": "number", "exclusiveMinimum": 0, "description": "Optional width in inches; keeps the aspect ratio."}},
        ["target", "image", "left", "top"])
def add_picture(presentation, args, edit):
    slide = resolve(presentation, args["target"])
    if not hasattr(slide, "shapes"):
        raise Refusal("target_not_found", f"{args['target']} is not a slide")
    image = edit.config.input(args["image"])
    if Path(image).suffix.lower() not in (".png", ".jpg", ".jpeg", ".gif"):
        raise Refusal("unsupported_format", "Pictures are PNG, JPEG or GIF")
    shape = slide.shapes.add_picture(str(image), Inches(args["left"]), Inches(args["top"]),
                                     width=Inches(args["width"]) if args.get("width") else None)
    return {"changed": [f"{args['target']}/shape[{shape.shape_id}]"], "inserted": "picture"}
