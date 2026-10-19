"""Local, closed-operation bridge. No arbitrary expressions or remote resources."""
import json
import re
import sys
from importlib import import_module
from importlib.metadata import version


class OperationError(Exception):
    def __init__(self, code, message):
        self.code, self.message = code, message


def require(ok, code, message):
    if not ok:
        raise OperationError(code, message)


def license_engine(module, config, kind):
    lic = module.License()
    license_file = config.get("asposeCellsLicense" if kind == "cells" else "asposeSlidesLicense") or config.get("asposeLicense")
    if license_file:
        try:
            lic.set_license(license_file)
        except Exception:
            raise OperationError("ENGINE_LICENSE_INVALID", "Aspose rejected the configured local license file.")
    licensed = module.Workbook().is_licensed if kind == "cells" else lic.is_licensed()
    return bool(licensed)


def cell_value(cell):
    value = cell.value
    return value if value is None or isinstance(value, (str, bool, int, float)) else str(value)


_FORMULA_STRING = re.compile(r'"(?:[^"]|"")*"')
_EXTERNAL_BOOK = re.compile(
    r"(?:'[^']*\[[^\]]+\][^']*'|(?<![A-Za-z0-9_])\[[^\]]+\][^!+\-*/^&,;()]*?)!|"
    r"(?<![A-Za-z0-9_])\[[^\]]+\.(?:xlsx|xlsm|xlsb|xls)\]",
    re.IGNORECASE,
)


def formula_uses_external_resources(formula):
    """Reject external workbooks/network functions without rejecting table refs."""
    expression = _FORMULA_STRING.sub('""', formula)
    upper = expression.upper()
    return (
        bool(_EXTERNAL_BOOK.search(expression))
        or "|" in expression
        or any(name in upper for name in ("WEBSERVICE(", "RTD(", "HYPERLINK(", "DDE("))
    )


def shape_by_id(shapes, target):
    for shape in shapes:
        if str(shape.office_interop_shape_id) == target:
            return shape
        if hasattr(shape, "shapes"):
            found = shape_by_id(shape.shapes, target)
            if found is not None:
                return found
    return None


def run(task):
    kind, config, args = task["kind"], task.get("engines", {}), task.get("args", {})
    if kind == "office_engine_status":
        engines = {}
        for name, package in [("cells", "aspose-cells-python"), ("slides", "aspose-slides")]:
            try:
                module = import_module("aspose." + name)
                licensed = license_engine(module, config, name)
                engines[name] = {"installed": True, "version": version(package), "licensed": licensed, "evaluationAllowed": config.get("allowEvaluation", False)}
            except OperationError as exc:
                engines[name] = {"installed": True, "licensed": False, "error": exc.code}
            except Exception:
                engines[name] = {"installed": False, "error": "ENGINE_UNAVAILABLE"}
        return {"aspose": engines}
    family = "cells" if kind.startswith("xlsx_") else "slides"
    module = import_module("aspose." + family)
    licensed = license_engine(module, config, family)
    require(licensed or config.get("allowEvaluation"), "ENGINE_LICENSE_REQUIRED", "Configure engines.asposeLicense, or explicitly enable engines.allowEvaluation for labeled local trials.")
    if family == "slides" and kind in ("pptx_edit_shape_text", "pptx_set_notes"):
        require(licensed, "ENGINE_LICENSE_REQUIRED", "Aspose.Slides evaluation truncates text on load, so guarded edits require a valid local license. Evaluation inspection remains available.")
    result = {"engine": "Aspose." + family.title(), "evaluation": not licensed}
    if family == "cells":
        book = module.Workbook(task["input"])
        names = [s.name for s in book.worksheets]
        targets = []
        if kind == "xlsx_engine_read":
            require(args["sheet"] in names, "TARGET_CONFLICT", "Worksheet name is unavailable.")
            sheet = next(s for s in book.worksheets if s.name == args["sheet"])
            start, end = (sheet.cells.get(x) for x in args["range"].split(":"))
            require(start.row <= end.row < 1048576 and start.column <= end.column < 16384 and (end.row - start.row + 1) * (end.column - start.column + 1) <= 1000, "LIMIT_EXCEEDED", "Range must be within XLSX bounds and contain at most 1000 cells.")
            cells = []
            for row in range(start.row, end.row + 1):
                for col in range(start.column, end.column + 1):
                    cell = sheet.cells.get(row, col); style = cell.get_style()
                    cells.append({"address": cell.name, "value": cell_value(cell), "formula": cell.formula or None, "displayText": cell.display_string_value, "format": {"bold": style.font.is_bold, "italic": style.font.is_italic, "fontSize": style.font.size, "numberFormat": style.custom, "wrapText": style.is_text_wrapped}})
            return {**result, "cells": cells, "sheet": args["sheet"], "range": args["range"]}
        if kind == "xlsx_edit_range":
            require(args["sheet"] in names, "TARGET_CONFLICT", "Worksheet name is unavailable.")
            sheet = next(s for s in book.worksheets if s.name == args["sheet"])
            seen = set()
            for edit in args["cells"]:
                address = edit["address"]
                require(address not in seen, "INVALID_ARGUMENT", "Duplicate cell targets are not allowed.")
                seen.add(address)
                cell = sheet.cells.get(address)
                require(cell.row < 1048576 and cell.column < 16384, "INVALID_ARGUMENT", "Cell address exceeds XLSX bounds.")
                require(cell_value(cell) == edit["expectedValue"] and (cell.formula or None) == edit["expectedFormula"], "TARGET_CONFLICT", "Cell value or formula differs from the expected value.")
            for edit in args["cells"]:
                cell = sheet.cells.get(edit["address"])
                cell.formula = None
                value = edit["value"]
                if isinstance(value, str):
                    cell.put_value(value, False)
                else:
                    cell.put_value(value)
                fmt = edit.get("format", {})
                if fmt:
                    style = cell.get_style()
                    if "numberFormat" in fmt: style.custom = fmt["numberFormat"]
                    if "bold" in fmt: style.font.is_bold = fmt["bold"]
                    if "italic" in fmt: style.font.is_italic = fmt["italic"]
                    if "fontSize" in fmt: style.font.size = int(fmt["fontSize"])
                    if "wrapText" in fmt: style.is_text_wrapped = fmt["wrapText"]
                    cell.set_style(style)
                targets.append(args["sheet"] + "!" + edit["address"])
        elif kind == "xlsx_formula":
            sheet = next((s for s in book.worksheets if s.name == args["sheet"]), None)
            require(sheet is not None, "TARGET_CONFLICT", "Worksheet does not exist.")
            cell = sheet.cells.get(args["address"])
            require((cell.formula or None) == args["expectedFormula"], "TARGET_CONFLICT", "Formula changed; read the current cell again.")
            formula = args["formula"] if args["formula"].startswith("=") else "=" + args["formula"]
            require(not formula_uses_external_resources(formula), "UNSUPPORTED_FEATURE", "External workbook/network formulas are outside this local calculation operation.")
            cell.formula = formula
            targets.append(args["sheet"] + "!" + args["address"])
        elif kind == "xlsx_worksheet":
            require(args["action"] == "add" or args["sheet"] in names, "TARGET_CONFLICT", "Worksheet does not exist.")
            if args["action"] == "add":
                require(args["sheet"] not in names, "TARGET_CONFLICT", "Worksheet name already exists.")
                book.worksheets.add(args["sheet"]); names.append(args["sheet"])
            elif args["action"] == "rename":
                require(args.get("name") and args["name"] not in names, "INVALID_ARGUMENT", "Provide a new unique worksheet name.")
                next(s for s in book.worksheets if s.name == args["sheet"]).name = args["name"]
                names[names.index(args["sheet"])] = args["name"]
            else:
                require(len(names) > 1, "UNSUPPORTED_FEATURE", "Workbook must retain a worksheet.")
                book.worksheets.remove_at(names.index(args["sheet"])); names.remove(args["sheet"])
            targets.append(args["sheet"])
        elif kind == "xlsx_rows_columns":
            sheet = next((s for s in book.worksheets if s.name == args["sheet"]), None)
            require(sheet is not None, "TARGET_CONFLICT", "Worksheet does not exist.")
            index, count = args["index"] - 1, args["count"]
            require(index + count <= (1048576 if args["axis"] == "rows" else 16384), "INVALID_ARGUMENT", "Row/column operation exceeds worksheet bounds.")
            method = ("insert_" if args["action"] == "insert" else "delete_") + args["axis"]
            getattr(sheet.cells, method)(index, count, True)
            targets.append(args["sheet"])
        elif kind != "xlsx_recalculate":
            raise OperationError("INVALID_ARGUMENT", "Unknown workbook operation.")
        # Never leave stale formula caches after a value edit. External links are
        # not refreshed; Aspose calculates against the values already in the file.
        book.calculate_formula()
        errors = []
        error_count = 0
        for sheet in book.worksheets:
            for cell in sheet.cells:
                if cell.type == module.CellValueType.IS_ERROR:
                    error_count += 1
                    if len(errors) < 100: errors.append({"sheet": sheet.name, "address": cell.name, "value": str(cell.value)})
        book.save(task["output"], module.SaveFormat.XLSX)
        reopened = module.Workbook(task["output"])
        # In evaluation mode Aspose may add an evaluation sheet; existing sheets
        # must remain, including order relative to one another.
        after_names = [s.name for s in reopened.worksheets]
        require([n for n in after_names if n in names] == names, "PRESERVATION_FAILED", "Worksheet inventory changed.")
        if kind == "xlsx_edit_range":
            for edit in args["cells"]:
                cell = next(s for s in reopened.worksheets if s.name == args["sheet"]).cells.get(edit["address"])
                require(cell_value(cell) == edit["value"] and not cell.formula, "VALIDATION_FAILED", "Saved cell differs from the requested literal value.")
                style = cell.get_style()
                actual = {"numberFormat": style.custom, "bold": style.font.is_bold, "italic": style.font.is_italic, "fontSize": style.font.size, "wrapText": style.is_text_wrapped}
                require(all(actual[k] == v for k, v in edit.get("format", {}).items()), "VALIDATION_FAILED", "Saved cell formatting differs from the requested properties.")
        if kind == "xlsx_formula":
            actual = next(s for s in reopened.worksheets if s.name == args["sheet"]).cells.get(args["address"]).formula
            require(actual == formula, "VALIDATION_FAILED", "Saved formula differs from the requested formula.")
        result.update(affectedTargets=targets or ["workbook-formula-caches"], formulaErrors=errors, formulaErrorCount=error_count, calculation="Aspose.Cells; external data not refreshed")
        return result
    with module.Presentation(task["input"]) as presentation:
        require(args["slide"] <= len(presentation.slides), "TARGET_CONFLICT", "Slide is unavailable.")
        slide = presentation.slides[args["slide"] - 1]
        if kind == "pptx_engine_read":
            shapes = []
            def visit(collection):
                for shape in collection:
                    shapes.append({"shapeId": str(shape.office_interop_shape_id), "name": shape.name, "text": shape.text_frame.text if hasattr(shape, "text_frame") and shape.text_frame is not None else None, "x": shape.x, "y": shape.y, "width": shape.width, "height": shape.height})
                    if hasattr(shape, "shapes"): visit(shape.shapes)
            visit(slide.shapes)
            require(len(shapes) <= 1000, "LIMIT_EXCEEDED", "Slide contains too many shapes for one response.")
            notes = slide.notes_slide_manager.notes_slide
            return {**result, "shapes": shapes, "notes": notes.notes_text_frame.text if notes is not None and notes.notes_text_frame is not None else "", "slide": args["slide"], "slideCount": len(presentation.slides)}
        if kind == "pptx_edit_shape_text":
            shape = shape_by_id(slide.shapes, args["shapeId"])
            require(shape is not None and hasattr(shape, "text_frame") and shape.text_frame is not None, "TARGET_CONFLICT", "Text shape is unavailable.")
            require(shape.text_frame.text == args["expectedText"], "TARGET_CONFLICT", "Shape text differs; paragraph breaks use newline characters.")
            shape.text_frame.text = args["replacement"]
        elif kind == "pptx_set_notes":
            notes = slide.notes_slide_manager.notes_slide
            current = notes.notes_text_frame.text if notes is not None and notes.notes_text_frame is not None else ""
            require(current == args["expectedText"], "TARGET_CONFLICT", "Speaker notes differ from expected text.")
            if notes is None: notes = slide.notes_slide_manager.add_notes_slide()
            notes.notes_text_frame.text = args["text"]
        else:
            raise OperationError("INVALID_ARGUMENT", "Unknown presentation operation.")
        count = len(presentation.slides)
        presentation.save(task["output"], module.export.SaveFormat.PPTX)
    with module.Presentation(task["output"]) as reopened:
        require(len(reopened.slides) == count, "PRESERVATION_FAILED", "Slide count changed.")
        slide = reopened.slides[args["slide"] - 1]
        actual = shape_by_id(slide.shapes, args["shapeId"]).text_frame.text if kind == "pptx_edit_shape_text" else slide.notes_slide_manager.notes_slide.notes_text_frame.text
        expected = args["replacement"] if kind == "pptx_edit_shape_text" else args["text"]
        require(actual == expected, "VALIDATION_FAILED", "Saved slide text differs from the requested text; evaluation restrictions may apply.")
    result["affectedTargets"] = [args["slide"]]
    return result


if __name__ == "__main__":
    task = json.load(sys.stdin)
    try:
        response = {"ok": True, "result": run(task)}
    except OperationError as exc:
        response = {"ok": False, "error": {"code": exc.code, "message": exc.message}}
    except ImportError:
        response = {"ok": False, "error": {"code": "ENGINE_UNAVAILABLE", "message": "Install the pinned Aspose Python packages in the configured interpreter."}}
    except Exception:
        response = {"ok": False, "error": {"code": "ENGINE_FAILED", "message": "Aspose could not complete this operation. No candidate was accepted."}}
    # A file keeps native runtime diagnostics out of the JSON protocol.
    with open(task["resultFile"], "x", encoding="utf-8") as output:
        json.dump(response, output, allow_nan=False)
