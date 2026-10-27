"""Rules every tool follows: paths, revisions, outputs, dry runs, unknown actions, the fidelity check."""

import io
from pathlib import Path
import os
import unittest
import zipfile

from officemcp import catalog
from officemcp.fidelity import losses, ooxml_inventory

from .support import Case, make_docx, make_xlsx


class ServiceTests(Case):
    def test_paths_outside_the_roots_are_refused(self):
        outside = self.root / "elsewhere.docx"
        make_docx(self.root, "elsewhere.docx")
        self.refused(self.service.inspect(str(outside)), "path_not_allowed")
        link = self.inputs / "link.docx"
        os.symlink(outside, link)
        self.refused(self.service.inspect(str(link)), "path_not_allowed")
        self.refused(self.service.inspect(str(self.inputs / "missing.docx")), "not_found")
        (self.inputs / "notes.txt").write_text("x")
        self.refused(self.service.inspect(str(self.inputs / "notes.txt")), "unsupported_format")

    def test_edits_save_in_place_after_a_revision_check(self):
        path = make_docx(self.inputs)
        os.chmod(path, 0o640)
        original = path.read_bytes()
        action = {"action": "set_properties", "title": "T"}
        self.refused(self.service.edit(str(path), "sha256:" + "0" * 64, [action]), "revision_mismatch")
        dry = self.ok(self.edit(path, [action], dryRun=True))
        self.assertTrue(dry["dryRun"])
        self.assertEqual(path.read_bytes(), original, "a dry run writes nothing")
        saved = self.ok(self.edit(path, [action]))
        self.assertEqual((saved["path"], saved["sourceRevision"] != saved["revision"]), (str(path), True))
        self.assertEqual(path.stat().st_mode & 0o777, 0o640, "the file keeps its permissions")
        relative = self.ok(self.service.inspect("report.docx"))
        self.assertEqual(relative["revision"], saved["revision"], "relative paths start at the first root")
        self.assertEqual(self.files(), ["report.docx"], "nothing is left beside the file")

    def test_new_files_never_overwrite(self):
        path = make_docx(self.inputs)
        action = {"action": "set_properties", "title": "T"}
        copy = self.ok(self.edit(path, [action], output={"path": "final.docx"}))
        self.assertEqual(copy["path"], str(self.inputs / "final.docx"))
        self.refused(self.edit(path, [action], output={"path": "final.docx"}), "output_exists")
        self.refused(self.edit(path, [action], output={"path": "final.pdf"}), "invalid_request")
        self.refused(self.edit(path, [action], output={"path": str(self.root / "out.docx")}), "path_not_allowed")
        self.refused(self.edit(path, [action], output={"path": "missing/out.docx"}), "not_found")
        self.refused(self.service.create("docx", "report.docx"), "output_exists")
        first = self.ok(self.service.convert(str(path), "txt"))
        second = self.ok(self.service.convert(str(path), "txt"))
        self.assertEqual([Path(first["path"]).name, Path(second["path"]).name], ["report.txt", "report-2.txt"])
        self.assertEqual(self.files(), ["final.docx", "report-2.txt", "report.docx", "report.txt"])

    def test_unknown_actions_and_bad_fields_are_refused(self):
        path = make_xlsx(self.inputs)
        self.refused(self.edit(path, [{"action": "explode"}]), "unknown_action")
        result = self.refused(self.edit(path, [{"action": "set_cells", "target": "Budget!A1", "values": "x"}]), "invalid_request")
        self.assertIn("values", result["message"])
        self.refused(self.edit(path, [{"action": "set_cells", "target": "Budget!A1", "values": [[1]], "extra": 1}]), "invalid_request")

    def test_bookkeeping_extensions_are_not_counted_as_losses(self):
        path = make_xlsx(self.inputs)
        source = zipfile.ZipFile(path)
        rewritten = io.BytesIO()
        with zipfile.ZipFile(rewritten, "w") as out:
            for item in source.infolist():
                data = source.read(item.filename)
                if item.filename == "xl/workbook.xml":
                    data = data.replace(b"<workbookPr", b'<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
                                        b'<mc:Choice Requires="x15"><x15ac:absPath xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac" url="/tmp/"/>'
                                        b'</mc:Choice></mc:AlternateContent><workbookPr', 1)
                if item.filename == "xl/styles.xml":
                    data = data.replace(b"</styleSheet>", b'<extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
                                        b'<x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext></extLst></styleSheet>')
                out.writestr(item, data)
        marked = self.inputs / "marked.xlsx"
        marked.write_bytes(rewritten.getvalue())
        self.assertEqual(self.ok(self.service.inspect(str(marked)))["cannotKeep"], [])

    def test_losses_count_only_what_no_action_declared(self):
        before = ooxml_inventory(make_docx(self.inputs))
        after = before.copy()
        after["feature:Word drawings and images"] -= 1
        self.assertEqual(losses(before, after, catalog.Edit(self.config).allowed), [{"what": "Word drawings and images", "count": 1}])
        allowed = catalog.Edit(self.config)
        allowed.allow("feature:Word drawings and images")
        self.assertEqual(losses(before, after, allowed.allowed), [])

    def test_text_conversion_of_every_format_reader(self):
        path = make_docx(self.inputs)
        result = self.ok(self.service.convert(str(path), "txt"))
        self.assertIn("Quarterly report", Path(result["path"]).read_text(encoding="utf-8"))
        self.refused(self.service.convert(str(path), "csv"), "unsupported_format")


if __name__ == "__main__":
    unittest.main()
