import unittest

from pypdf import PdfReader

from .support import Case, make_pdf


class PdfTests(Case):
    def setUp(self):
        super().setUp()
        self.path = make_pdf(self.inputs)

    def test_inspect_read_and_find(self):
        outline = self.ok(self.service.inspect(str(self.path)))["outline"]
        self.assertEqual(outline["pages"], 3)
        self.assertEqual(outline["outline"], [{"title": "Start", "page": 1}])
        self.assertEqual(outline["formFields"][0]["address"], "form/field[Customer]")
        self.assertIn("Page 2", self.ok(self.service.read(str(self.path), "page[2]"))["text"])
        self.assertEqual([m["address"] for m in self.ok(self.service.find(str(self.path), "page 3"))["matches"]], ["page[3]"])

    def test_pages_rotate_reorder_and_delete(self):
        result = self.ok(self.edit(self.path, [{"action": "rotate_pages", "pages": [2], "degrees": 90},
                                               {"action": "select_pages", "pages": [1, 3, 2]}]))
        reader = PdfReader(result["path"])
        self.assertEqual([("Page 3" in p.extract_text(), p.rotation) for p in reader.pages], [(False, 0), (True, 0), (False, 90)])
        deleted = self.ok(self.edit(self.path, [{"action": "delete_pages", "pages": [3]}]))
        self.assertEqual(len(PdfReader(deleted["path"]).pages), 2)

    def test_removing_the_page_an_outline_or_field_points_to_is_refused(self):
        self.refused(self.edit(self.path, [{"action": "delete_pages", "pages": [1]}]), "unsupported_content")

    def test_forms_metadata_and_appending(self):
        other = make_pdf(self.inputs, "annex.pdf", pages=2, outline=False, form=False)
        result = self.ok(self.edit(self.path, [{"action": "fill_form", "fields": {"Customer": "Acme 中文"}, "expected": {"Customer": ""}},
                                               {"action": "set_metadata", "title": "Contract"},
                                               {"action": "append_pdf", "paths": [str(other)]}]))
        reader = PdfReader(result["path"])
        self.assertEqual(reader.get_fields()["Customer"]["/V"], "Acme 中文")
        self.assertEqual(reader.metadata.title, "Contract")
        self.assertEqual(len(reader.pages), 5)
        self.refused(self.edit(self.path, [{"action": "fill_form", "fields": {"Nope": "x"}}]), "target_not_found")

    def test_create_a_pdf_with_compose_actions_including_chinese(self):
        result = self.ok(self.service.create("pdf", str(self.inputs / "memo.pdf"), [
            {"action": "add_heading", "text": "Memo 备忘录", "level": 1},
            {"action": "add_paragraph", "text": "First line\nSecond line"},
            {"action": "add_table", "rows": [["Item", "Qty"], ["Paper", "10"]]},
            {"action": "add_page_break"},
            {"action": "add_paragraph", "text": "Page two"},
        ]))
        reader = PdfReader(result["path"])
        self.assertEqual(len(reader.pages), 2)
        self.assertIn("Paper", reader.pages[0].extract_text())
        self.refused(self.edit(self.path, [{"action": "add_paragraph", "text": "x"}]), "invalid_request")


if __name__ == "__main__":
    unittest.main()
