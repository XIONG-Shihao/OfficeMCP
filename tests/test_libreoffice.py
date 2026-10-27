"""Conversion to PDF and formula recalculation through LibreOffice, when it is installed."""

import unittest

import openpyxl
from pypdf import PdfReader

from .support import Case, make_docx, make_xlsx


class LibreOfficeTests(Case):
    def setUp(self):
        super().setUp()
        if not self.config.libreoffice:
            self.skipTest("LibreOffice is not installed here; these requests are refused with engine_unavailable instead")

    def test_a_document_converts_to_pdf(self):
        result = self.ok(self.service.convert(str(make_docx(self.inputs)), "pdf"))
        self.assertIn("Quarterly report", PdfReader(result["path"]).pages[0].extract_text())

    def test_recalculate_writes_fresh_formula_results(self):
        path = make_xlsx(self.inputs)
        result = self.ok(self.edit(path, [{"action": "set_cells", "target": "Budget!B2", "values": [[20]]}], recalculate=True))
        self.assertFalse(result["recalculationNeeded"])
        values = openpyxl.load_workbook(result["path"], data_only=True)
        self.assertEqual(values["Budget"]["D2"].value, 50)


class WithoutLibreOfficeTests(Case):
    def test_requests_that_need_it_are_refused_when_it_is_missing(self):
        self.config.libreoffice = None
        path = make_xlsx(self.inputs)
        self.refused(self.service.convert(str(path), "pdf"), "engine_unavailable")
        self.refused(self.edit(path, [{"action": "set_cells", "target": "Budget!B2", "values": [[1]]}], recalculate=True), "engine_unavailable")


if __name__ == "__main__":
    unittest.main()
