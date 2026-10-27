"""The catalog is complete and consistent, and the generated files match it."""

import unittest

from jsonschema import Draft202012Validator

from officemcp import catalog, contracts
from officemcp.errors import CODES
from officemcp.tools import TOOLS


class ContractTests(unittest.TestCase):
    def test_generated_catalog_and_guide_are_current(self):
        self.assertEqual(contracts.write(check=True), [], "run: python -m officemcp contracts")

    def test_every_action_has_a_valid_schema_family_and_summary(self):
        names = set()
        for a in catalog.actions():
            Draft202012Validator.check_schema(a.schema())
            self.assertTrue(a.summary.endswith("."), a.name)
            self.assertIn(a.format, ("docx", "xlsx", "pptx", "pdf"))
            self.assertIn(a.tier, ("primitive", "recipe"))
            self.assertNotIn((a.format, a.name), names)
            names.add((a.format, a.name))
            for field in a.required:
                self.assertIn(field, a.params, f"{a.format} {a.name}: required {field} is not a field")
        for t in TOOLS:
            Draft202012Validator.check_schema(t["inputSchema"])
        self.assertEqual([t["name"] for t in TOOLS], ["office_actions", "office_inspect", "office_read", "office_find",
                                                     "office_edit", "office_create", "office_convert"])

    def test_the_guide_names_every_action_and_refusal_code(self):
        guide = contracts.guide_markdown()
        for a in catalog.actions():
            self.assertIn(f"`{a.name}`", guide)
        for code in CODES:
            self.assertIn(f"`{code}`", guide)


if __name__ == "__main__":
    unittest.main()
