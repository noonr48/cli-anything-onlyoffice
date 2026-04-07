import os
import tempfile
import unittest

from cli_anything.onlyoffice.utils.docserver import get_client


class OnlyOfficeFormulaSafetyTests(unittest.TestCase):
    def setUp(self):
        self.client = get_client()
        self.tmp = tempfile.TemporaryDirectory(prefix="oo-formula-")
        self.path = os.path.join(self.tmp.name, "formula_safety.xlsx")

        self.client.write_spreadsheet(
            output_path=self.path,
            headers=["A", "B", "Calc"],
            data=[
                [1, 2, "=A2+B2"],
                [3, 4, "=IF(A3>0,B3,0)"],
                [5, 6, "=[Book2.xlsx]Sheet1!A1"],
            ],
            sheet_name="Sheet0",
            overwrite_workbook=True,
        )

    def tearDown(self):
        self.tmp.cleanup()

    def test_formula_audit_detects_risks(self):
        audit = self.client.audit_spreadsheet_formulas(self.path, sheet_name="Sheet0")
        self.assertTrue(audit["success"])
        self.assertGreaterEqual(audit["formula_count"], 3)
        self.assertIn("IF", audit["unsupported_functions"])
        self.assertFalse(audit["safe_for_cli_formula_eval"])

    def test_strict_formula_mode_blocks_unreliable_eval(self):
        calc = self.client.calculate_column(
            self.path,
            "C",
            "avg",
            sheet_name="Sheet0",
            include_formulas=True,
            strict_formula_safety=True,
        )
        self.assertFalse(calc["success"])
        self.assertIn("Strict formula safety", calc["error"])

    def test_inferential_outputs_include_interpretation_and_apa(self):
        # Build simple inferential-safe dataset
        p2 = os.path.join(self.tmp.name, "infer.xlsx")
        self.client.write_spreadsheet(
            output_path=p2,
            headers=["G", "X", "Y", "C"],
            data=[[1, 10, 20, 1], [1, 11, 22, 1], [2, 20, 40, 2], [2, 21, 42, 2]],
            sheet_name="Sheet0",
            overwrite_workbook=True,
        )
        corr = self.client.correlation_test(p2, "B", "C", sheet_name="Sheet0")
        self.assertTrue(corr["success"])
        self.assertIn("interpretation", corr)
        self.assertIn("apa", corr)

        ttest = self.client.ttest_independent(
            p2, "B", "A", "1", "2", sheet_name="Sheet0"
        )
        self.assertTrue(ttest["success"])
        self.assertIn("interpretation", ttest)
        self.assertIn("apa", ttest)

        chi = self.client.chi_square_test(p2, "A", "D", sheet_name="Sheet0")
        self.assertTrue(chi["success"])
        self.assertIn("interpretation", chi)
        self.assertIn("apa", chi)


if __name__ == "__main__":
    unittest.main()
