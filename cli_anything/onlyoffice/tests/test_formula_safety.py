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
                [1, 2, ""],
                [3, 4, ""],
                [5, 6, ""],
            ],
            sheet_name="Sheet0",
            overwrite_workbook=True,
        )
        self.client.add_formula(self.path, "C2", "=A2+B2", sheet_name="Sheet0")
        self.client.add_formula(self.path, "C3", "=IF(A3>0,B3,0)", sheet_name="Sheet0")
        self.client.add_formula(
            self.path, "C4", "=[Book2.xlsx]Sheet1!A1", sheet_name="Sheet0"
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

    def test_bulk_write_neutralizes_formulas_by_default(self):
        p2 = os.path.join(self.tmp.name, "neutralized.xlsx")
        result = self.client.write_spreadsheet(
            output_path=p2,
            headers=["Name", "FormulaLike"],
            data=[["Alice", "=1+1"], ["Bob", "@SUM(A1:A2)"]],
            sheet_name="Sheet0",
            overwrite_workbook=True,
        )

        self.assertTrue(result["success"])
        self.assertEqual(result["neutralized_cells"], 2)

        audit = self.client.audit_spreadsheet_formulas(p2, sheet_name="Sheet0")
        self.assertTrue(audit["success"])
        self.assertEqual(audit["formula_count"], 0)

    def test_formula_evaluator_rejects_exponent_bombs(self):
        p2 = os.path.join(self.tmp.name, "exponent_bomb.xlsx")
        self.client.write_spreadsheet(
            output_path=p2,
            headers=["A", "B", "Calc"],
            data=[[9, 2, ""]],
            sheet_name="Sheet0",
            overwrite_workbook=True,
        )
        self.assertTrue(
            self.client.add_formula(p2, "C2", "=9**99999999", sheet_name="Sheet0")[
                "success"
            ]
        )

        calc = self.client.calculate_column(
            p2, "C", "avg", sheet_name="Sheet0", include_formulas=True
        )

        self.assertFalse(calc["success"])
        self.assertIn("No numeric values", calc["error"])

    def test_formula_evaluator_handles_basic_arithmetic(self):
        p2 = os.path.join(self.tmp.name, "basic_formula.xlsx")
        self.client.write_spreadsheet(
            output_path=p2,
            headers=["A", "B", "Calc"],
            data=[[9, 3, ""], [4, 2, ""]],
            sheet_name="Sheet0",
            overwrite_workbook=True,
        )
        self.assertTrue(
            self.client.add_formula(p2, "C2", "=(A2+B2)/B2", sheet_name="Sheet0")[
                "success"
            ]
        )
        self.assertTrue(
            self.client.add_formula(p2, "C3", "=SUM(A3:B3)", sheet_name="Sheet0")[
                "success"
            ]
        )

        calc = self.client.calculate_column(
            p2, "C", "sum", sheet_name="Sheet0", include_formulas=True
        )

        self.assertTrue(calc["success"])
        self.assertEqual(calc["sum"], 10.0)

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
