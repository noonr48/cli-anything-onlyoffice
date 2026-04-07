import json
import os
import subprocess
import tempfile
import unittest

from openpyxl import Workbook, load_workbook

from cli_anything.onlyoffice.utils.docserver import get_client


class OnlyOfficeProductionReadinessTests(unittest.TestCase):
    def setUp(self):
        self.client = get_client()
        self.tmpdir = tempfile.TemporaryDirectory(prefix="oo-prod-test-")
        self.base = self.tmpdir.name

    def tearDown(self):
        self.tmpdir.cleanup()

    def _path(self, name: str) -> str:
        return os.path.join(self.base, name)

    def test_xlsx_write_is_non_destructive_by_default(self):
        path = self._path("grades.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws["A1"] = "Header"
        keep = wb.create_sheet("KeepSheet")
        keep["A1"] = "KeepMe"
        wb.save(path)

        result = self.client.write_spreadsheet(
            output_path=path,
            headers=["Student", "Score"],
            data=[["Alice", 91], ["Bob", 88]],
            sheet_name="Sheet1",
            overwrite_workbook=False,
        )
        self.assertTrue(result["success"])

        wb2 = load_workbook(path)
        self.assertIn("KeepSheet", wb2.sheetnames)
        self.assertEqual(wb2["KeepSheet"]["A1"].value, "KeepMe")

    def test_xlsx_write_can_explicitly_overwrite(self):
        path = self._path("grades_overwrite.xlsx")
        wb = Workbook()
        wb.active.title = "Legacy"
        wb.create_sheet("KeepSheet")
        wb.save(path)

        result = self.client.write_spreadsheet(
            output_path=path,
            headers=["Student", "Score"],
            data=[["Alice", 91]],
            sheet_name="NewSheet",
            overwrite_workbook=True,
        )
        self.assertTrue(result["success"])

        wb2 = load_workbook(path)
        self.assertEqual(wb2.sheetnames, ["NewSheet"])

    def test_formula_aware_calculation(self):
        path = self._path("formula_calc.xlsx")
        self.client.write_spreadsheet(
            output_path=path,
            headers=["A", "B", "Total"],
            data=[[10, 20, "=A2+B2"], [15, 25, "=A3+B3"]],
            sheet_name="Sheet1",
            overwrite_workbook=True,
        )

        raw = self.client.calculate_column(
            path, "C", "avg", sheet_name="Sheet1", include_formulas=False
        )
        self.assertFalse(raw["success"])

        formula = self.client.calculate_column(
            path, "C", "avg", sheet_name="Sheet1", include_formulas=True
        )
        self.assertTrue(formula["success"])
        self.assertEqual(formula["average"], 35.0)

    def test_chart_validation_rejects_invalid_ranges(self):
        path = self._path("chart_invalid.xlsx")
        self.client.write_spreadsheet(
            output_path=path,
            headers=["Student", "Score"],
            data=[["Alice", 90], ["Bob", 85]],
            sheet_name="Sheet1",
            overwrite_workbook=True,
        )

        result = self.client.create_chart(
            file_path=path,
            chart_type="bar",
            data_range="Z1:Z5",
            categories_range="A2:A3",
            title="Invalid",
        )
        self.assertFalse(result["success"])

    def test_backup_snapshot_is_created_for_mutation(self):
        path = self._path("paper.docx")
        create = self.client.create_document(path, "Title", "Intro")
        self.assertTrue(create["success"])

        append = self.client.append_to_document(path, "New body paragraph")
        self.assertTrue(append["success"])
        self.assertTrue(append.get("backup"))
        self.assertTrue(os.path.exists(append["backup"]))

    def test_cli_help_exposes_hardened_commands(self):
        proc = subprocess.run(
            ["cli-anything-onlyoffice", "help", "--json"],
            capture_output=True,
            text=True,
            check=True,
        )
        payload = json.loads(proc.stdout)
        docs = payload["categories"]["DOCUMENTS (.docx)"]
        sheet = payload["categories"]["SPREADSHEETS (.xlsx)"]
        self.assertIn(
            "doc-format <file> <paragraph_index> [--bold] [--italic] [--underline] [--font-name <name>] [--font-size <n>] [--color <hex>] [--align <left|center|right|justify>]",
            docs,
        )
        self.assertIn(
            "xlsx-calc <file> <column> <op> [--sheet <name>] [--include-formulas] [--strict-formulas]",
            sheet,
        )

    def test_strict_schema_validation_rejects_row_width_mismatch(self):
        path = self._path("schema_strict.xlsx")
        result = self.client.write_spreadsheet(
            output_path=path,
            headers=["A", "B", "C"],
            data=[[1, 2], [3, 4, 5]],
            sheet_name="Sheet1",
            overwrite_workbook=True,
            coerce_rows=False,
        )
        self.assertFalse(result["success"])
        self.assertIn("expected 3", result["error"])

    def test_coerce_rows_allows_width_normalization(self):
        path = self._path("schema_coerce.xlsx")
        result = self.client.write_spreadsheet(
            output_path=path,
            headers=["A", "B", "C"],
            data=[[1, 2], [3, 4, 5, 6]],
            sheet_name="Sheet1",
            overwrite_workbook=True,
            coerce_rows=True,
        )
        self.assertTrue(result["success"])
        wb = load_workbook(path)
        ws = wb["Sheet1"]
        self.assertIn(ws["C2"].value, ("", None))
        self.assertEqual(ws["C3"].value, 5)

    def test_backup_list_restore_and_prune(self):
        path = self._path("restore_target.docx")
        self.client.create_document(path, "Title", "Version 1")
        self.client.append_to_document(path, "Version 2")
        self.client.append_to_document(path, "Version 3")

        listing = self.client.list_backups(path, limit=10)
        self.assertTrue(listing["success"])
        self.assertGreaterEqual(listing["count"], 1)

        before_restore = self.client.read_document(path)
        self.assertIn("Version 3", before_restore["full_text"])

        restore = self.client.restore_backup(
            path, backup=listing["backups"][-1]["name"]
        )
        self.assertTrue(restore["success"])
        after_restore = self.client.read_document(path)
        self.assertNotIn("Version 3", after_restore["full_text"])

        pruned = self.client.prune_backups(file_path=path, keep=1)
        self.assertTrue(pruned["success"])
        listing2 = self.client.list_backups(path, limit=10)
        self.assertLessEqual(listing2["count"], 1)


if __name__ == "__main__":
    unittest.main()
