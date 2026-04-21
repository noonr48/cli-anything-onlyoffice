import os
import tempfile
import unittest
from unittest import mock

from openpyxl import Workbook, load_workbook

from cli_anything.onlyoffice.utils.docserver import get_client


class OnlyOfficeXLSXOpsTests(unittest.TestCase):
    def setUp(self):
        self.client = get_client()
        self.ops = self.client._xlsx_ops
        self.tmpdir = tempfile.TemporaryDirectory(prefix="oo-xlsx-ops-test-")
        self.base = self.tmpdir.name

    def tearDown(self):
        self.tmpdir.cleanup()

    def _path(self, name: str) -> str:
        return os.path.join(self.base, name)

    def test_xlsx_ops_create_write_and_sheet_list_roundtrip(self):
        path = self._path("roundtrip.xlsx")

        create_result = self.ops.create_spreadsheet(path, sheet_name="Data")
        write_result = self.ops.write_spreadsheet(
            path,
            ["Name", "Score"],
            [["Alice", "90"], ["Bob", "85"]],
            sheet_name="Data",
        )
        list_result = self.ops.sheet_list(path)

        self.assertTrue(create_result["success"])
        self.assertTrue(write_result["success"])
        self.assertTrue(list_result["success"])
        self.assertEqual(list_result["sheet_count"], 1)
        self.assertEqual(list_result["sheets"][0]["name"], "Data")

        wb = load_workbook(path)
        ws = wb["Data"]
        self.assertEqual(ws["A2"].value, "Alice")
        self.assertEqual(ws["B2"].value, 90)
        self.assertEqual(ws["A3"].value, "Bob")
        self.assertEqual(ws["B3"].value, 85)
        wb.close()

    def test_xlsx_ops_formula_audit_detects_unsupported_functions(self):
        path = self._path("formula_audit.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = "Audit"
        ws["A1"] = 10
        ws["A2"] = 20
        ws["B1"] = "=SUM(A1:A2)"
        ws["B2"] = "=XLOOKUP(A1,A1:A2,A1:A2)"
        wb.save(path)
        wb.close()

        result = self.ops.audit_spreadsheet_formulas(path, sheet_name="Audit")

        self.assertTrue(result["success"])
        self.assertEqual(result["formula_count"], 2)
        self.assertIn("XLOOKUP", result["unsupported_functions"])
        self.assertFalse(result["safe_for_cli_formula_eval"])
        self.assertGreaterEqual(len(result["risks"]), 1)

    def test_xlsx_ops_preview_uses_host_spreadsheet_pdf_pipeline(self):
        path = self._path("preview.xlsx")
        output_dir = self._path("preview-out")
        Workbook().save(path)
        captured = {}

        def fake_spreadsheet_to_pdf(file_path, output_path=None):
            self.assertEqual(file_path, path)
            self.assertIsNotNone(output_path)
            with open(output_path, "wb") as handle:
                handle.write(b"%PDF-1.4 fake")
            captured["pdf_path"] = output_path
            return {
                "success": True,
                "input_file": file_path,
                "output_file": output_path,
                "pages": 1,
            }

        def fake_pdf_page_to_image(file_path, render_dir, pages=None, dpi=150, fmt="png"):
            self.assertEqual(file_path, captured["pdf_path"])
            self.assertEqual(render_dir, output_dir)
            self.assertEqual(pages, "0")
            self.assertEqual(dpi, 200)
            self.assertEqual(fmt, "jpg")
            return {
                "success": True,
                "total_pages": 1,
                "pages_rendered": 1,
                "images": [{"page": 0, "file": os.path.join(render_dir, "page_000.jpg")}],
            }

        with mock.patch.object(
            self.client, "spreadsheet_to_pdf", side_effect=fake_spreadsheet_to_pdf
        ):
            with mock.patch.object(
                self.client, "pdf_page_to_image", side_effect=fake_pdf_page_to_image
            ):
                result = self.ops.preview_spreadsheet(
                    path, output_dir, pages="0", dpi=200, fmt="jpg"
                )

        self.assertTrue(result["success"])
        self.assertEqual(result["pages_rendered"], 1)
        self.assertEqual(result["format"], "jpg")
        self.assertFalse(os.path.exists(captured["pdf_path"]))
