import io
import json
import os
import tempfile
import unittest
from contextlib import redirect_stdout
from unittest import mock

from cli_anything.onlyoffice.core import cli as cli_module


class OnlyOfficeXLSXCLITests(unittest.TestCase):
    def setUp(self):
        self.tmpdir = tempfile.TemporaryDirectory(prefix="oo-xlsx-cli-test-")
        self.base = self.tmpdir.name

    def tearDown(self):
        self.tmpdir.cleanup()

    def _path(self, name: str) -> str:
        return os.path.join(self.base, name)

    def test_cli_xlsx_write_dispatches_via_xlsx_handler(self):
        path = self._path("dispatch_write.xlsx")
        stdout = io.StringIO()

        with mock.patch.object(
            cli_module.doc_server,
            "write_spreadsheet",
            return_value={"success": True, "file": path, "rows_written": 2},
        ) as write_spreadsheet:
            with mock.patch(
                "sys.argv",
                [
                    "cli-anything-onlyoffice",
                    "xlsx-write",
                    path,
                    "Name,Score",
                    "Alice,90;Bob,85",
                    "--sheet",
                    "Grades",
                    "--overwrite",
                    "--coerce-rows",
                    "--text-columns",
                    "A,C",
                    "--json",
                ],
            ):
                with redirect_stdout(stdout):
                    cli_module.main()

        payload = json.loads(stdout.getvalue())
        self.assertTrue(payload["success"])
        write_spreadsheet.assert_called_once_with(
            path,
            ["Name", "Score"],
            [["Alice", "90"], ["Bob", "85"]],
            sheet_name="Grades",
            overwrite_workbook=True,
            coerce_rows=True,
            text_columns=["A", "C"],
        )

    def test_cli_xlsx_calc_dispatches_via_xlsx_handler(self):
        path = self._path("dispatch_calc.xlsx")
        stdout = io.StringIO()

        with mock.patch.object(
            cli_module.doc_server,
            "calculate_column",
            return_value={"success": True, "average": 91.5},
        ) as calculate_column:
            with mock.patch(
                "sys.argv",
                [
                    "cli-anything-onlyoffice",
                    "xlsx-calc",
                    path,
                    "B",
                    "avg",
                    "--sheet",
                    "Grades",
                    "--include-formulas",
                    "--strict-formulas",
                    "--json",
                ],
            ):
                with redirect_stdout(stdout):
                    cli_module.main()

        payload = json.loads(stdout.getvalue())
        self.assertTrue(payload["success"])
        calculate_column.assert_called_once_with(
            path,
            "B",
            "avg",
            sheet_name="Grades",
            include_formulas=True,
            strict_formula_safety=True,
        )

    def test_cli_xlsx_formula_dispatches_explicit_formula_entrypoint(self):
        path = self._path("dispatch_formula.xlsx")
        stdout = io.StringIO()

        with mock.patch.object(
            cli_module.doc_server,
            "add_formula",
            return_value={"success": True, "cell": "C2", "formula": "=A2+B2"},
        ) as add_formula:
            with mock.patch(
                "sys.argv",
                [
                    "cli-anything-onlyoffice",
                    "xlsx-formula",
                    path,
                    "C2",
                    "=A2+B2",
                    "--sheet",
                    "Grades",
                    "--json",
                ],
            ):
                with redirect_stdout(stdout):
                    cli_module.main()

        payload = json.loads(stdout.getvalue())
        self.assertTrue(payload["success"])
        add_formula.assert_called_once_with(path, "C2", "=A2+B2", sheet_name="Grades")

    def test_cli_xlsx_add_validation_dispatches_via_xlsx_handler(self):
        path = self._path("dispatch_validate.xlsx")
        stdout = io.StringIO()

        with mock.patch.object(
            cli_module.doc_server,
            "add_validation",
            return_value={"success": True, "range": "A2:A10"},
        ) as add_validation:
            with mock.patch(
                "sys.argv",
                [
                    "cli-anything-onlyoffice",
                    "xlsx-add-validation",
                    path,
                    "A2:A10",
                    "whole",
                    "--operator",
                    "between",
                    "--formula1",
                    "1",
                    "--formula2",
                    "5",
                    "--sheet",
                    "Grades",
                    "--error",
                    "Enter 1-5",
                    "--prompt",
                    "Choose a score",
                    "--error-style",
                    "warning",
                    "--no-blank",
                    "--json",
                ],
            ):
                with redirect_stdout(stdout):
                    cli_module.main()

        payload = json.loads(stdout.getvalue())
        self.assertTrue(payload["success"])
        add_validation.assert_called_once_with(
            path,
            "A2:A10",
            "whole",
            operator="between",
            formula1="1",
            formula2="5",
            allow_blank=False,
            sheet_name="Grades",
            error_message="Enter 1-5",
            error_title=None,
            prompt_message="Choose a score",
            prompt_title=None,
            error_style="warning",
        )

    def test_cli_xlsx_preview_dispatches_via_xlsx_handler(self):
        path = self._path("dispatch_preview.xlsx")
        out_dir = self._path("preview")
        stdout = io.StringIO()

        with mock.patch.object(
            cli_module.doc_server,
            "preview_spreadsheet",
            return_value={"success": True, "pages_rendered": 1},
        ) as preview_spreadsheet:
            with mock.patch(
                "sys.argv",
                [
                    "cli-anything-onlyoffice",
                    "xlsx-preview",
                    path,
                    out_dir,
                    "--pages",
                    "0-1",
                    "--dpi",
                    "200",
                    "--format",
                    "jpg",
                    "--json",
                ],
            ):
                with redirect_stdout(stdout):
                    cli_module.main()

        payload = json.loads(stdout.getvalue())
        self.assertTrue(payload["success"])
        preview_spreadsheet.assert_called_once_with(
            path,
            out_dir,
            pages="0-1",
            dpi=200,
            fmt="jpg",
        )

    def test_cli_chart_create_dispatches_via_xlsx_handler(self):
        path = self._path("dispatch_chart.xlsx")
        stdout = io.StringIO()

        with mock.patch.object(
            cli_module.doc_server,
            "create_chart",
            return_value={"success": True, "chart_type": "bar"},
        ) as create_chart:
            with mock.patch(
                "sys.argv",
                [
                    "cli-anything-onlyoffice",
                    "chart-create",
                    path,
                    "bar",
                    "B1:B5",
                    "A1:A5",
                    "Scores",
                    "--sheet",
                    "Grades",
                    "--output-sheet",
                    "Charts",
                    "--x-label",
                    "Students",
                    "--y-label",
                    "Score",
                    "--labels",
                    "--no-legend",
                    "--legend-pos",
                    "bottom",
                    "--colors",
                    "112233,445566",
                    "--json",
                ],
            ):
                with redirect_stdout(stdout):
                    cli_module.main()

        payload = json.loads(stdout.getvalue())
        self.assertTrue(payload["success"])
        create_chart.assert_called_once_with(
            file_path=path,
            chart_type="bar",
            data_range="B1:B5",
            categories_range="A1:A5",
            title="Scores",
            sheet_name="Grades",
            output_sheet="Charts",
            x_label="Students",
            y_label="Score",
            show_data_labels=True,
            show_legend=False,
            legend_pos="bottom",
            colors=["112233", "445566"],
        )

    def test_cli_xlsx_mannwhitney_dispatches_via_xlsx_handler(self):
        path = self._path("dispatch_mw.xlsx")
        stdout = io.StringIO()

        with mock.patch.object(
            cli_module.doc_server,
            "mann_whitney_test",
            return_value={"success": True, "p_value": 0.04},
        ) as mann_whitney_test:
            with mock.patch(
                "sys.argv",
                [
                    "cli-anything-onlyoffice",
                    "xlsx-mannwhitney",
                    path,
                    "B",
                    "C",
                    "Lower",
                    "Higher",
                    "--sheet",
                    "Analysis",
                    "--json",
                ],
            ):
                with redirect_stdout(stdout):
                    cli_module.main()

        payload = json.loads(stdout.getvalue())
        self.assertTrue(payload["success"])
        mann_whitney_test.assert_called_once_with(
            path,
            "B",
            "C",
            "Lower",
            "Higher",
            sheet_name="Analysis",
        )
