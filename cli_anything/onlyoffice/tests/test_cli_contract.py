import io
import json
import os
import tempfile
import unittest
from contextlib import redirect_stdout
from unittest import mock

from cli_anything.onlyoffice.core import cli as cli_module
from cli_anything.onlyoffice.core.command_registry import (
    CLI_SCHEMA_VERSION,
    TOTAL_COMMANDS,
    command_signature,
    command_usage,
)


class OnlyOfficeCLIContractTests(unittest.TestCase):
    def setUp(self):
        self.tmpdir = tempfile.TemporaryDirectory(prefix="oo-cli-contract-test-")
        self.base = self.tmpdir.name

    def tearDown(self):
        self.tmpdir.cleanup()

    def _path(self, name: str) -> str:
        return os.path.join(self.base, name)

    def _run_cli_with_status(self, argv, *, exit_on_error=False):
        stdout = io.StringIO()
        exit_code = 0
        with mock.patch("sys.argv", ["cli-anything-onlyoffice", *argv]):
            with redirect_stdout(stdout):
                try:
                    result = cli_module.main(exit_on_error=exit_on_error)
                except SystemExit as exc:
                    exit_code = exc.code if isinstance(exc.code, int) else 1
                else:
                    exit_code = result if isinstance(result, int) else 0
        return json.loads(stdout.getvalue()), exit_code

    def _run_cli(self, argv):
        payload, _exit_code = self._run_cli_with_status(argv)
        return payload

    def test_top_level_unknown_option_does_not_return_help_success(self):
        payload, exit_code = self._run_cli_with_status(
            ["--bogus", "--json"],
            exit_on_error=True,
        )

        self.assertNotEqual(exit_code, 0)
        self.assertFalse(payload["success"])
        self.assertIn("Unknown global option", payload["error"])
        self.assertNotIn("categories", payload)

    def test_unknown_command_returns_nonzero_exit_status(self):
        payload, exit_code = self._run_cli_with_status(
            ["not-a-real-command", "--json"],
            exit_on_error=True,
        )

        self.assertNotEqual(exit_code, 0)
        self.assertFalse(payload["success"])
        self.assertIn("Unknown command", payload["error"])

    def test_usage_error_returns_nonzero_exit_status(self):
        doc_path = self._path("preview.docx")
        output_dir = self._path("preview")

        with mock.patch.object(cli_module.doc_server, "preview_document") as preview:
            payload, exit_code = self._run_cli_with_status(
                ["doc-preview", doc_path, output_dir, "--dpi", "--json"],
                exit_on_error=True,
            )

        self.assertNotEqual(exit_code, 0)
        self.assertFalse(payload["success"])
        self.assertEqual(payload["error_code"], "usage_error")
        self.assertIn("--dpi", payload["error"])
        preview.assert_not_called()

    def test_missing_file_returns_nonzero_exit_status(self):
        missing_path = self._path("missing.docx")

        payload, exit_code = self._run_cli_with_status(
            ["open", missing_path, "--json"],
            exit_on_error=True,
        )

        self.assertNotEqual(exit_code, 0)
        self.assertFalse(payload["success"])
        self.assertIn("File not found", payload["error"])

    def test_help_json_exposes_stable_schema_and_validation_usage(self):
        payload = self._run_cli(["help", "--json"])

        self.assertTrue(payload["success"])
        self.assertEqual(payload["schema_version"], CLI_SCHEMA_VERSION)
        self.assertEqual(payload["command_count"], TOTAL_COMMANDS)
        self.assertIn("categories", payload)
        self.assertIn("commands", payload)
        self.assertIn("usage", payload)
        self.assertIn("capability_metadata", payload)
        self.assertEqual(
            payload["commands"]["xlsx-add-validation"]["usage"],
            command_usage("xlsx-add-validation"),
        )
        validation_usage = payload["usage"]["xlsx-add-validation"]
        self.assertIn("--no-blank", validation_usage)
        self.assertNotIn("--allow-blank", validation_usage)
        validation_signature = next(
            signature
            for signature in payload["categories"]["SPREADSHEETS (.xlsx)"]
            if signature.startswith("xlsx-add-validation ")
        )
        self.assertIn("--no-blank", validation_signature)
        self.assertNotIn("--allow-blank", validation_signature)

    def test_help_category_signatures_match_command_usage_for_overrides(self):
        payload = self._run_cli(["help", "--json"])

        self.assertTrue(payload["success"])
        all_signatures = [
            signature
            for signatures in payload["categories"].values()
            for signature in signatures
        ]
        for command in [
            "doc-sanitize",
            "pdf-sanitize",
            "rdf-add",
            "rdf-remove",
            "rdf-namespace",
        ]:
            category_signature = next(
                signature
                for signature in all_signatures
                if signature.split()[0] == command
            )
            self.assertEqual(category_signature, command_signature(command))

    def test_status_json_reports_registry_and_capability_metadata(self):
        with mock.patch.object(cli_module.doc_server, "check_health", return_value=True):
            payload = self._run_cli(["status", "--json"])

        self.assertTrue(payload["success"])
        self.assertEqual(payload["schema_version"], CLI_SCHEMA_VERSION)
        self.assertEqual(payload["command_count"], TOTAL_COMMANDS)
        self.assertEqual(payload["registry"]["total_commands"], TOTAL_COMMANDS)
        self.assertEqual(payload["registry"]["schema_version"], CLI_SCHEMA_VERSION)
        self.assertEqual(payload["registry"]["category_counts"], payload["category_counts"])
        self.assertEqual(payload["dependencies"]["openpyxl"], payload["openpyxl"])
        self.assertEqual(
            payload["capability_metadata"]["openpyxl"]["available"],
            payload["openpyxl"],
        )
        self.assertIn("xlsx-*", payload["capability_metadata"]["openpyxl"]["commands"])
        self.assertEqual(
            payload["capability_metadata"]["xlsx_create"]["requires"],
            ["openpyxl"],
        )

    def test_trailing_json_flag_does_not_strip_literal_mid_content(self):
        path = self._path("literal-json.docx")

        with mock.patch.object(cli_module, "DOCX_AVAILABLE", True):
            with mock.patch.object(
                cli_module.doc_server,
                "create_document",
                return_value={"success": True, "file": path},
            ) as create_document:
                payload = self._run_cli(
                    [
                        "doc-create",
                        path,
                        "Title",
                        "literal",
                        "--json",
                        "content",
                        "--json",
                    ]
                )

        self.assertTrue(payload["success"])
        create_document.assert_called_once_with(path, "Title", "literal --json content")

    def test_delimiter_preserves_literal_trailing_json_content(self):
        path = self._path("literal-trailing-json.docx")

        with mock.patch.object(cli_module, "DOCX_AVAILABLE", True):
            with mock.patch.object(
                cli_module.doc_server,
                "create_document",
                return_value={"success": True, "file": path},
            ) as create_document:
                payload = self._run_cli(
                    [
                        "doc-create",
                        path,
                        "Title",
                        "--json",
                        "--",
                        "literal",
                        "--json",
                    ]
                )

        self.assertTrue(payload["success"])
        create_document.assert_called_once_with(path, "Title", "literal --json")

    def test_pdf_page_to_image_bad_dpi_returns_json_error(self):
        pdf_path = self._path("input.pdf")
        output_dir = self._path("pages")

        with mock.patch.object(cli_module.doc_server, "pdf_page_to_image") as page_to_image:
            payload = self._run_cli(
                [
                    "pdf-page-to-image",
                    pdf_path,
                    output_dir,
                    "--dpi",
                    "bad",
                    "--json",
                ]
            )

        self.assertFalse(payload["success"])
        self.assertEqual(payload["error_code"], "usage_error")
        self.assertIn("--dpi", payload["error"])
        page_to_image.assert_not_called()

    def test_extract_render_preview_commands_reject_unknown_flags_and_missing_values(self):
        pdf_path = self._path("input.pdf")
        pptx_path = self._path("deck.pptx")
        output_dir = self._path("out")
        cases = [
            (
                "pdf extract unknown",
                "pdf_extract_images",
                ["pdf-extract-images", pdf_path, output_dir, "--bogus", "--json"],
                "--bogus",
            ),
            (
                "pdf extract missing",
                "pdf_extract_images",
                ["pdf-extract-images", pdf_path, output_dir, "--format", "--json"],
                "--format",
            ),
            (
                "pdf render unknown",
                "pdf_page_to_image",
                ["pdf-page-to-image", pdf_path, output_dir, "--bogus", "--json"],
                "--bogus",
            ),
            (
                "pdf render missing",
                "pdf_page_to_image",
                ["pdf-page-to-image", pdf_path, output_dir, "--dpi", "--json"],
                "--dpi",
            ),
            (
                "pdf read unknown",
                "pdf_read_blocks",
                ["pdf-read-blocks", pdf_path, "--bogus", "--json"],
                "--bogus",
            ),
            (
                "pdf read missing",
                "pdf_read_blocks",
                ["pdf-read-blocks", pdf_path, "--pages", "--json"],
                "--pages",
            ),
            (
                "pdf search unknown",
                "pdf_search_blocks",
                ["pdf-search-blocks", pdf_path, "Results", "--bogus", "--json"],
                "--bogus",
            ),
            (
                "pdf search missing",
                "pdf_search_blocks",
                ["pdf-search-blocks", pdf_path, "Results", "--pages", "--json"],
                "--pages",
            ),
            (
                "pdf sanitize unknown",
                "pdf_sanitize",
                ["pdf-sanitize", pdf_path, "--bogus", "--json"],
                "--bogus",
            ),
            (
                "pdf sanitize missing",
                "pdf_sanitize",
                ["pdf-sanitize", pdf_path, "--author", "--json"],
                "--author",
            ),
            (
                "pptx extract unknown",
                "extract_images_from_pptx",
                ["pptx-extract-images", pptx_path, output_dir, "--bogus", "--json"],
                "--bogus",
            ),
            (
                "pptx extract missing",
                "extract_images_from_pptx",
                ["pptx-extract-images", pptx_path, output_dir, "--prefix", "--json"],
                "--prefix",
            ),
            (
                "pptx preview unknown",
                "preview_slide",
                ["pptx-preview", pptx_path, output_dir, "--bogus", "--json"],
                "--bogus",
            ),
            (
                "pptx preview missing",
                "preview_slide",
                ["pptx-preview", pptx_path, output_dir, "--slide", "--json"],
                "--slide",
            ),
        ]

        for name, method_name, argv, expected_text in cases:
            with self.subTest(name=name):
                with mock.patch.object(cli_module.doc_server, method_name) as operation:
                    payload = self._run_cli(argv)

                self.assertFalse(payload["success"])
                self.assertEqual(payload["error_code"], "usage_error")
                self.assertIn(expected_text, payload["error"])
                operation.assert_not_called()

    def test_doc_render_audit_bad_tolerance_returns_json_error(self):
        doc_path = self._path("audit.docx")

        with mock.patch.object(cli_module, "DOCX_AVAILABLE", True):
            with mock.patch.object(cli_module.doc_server, "rendered_layout_audit") as audit:
                payload = self._run_cli(
                    [
                        "doc-render-audit",
                        doc_path,
                        "--tolerance-points",
                        "bad",
                        "--json",
                    ]
                )

        self.assertFalse(payload["success"])
        self.assertEqual(payload["error_code"], "usage_error")
        self.assertIn("--tolerance-points", payload["error"])
        audit.assert_not_called()

    def test_non_finite_float_options_return_json_usage_error(self):
        doc_path = self._path("layout.docx")

        with mock.patch.object(cli_module, "DOCX_AVAILABLE", True):
            with mock.patch.object(cli_module.doc_server, "set_page_layout") as layout:
                payload = self._run_cli(
                    [
                        "doc-layout",
                        doc_path,
                        "--margin-left",
                        "nan",
                        "--json",
                    ]
                )

        self.assertFalse(payload["success"])
        self.assertEqual(payload["error_code"], "usage_error")
        self.assertIn("finite number", payload["error"])
        layout.assert_not_called()

    def test_doc_formatting_info_bad_start_returns_json_error(self):
        doc_path = self._path("formatting.docx")

        with mock.patch.object(cli_module, "DOCX_AVAILABLE", True):
            with mock.patch.object(cli_module.doc_server, "get_formatting_info") as formatting:
                payload = self._run_cli(
                    [
                        "doc-formatting-info",
                        doc_path,
                        "--start",
                        "bad",
                        "--json",
                    ]
                )

        self.assertFalse(payload["success"])
        self.assertEqual(payload["error_code"], "usage_error")
        self.assertIn("--start", payload["error"])
        formatting.assert_not_called()

    def test_xlsx_preview_bad_dpi_returns_json_usage_error(self):
        path = self._path("preview.xlsx")
        output_dir = self._path("preview")

        with mock.patch.object(cli_module.doc_server, "preview_spreadsheet") as preview:
            payload = self._run_cli(
                [
                    "xlsx-preview",
                    path,
                    output_dir,
                    "--dpi",
                    "bad",
                    "--json",
                ]
            )

        self.assertFalse(payload["success"])
        self.assertEqual(payload["error_code"], "usage_error")
        self.assertIn("--dpi", payload["error"])
        preview.assert_not_called()

    def test_chart_comparison_bad_value_cols_returns_json_usage_error(self):
        path = self._path("chart.xlsx")

        with mock.patch.object(cli_module.doc_server, "create_comparison_chart") as chart:
            payload = self._run_cli(
                [
                    "chart-comparison",
                    path,
                    "bar",
                    "Scores",
                    "--value-cols",
                    "2,bad",
                    "--json",
                ]
            )

        self.assertFalse(payload["success"])
        self.assertEqual(payload["error_code"], "usage_error")
        self.assertIn("--value-cols", payload["error"])
        chart.assert_not_called()

    def test_pptx_add_textbox_bad_left_returns_json_usage_error(self):
        path = self._path("deck.pptx")

        with mock.patch.object(cli_module.doc_server, "add_textbox") as add_textbox:
            payload = self._run_cli(
                [
                    "pptx-add-textbox",
                    path,
                    "0",
                    "Callout",
                    "--left",
                    "bad",
                    "--json",
                ]
            )

        self.assertFalse(payload["success"])
        self.assertEqual(payload["error_code"], "usage_error")
        self.assertIn("--left", payload["error"])
        add_textbox.assert_not_called()

    def test_editor_capture_bad_zoom_returns_json_usage_error(self):
        path = self._path("capture.xlsx")
        output_path = self._path("capture.png")

        with mock.patch.object(cli_module.doc_server, "capture_editor_view") as capture:
            payload = self._run_cli(
                [
                    "editor-capture",
                    path,
                    output_path,
                    "--zoom-in",
                    "bad",
                    "--json",
                ]
            )

        self.assertFalse(payload["success"])
        self.assertEqual(payload["error_code"], "usage_error")
        self.assertIn("--zoom-in", payload["error"])
        capture.assert_not_called()

    def test_rdf_remove_without_selectors_requires_all(self):
        rdf_path = self._path("graph.ttl")
        with open(rdf_path, "w", encoding="utf-8") as handle:
            handle.write("@prefix ex: <http://example.org/> . ex:A ex:p ex:B .\n")

        with mock.patch.object(cli_module.doc_server, "rdf_remove") as rdf_remove:
            payload = self._run_cli(["rdf-remove", rdf_path, "--json"])

        self.assertFalse(payload["success"])
        self.assertEqual(payload["error_code"], "usage_error")
        self.assertIn("selector", payload["error"])
        self.assertIn("--all", payload["usage"])
        rdf_remove.assert_not_called()

    def test_rdf_remove_dry_run_reports_without_mutating(self):
        rdf_path = self._path("dry-run.ttl")
        content = (
            "@prefix ex: <http://example.org/> .\n"
            "ex:A ex:p ex:B .\n"
            "ex:C ex:p ex:D .\n"
        )
        with open(rdf_path, "w", encoding="utf-8") as handle:
            handle.write(content)

        with mock.patch.object(cli_module.doc_server, "rdf_remove") as rdf_remove:
            payload = self._run_cli(
                [
                    "rdf-remove",
                    rdf_path,
                    "--predicate",
                    "http://example.org/p",
                    "--dry-run",
                    "--json",
                ]
            )

        self.assertTrue(payload["success"])
        self.assertTrue(payload["dry_run"])
        self.assertEqual(payload["would_remove"], 2)
        self.assertEqual(payload["removed"], 0)
        rdf_remove.assert_not_called()
        with open(rdf_path, "r", encoding="utf-8") as handle:
            self.assertEqual(handle.read(), content)

    def test_rdf_add_dispatches_without_remove_options(self):
        rdf_path = self._path("add.ttl")

        with mock.patch.object(
            cli_module.doc_server,
            "rdf_add",
            return_value={"success": True, "triples": 1},
        ) as rdf_add:
            payload = self._run_cli(
                [
                    "rdf-add",
                    rdf_path,
                    "http://example.org/s",
                    "http://example.org/p",
                    "label",
                    "--type",
                    "literal",
                    "--lang",
                    "en",
                    "--format",
                    "turtle",
                    "--json",
                ]
            )

        self.assertTrue(payload["success"])
        rdf_add.assert_called_once_with(
            rdf_path,
            "http://example.org/s",
            "http://example.org/p",
            "label",
            object_type="literal",
            lang="en",
            datatype=None,
            format="turtle",
        )

    def test_rdf_add_infers_jsonld_format_without_format_option(self):
        rdf_path = self._path("graph.jsonld")

        create_payload = self._run_cli(
            ["rdf-create", rdf_path, "--format", "json-ld", "--json"]
        )
        self.assertTrue(create_payload["success"])

        add_payload = self._run_cli(
            [
                "rdf-add",
                rdf_path,
                "http://example.org/s",
                "http://example.org/p",
                "label",
                "--type",
                "literal",
                "--lang",
                "en",
                "--json",
            ]
        )
        self.assertTrue(add_payload["success"], add_payload)

        with open(rdf_path, "r", encoding="utf-8") as handle:
            json.load(handle)

        read_payload = self._run_cli(["rdf-read", rdf_path, "--json"])
        self.assertTrue(read_payload["success"], read_payload)
        self.assertEqual(read_payload["total_triples"], 1)


if __name__ == "__main__":
    unittest.main()
