import io
import json
import os
import tempfile
import unittest
from contextlib import redirect_stdout
from unittest import mock

from cli_anything.onlyoffice.core import cli as cli_module


class OnlyOfficeDocCLITests(unittest.TestCase):
    def setUp(self):
        self.tmpdir = tempfile.TemporaryDirectory(prefix="oo-doc-cli-test-")
        self.base = self.tmpdir.name

    def tearDown(self):
        self.tmpdir.cleanup()

    def _path(self, name: str) -> str:
        return os.path.join(self.base, name)

    def _run_cli(self, argv):
        stdout = io.StringIO()
        exit_code = 0
        with mock.patch("sys.argv", ["cli-anything-onlyoffice", *argv]):
            with redirect_stdout(stdout):
                try:
                    result = cli_module.main()
                except SystemExit as exc:
                    exit_code = exc.code if isinstance(exc.code, int) else 1
                else:
                    exit_code = result if isinstance(result, int) else 0
        return json.loads(stdout.getvalue()), exit_code

    def test_cli_doc_create_dispatches_via_doc_handler(self):
        path = self._path("dispatch_create.docx")
        stdout = io.StringIO()

        with mock.patch.object(
            cli_module.doc_server,
            "create_document",
            return_value={"success": True, "file": path},
        ) as create_document:
            with mock.patch(
                "sys.argv",
                [
                    "cli-anything-onlyoffice",
                    "doc-create",
                    path,
                    "My Title",
                    "First",
                    "body",
                    "paragraph",
                    "--json",
                ],
            ):
                with redirect_stdout(stdout):
                    cli_module.main()

        payload = json.loads(stdout.getvalue())
        self.assertTrue(payload["success"])
        create_document.assert_called_once_with(
            path,
            "My Title",
            "First body paragraph",
        )

    def test_cli_doc_layout_dispatches_via_doc_handler(self):
        path = self._path("dispatch_layout.docx")
        stdout = io.StringIO()

        with mock.patch.object(
            cli_module.doc_server,
            "set_page_layout",
            return_value={"success": True, "page_size": "A4"},
        ) as set_page_layout:
            with mock.patch(
                "sys.argv",
                [
                    "cli-anything-onlyoffice",
                    "doc-layout",
                    path,
                    "--size",
                    "A4",
                    "--orientation",
                    "landscape",
                    "--margin-top",
                    "0.5",
                    "--margin-left",
                    "0.75",
                    "--header",
                    "Header text",
                    "--page-numbers",
                    "--json",
                ],
            ):
                with redirect_stdout(stdout):
                    cli_module.main()

        payload = json.loads(stdout.getvalue())
        self.assertTrue(payload["success"])
        set_page_layout.assert_called_once_with(
            path,
            page_size="A4",
            orientation="landscape",
            margins={"top": 0.5, "left": 0.75},
            header_text="Header text",
            page_numbers=True,
        )

    def test_cli_doc_sanitize_dispatches_via_doc_handler(self):
        path = self._path("dispatch_sanitize.docx")
        output_path = self._path("dispatch_sanitize_clean.docx")
        stdout = io.StringIO()

        with mock.patch.object(
            cli_module.doc_server,
            "sanitize_document",
            return_value={"success": True, "output_file": output_path},
        ) as sanitize_document:
            with mock.patch(
                "sys.argv",
                [
                    "cli-anything-onlyoffice",
                    "doc-sanitize",
                    path,
                    output_path,
                    "--remove-comments",
                    "--accept-revisions",
                    "--clear-metadata",
                    "--remove-custom-xml",
                    "--set-remove-personal-information",
                    "--canonicalize-ooxml",
                    "--author",
                    "benbi",
                    "--title",
                    "Clean Title",
                    "--subject",
                    "Methods",
                    "--keywords",
                    "audit,submission",
                    "--json",
                ],
            ):
                with redirect_stdout(stdout):
                    cli_module.main()

        payload = json.loads(stdout.getvalue())
        self.assertTrue(payload["success"])
        sanitize_document.assert_called_once_with(
            path,
            output_path=output_path,
            remove_comments=True,
            accept_revisions=True,
            clear_metadata=True,
            remove_custom_xml=True,
            set_remove_personal_information=True,
            canonicalize_ooxml=True,
            author="benbi",
            title="Clean Title",
            subject="Methods",
            keywords="audit,submission",
        )

    def test_cli_doc_sanitize_delimiter_keeps_option_like_output_path_literal(self):
        path = self._path("delimiter_sanitize.docx")
        output_path = "--canonicalize-ooxml"

        with mock.patch.object(
            cli_module.doc_server,
            "sanitize_document",
            return_value={"success": True, "output_file": output_path},
        ) as sanitize_document:
            payload, exit_code = self._run_cli(
                [
                    "doc-sanitize",
                    path,
                    "--json",
                    "--",
                    output_path,
                ]
            )

        self.assertEqual(exit_code, 0)
        self.assertTrue(payload["success"])
        sanitize_document.assert_called_once_with(
            path,
            output_path=output_path,
            remove_comments=False,
            accept_revisions=False,
            clear_metadata=False,
            remove_custom_xml=False,
            set_remove_personal_information=False,
            canonicalize_ooxml=False,
            author=None,
            title=None,
            subject=None,
            keywords=None,
        )

    def test_cli_doc_preview_rejects_unknown_and_missing_options(self):
        path = self._path("strict_preview.docx")
        output_dir = self._path("strict_preview")
        cases = [
            (["doc-preview", path, output_dir, "--bogus", "--json"], "--bogus"),
            (["doc-preview", path, output_dir, "--format", "--json"], "--format"),
        ]

        for argv, expected_text in cases:
            with self.subTest(argv=argv):
                with mock.patch.object(cli_module.doc_server, "preview_document") as preview:
                    payload, exit_code = self._run_cli(argv)

                self.assertNotEqual(exit_code, 0)
                self.assertFalse(payload["success"])
                self.assertEqual(payload["error_code"], "usage_error")
                self.assertIn(expected_text, payload["error"])
                preview.assert_not_called()

    def test_cli_doc_preview_dispatches_via_doc_handler(self):
        path = self._path("dispatch_preview.docx")
        output_dir = self._path("preview")
        stdout = io.StringIO()

        with mock.patch.object(
            cli_module.doc_server,
            "preview_document",
            return_value={"success": True, "pages_rendered": 2},
        ) as preview_document:
            with mock.patch(
                "sys.argv",
                [
                    "cli-anything-onlyoffice",
                    "doc-preview",
                    path,
                    output_dir,
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
        preview_document.assert_called_once_with(
            path,
            output_dir,
            pages="0-1",
            dpi=200,
            fmt="jpg",
        )

    def test_cli_doc_render_map_dispatches_via_doc_handler(self):
        path = self._path("dispatch_render_map.docx")
        stdout = io.StringIO()

        with mock.patch.object(
            cli_module.doc_server,
            "doc_render_map",
            return_value={"success": True, "mapped_entries": 12},
        ) as doc_render_map:
            with mock.patch(
                "sys.argv",
                [
                    "cli-anything-onlyoffice",
                    "doc-render-map",
                    path,
                    "--json",
                ],
            ):
                with redirect_stdout(stdout):
                    cli_module.main()

        payload = json.loads(stdout.getvalue())
        self.assertTrue(payload["success"])
        doc_render_map.assert_called_once_with(path)

    def test_cli_doc_render_audit_dispatches_profile(self):
        path = self._path("dispatch_render_audit.docx")
        pdf_path = self._path("dispatch_render_audit.pdf")
        stdout = io.StringIO()

        with mock.patch.object(
            cli_module.doc_server,
            "rendered_layout_audit",
            return_value={"success": True, "profile": "generic"},
        ) as rendered_layout_audit:
            with mock.patch(
                "sys.argv",
                [
                    "cli-anything-onlyoffice",
                    "doc-render-audit",
                    path,
                    "--pdf",
                    pdf_path,
                    "--tolerance-points",
                    "5.5",
                    "--profile",
                    "generic",
                    "--json",
                ],
            ):
                with redirect_stdout(stdout):
                    cli_module.main()

        payload = json.loads(stdout.getvalue())
        self.assertTrue(payload["success"])
        rendered_layout_audit.assert_called_once_with(
            path,
            pdf_path=pdf_path,
            tolerance_points=5.5,
            profile="generic",
        )

    def test_cli_doc_normalize_format_dispatches_options(self):
        path = self._path("dispatch_normalize.docx")
        output_path = self._path("dispatch_normalize_out.docx")

        with mock.patch.object(
            cli_module.doc_server,
            "normalize_document_format",
            return_value={"success": True, "output_file": output_path},
        ) as normalize_document_format:
            payload, exit_code = self._run_cli(
                [
                    "doc-normalize-format",
                    path,
                    output_path,
                    "--font",
                    "Times New Roman",
                    "--body-size",
                    "11",
                    "--title-size",
                    "12",
                    "--line-spacing",
                    "double",
                    "--paragraph-after",
                    "12",
                    "--clear-theme-fonts",
                    "--remove-style-borders",
                    "--reference-hanging",
                    "0.5",
                    "--json",
                ]
            )

        self.assertEqual(exit_code, 0)
        self.assertTrue(payload["success"])
        normalize_document_format.assert_called_once_with(
            path,
            output_path=output_path,
            font_name="Times New Roman",
            body_font_size=11.0,
            title_font_size=12.0,
            line_spacing="double",
            paragraph_after=12.0,
            clear_theme_fonts=True,
            include_header_footer=True,
            remove_style_borders=True,
            reference_hanging_inches=0.5,
        )

    def test_cli_doc_font_audit_dispatches_rendered_options(self):
        path = self._path("dispatch_font_audit.docx")
        pdf_path = self._path("dispatch_font_audit.pdf")

        with mock.patch.object(
            cli_module.doc_server,
            "audit_document_fonts",
            return_value={"success": True, "overall_status": "pass"},
        ) as audit_document_fonts:
            payload, exit_code = self._run_cli(
                [
                    "doc-font-audit",
                    path,
                    "--expected-font",
                    "Times New Roman",
                    "--expected-font-size",
                    "12",
                    "--rendered",
                    "--pdf",
                    pdf_path,
                    "--json",
                ]
            )

        self.assertEqual(exit_code, 0)
        self.assertTrue(payload["success"])
        audit_document_fonts.assert_called_once_with(
            path,
            expected_font_name="Times New Roman",
            expected_font_size=12.0,
            rendered=True,
            pdf_path=pdf_path,
        )

    def test_cli_doc_submission_pack_dispatches_options(self):
        path = self._path("dispatch_pack.docx")
        output_dir = self._path("dispatch_pack_out")

        with mock.patch.object(
            cli_module.doc_server,
            "submission_pack",
            return_value={"success": True, "submission_ready": True},
        ) as submission_pack:
            payload, exit_code = self._run_cli(
                [
                    "doc-submission-pack",
                    path,
                    output_dir,
                    "--basename",
                    "final",
                    "--expected-page-size",
                    "A4",
                    "--expected-font",
                    "Times New Roman",
                    "--expected-font-size",
                    "12",
                    "--profile",
                    "apa-references",
                    "--skip-pdf-sanitize",
                    "--json",
                ]
            )

        self.assertEqual(exit_code, 0)
        self.assertTrue(payload["success"])
        submission_pack.assert_called_once_with(
            path,
            output_dir,
            basename="final",
            expected_page_size="A4",
            expected_font_name="Times New Roman",
            expected_font_size=12.0,
            render_profile="apa-references",
            sanitize_docx=True,
            sanitize_pdf=False,
            rendered_layout=True,
        )

    def test_cli_doc_create_respects_docx_availability_guard(self):
        path = self._path("unavailable.docx")
        stdout = io.StringIO()

        with mock.patch.object(cli_module, "DOCX_AVAILABLE", False):
            with mock.patch.object(cli_module.doc_server, "create_document") as create_document:
                with mock.patch(
                    "sys.argv",
                    [
                        "cli-anything-onlyoffice",
                        "doc-create",
                        path,
                        "Title",
                        "Body",
                        "--json",
                    ],
                ):
                    with redirect_stdout(stdout):
                        cli_module.main()

        payload = json.loads(stdout.getvalue())
        self.assertFalse(payload["success"])
        self.assertEqual(payload["error"], "python-docx not installed")
        create_document.assert_not_called()
