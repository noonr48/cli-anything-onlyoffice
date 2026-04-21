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
            author="benbi",
            title="Clean Title",
            subject="Methods",
            keywords="audit,submission",
        )

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
