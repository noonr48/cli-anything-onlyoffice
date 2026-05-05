import io
import json
import os
import tempfile
import unittest
from contextlib import redirect_stdout
from unittest import mock

import fitz
from PIL import Image

from cli_anything.onlyoffice.core import cli as cli_module
from cli_anything.onlyoffice.utils.docserver import get_client
from cli_anything.onlyoffice.utils.pdf_ops import PDFOperations


class OnlyOfficePDFTests(unittest.TestCase):
    def setUp(self):
        self.client = get_client()
        self.tmpdir = tempfile.TemporaryDirectory(prefix="oo-pdf-test-")
        self.base = self.tmpdir.name

    def tearDown(self):
        self.tmpdir.cleanup()

    def _path(self, name: str) -> str:
        return os.path.join(self.base, name)

    def _build_pdf_with_image(self, pdf_path: str, image_path: str) -> None:
        Image.new("RGB", (64, 32), color="navy").save(image_path)
        doc = fitz.open()
        page = doc.new_page(width=595.0, height=842.0)
        page.insert_text((72, 72), "PDF extraction target")
        page.insert_image(fitz.Rect(72, 120, 200, 184), filename=image_path)
        doc.save(pdf_path)
        doc.close()

    def test_pdf_extract_images_and_page_to_image_create_outputs(self):
        pdf_path = self._path("images.pdf")
        image_path = self._path("source.png")
        extract_dir = self._path("extract")
        render_dir = self._path("render")
        self._build_pdf_with_image(pdf_path, image_path)

        extract_result = self.client.pdf_extract_images(pdf_path, extract_dir, fmt="png")
        render_result = self.client.pdf_page_to_image(pdf_path, render_dir, pages="0", dpi=100, fmt="png")

        self.assertTrue(extract_result["success"])
        self.assertGreaterEqual(extract_result["images_extracted"], 1)
        self.assertTrue(os.path.exists(extract_result["images"][0]["file"]))

        self.assertTrue(render_result["success"])
        self.assertEqual(render_result["pages_rendered"], 1)
        self.assertTrue(os.path.exists(render_result["images"][0]["file"]))
        self.assertEqual(render_result["preflight"]["status"], "pass")
        self.assertFalse([name for name in os.listdir(render_dir) if name.startswith(".")])

    def test_pdf_image_outputs_reject_unsafe_resource_and_format_requests(self):
        pdf_path = self._path("unsafe_render.pdf")
        image_path = self._path("source.png")
        render_dir = self._path("render_unsafe")
        self._build_pdf_with_image(pdf_path, image_path)

        dpi_result = self.client.pdf_page_to_image(
            pdf_path,
            render_dir,
            pages="0",
            dpi=5000,
            fmt="png",
        )
        bad_format_result = self.client.pdf_extract_images(
            pdf_path,
            render_dir,
            fmt="../escape",
        )

        self.assertFalse(dpi_result["success"])
        self.assertEqual(dpi_result["error_code"], "unsafe_pdf_render_request")
        self.assertEqual(dpi_result["preflight"]["status"], "fail")

        self.assertFalse(bad_format_result["success"])
        self.assertEqual(bad_format_result["error_code"], "unsupported_image_format")

    def test_pdf_page_ranges_reject_out_of_range_and_over_limit_requests(self):
        pdf_path = self._path("ranges.pdf")
        image_path = self._path("source.png")
        render_dir = self._path("render_ranges")
        self._build_pdf_with_image(pdf_path, image_path)

        out_of_range = self.client.pdf_page_to_image(
            pdf_path,
            render_dir,
            pages="2",
            dpi=100,
            fmt="png",
        )
        with mock.patch.object(PDFOperations, "MAX_RENDER_PAGES", 0):
            over_limit = self.client.pdf_page_to_image(
                pdf_path,
                render_dir,
                pages="0",
                dpi=100,
                fmt="png",
            )

        self.assertFalse(out_of_range["success"])
        self.assertEqual(out_of_range["error_code"], "invalid_page_range")
        self.assertIn("out of range", out_of_range["error"])

        self.assertFalse(over_limit["success"])
        self.assertEqual(over_limit["error_code"], "unsafe_pdf_resource_request")
        self.assertEqual(over_limit["preflight"]["status"], "fail")

        with self.assertRaisesRegex(ValueError, "safety limit"):
            PDFOperations.parse_page_range(None, 10_000_000, max_pages=1)
        with self.assertRaisesRegex(ValueError, "out of range"):
            PDFOperations.parse_page_range("0-999999999", 1, max_pages=50)

    def test_pdf_extract_images_enforces_embedded_image_resource_limits(self):
        pdf_path = self._path("bounded_extract.pdf")
        image_path = self._path("bounded_source.png")
        self._build_pdf_with_image(pdf_path, image_path)

        with mock.patch.object(PDFOperations, "MAX_EXTRACT_IMAGES", 0):
            count_limited = self.client.pdf_extract_images(
                pdf_path,
                self._path("count_limited"),
                fmt="png",
            )

        self.assertTrue(count_limited["success"])
        self.assertTrue(count_limited["truncated"])
        self.assertEqual(count_limited["images_extracted"], 0)
        self.assertEqual(count_limited["resource_limits"]["max_images"], 0)
        self.assertTrue(any("Stopped after 0 images" in warning for warning in count_limited["warnings"]))

        with mock.patch.object(PDFOperations, "MAX_EXTRACT_IMAGE_COMPRESSED_BYTES", 1):
            byte_limited = self.client.pdf_extract_images(
                pdf_path,
                self._path("byte_limited"),
                fmt="png",
            )

        self.assertTrue(byte_limited["success"])
        self.assertEqual(byte_limited["images_extracted"], 0)
        self.assertEqual(byte_limited["images_skipped"], 1)
        self.assertTrue(byte_limited["images"][0]["skipped"])
        self.assertEqual(byte_limited["images"][0]["error_code"], "image_compressed_bytes_limit_exceeded")

        with mock.patch.object(PDFOperations, "MAX_EXTRACT_IMAGE_PIXELS", 1):
            with mock.patch.object(Image.Image, "save") as image_save:
                pixel_limited = self.client.pdf_extract_images(
                    pdf_path,
                    self._path("pixel_limited"),
                    fmt="png",
                )

        self.assertTrue(pixel_limited["success"])
        self.assertEqual(pixel_limited["images_extracted"], 0)
        self.assertEqual(pixel_limited["images_skipped"], 1)
        self.assertEqual(pixel_limited["images"][0]["error_code"], "image_pixel_limit_exceeded")
        image_save.assert_not_called()

        with mock.patch.object(PDFOperations, "MAX_EXTRACT_TOTAL_COMPRESSED_BYTES", 1):
            aggregate_byte_limited = self.client.pdf_extract_images(
                pdf_path,
                self._path("aggregate_byte_limited"),
                fmt="png",
            )

        self.assertTrue(aggregate_byte_limited["success"])
        self.assertTrue(aggregate_byte_limited["truncated"])
        self.assertEqual(aggregate_byte_limited["images_extracted"], 0)
        self.assertEqual(aggregate_byte_limited["images_skipped"], 1)
        self.assertEqual(
            aggregate_byte_limited["images"][0]["error_code"],
            "aggregate_image_compressed_bytes_limit_exceeded",
        )
        self.assertEqual(
            aggregate_byte_limited["resource_limits"]["max_total_compressed_image_bytes"],
            1,
        )

        with mock.patch.object(PDFOperations, "MAX_EXTRACT_TOTAL_PIXELS", 1):
            with mock.patch.object(Image.Image, "save") as image_save:
                aggregate_pixel_limited = self.client.pdf_extract_images(
                    pdf_path,
                    self._path("aggregate_pixel_limited"),
                    fmt="png",
                )

        self.assertTrue(aggregate_pixel_limited["success"])
        self.assertTrue(aggregate_pixel_limited["truncated"])
        self.assertEqual(aggregate_pixel_limited["images_extracted"], 0)
        self.assertEqual(aggregate_pixel_limited["images_skipped"], 1)
        self.assertEqual(
            aggregate_pixel_limited["images"][0]["error_code"],
            "aggregate_image_pixel_limit_exceeded",
        )
        image_save.assert_not_called()

    def test_pdf_read_blocks_enforces_resource_limits_and_reports_truncation(self):
        pdf_path = self._path("blocks_limit.pdf")
        doc = fitz.open()
        page = doc.new_page()
        page.insert_text((72, 72), "Alpha Results")
        doc.save(pdf_path)
        doc.close()

        with mock.patch.object(PDFOperations, "MAX_READ_TEXT_CHARS", 5):
            result = self.client.pdf_read_blocks(pdf_path, pages="0")

        self.assertTrue(result["success"])
        self.assertTrue(result["truncated"])
        self.assertEqual(result["resource_limits"]["max_text_chars"], 5)
        self.assertTrue(any("text characters" in warning for warning in result["warnings"]))

        with mock.patch.object(PDFOperations, "MAX_READ_PAGES", 0):
            blocked = self.client.pdf_read_blocks(pdf_path, pages="0")

        self.assertFalse(blocked["success"])
        self.assertEqual(blocked["error_code"], "unsafe_pdf_resource_request")
        self.assertEqual(blocked["preflight"]["status"], "fail")

    def test_pdf_read_blocks_and_search_blocks_find_expected_text(self):
        pdf_path = self._path("blocks.pdf")
        doc = fitz.open()
        page = doc.new_page()
        page.insert_text((72, 72), "Alpha Results")
        page.insert_text((72, 120), "Beta Findings")
        doc.save(pdf_path)
        doc.close()

        read_result = self.client.pdf_read_blocks(pdf_path, pages="0")
        search_result = self.client.pdf_search_blocks(pdf_path, "Results", pages="0")

        self.assertTrue(read_result["success"])
        self.assertEqual(read_result["pages_scanned"], 1)
        self.assertGreaterEqual(read_result["text_block_count"], 1)
        self.assertTrue(any(block["type"] == "text" for block in read_result["pages"][0]["blocks"]))

        self.assertTrue(search_result["success"])
        self.assertGreaterEqual(search_result["match_count"], 1)
        self.assertEqual(search_result["matches"][0]["scope"], "block")

    def test_pdf_inspect_hidden_data_and_sanitize_output(self):
        pdf_path = self._path("sanitize.pdf")
        clean_path = self._path("sanitize-clean.pdf")
        doc = fitz.open()
        page = doc.new_page(width=595.0, height=842.0)
        page.insert_text((72, 72), "Sanitize Me")
        page.add_highlight_annot(fitz.Rect(70, 60, 150, 84))
        doc.set_metadata({"author": "benbi", "title": "Draft"})
        if hasattr(doc, "set_xml_metadata"):
            doc.set_xml_metadata(
                "<x:xmpmeta xmlns:x='adobe:ns:meta/'><rdf:RDF xmlns:rdf='http://www.w3.org/1999/02/22-rdf-syntax-ns#'></rdf:RDF></x:xmpmeta>"
            )
        doc.save(pdf_path)
        doc.close()

        inspect_result = self.client.inspect_pdf_hidden_data(pdf_path)
        sanitize_result = self.client.pdf_sanitize(
            pdf_path,
            output_path=clean_path,
            clear_metadata=True,
            remove_xml_metadata=True,
            author="benbi",
        )

        self.assertTrue(inspect_result["success"])
        self.assertEqual(inspect_result["page_size_labels"], ["A4"])
        self.assertGreaterEqual(inspect_result["annotations_count"], 1)
        self.assertTrue(inspect_result["has_xml_metadata"])

        self.assertTrue(sanitize_result["success"])
        self.assertEqual(sanitize_result["after"]["author"], "benbi")
        self.assertFalse(sanitize_result["has_xml_metadata"])
        self.assertEqual(sanitize_result["sanitization_scope"]["annotations"], "reported_not_removed")
        self.assertGreaterEqual(sanitize_result["after_hidden_data"]["annotations_count"], 1)
        self.assertTrue(any("annotations" in warning for warning in sanitize_result["warnings"]))
        self.assertTrue(os.path.exists(clean_path))

    def test_cli_pdf_extract_images_dispatches_via_pdf_handler(self):
        pdf_path = self._path("dispatch_extract.pdf")
        out_dir = self._path("out")
        with open(pdf_path, "wb") as handle:
            handle.write(b"%PDF-1.4")
        stdout = io.StringIO()

        with mock.patch.object(
            cli_module.doc_server,
            "pdf_extract_images",
            return_value={"success": True, "file": pdf_path, "images_extracted": 1},
        ) as pdf_extract_images:
            with mock.patch(
                "sys.argv",
                [
                    "cli-anything-onlyoffice",
                    "pdf-extract-images",
                    pdf_path,
                    out_dir,
                    "--format",
                    "jpg",
                    "--pages",
                    "0",
                    "--json",
                ],
            ):
                with redirect_stdout(stdout):
                    cli_module.main()

        payload = json.loads(stdout.getvalue())
        self.assertTrue(payload["success"])
        pdf_extract_images.assert_called_once_with(pdf_path, out_dir, fmt="jpg", pages="0")

    def test_cli_pdf_read_blocks_dispatches_via_pdf_handler(self):
        pdf_path = self._path("dispatch_blocks.pdf")
        with open(pdf_path, "wb") as handle:
            handle.write(b"%PDF-1.4")
        stdout = io.StringIO()

        with mock.patch.object(
            cli_module.doc_server,
            "pdf_read_blocks",
            return_value={"success": True, "file": pdf_path, "pages_scanned": 1, "pages": []},
        ) as pdf_read_blocks:
            with mock.patch(
                "sys.argv",
                [
                    "cli-anything-onlyoffice",
                    "pdf-read-blocks",
                    pdf_path,
                    "--pages",
                    "0-1",
                    "--no-spans",
                    "--include-empty",
                    "--json",
                ],
            ):
                with redirect_stdout(stdout):
                    cli_module.main()

        payload = json.loads(stdout.getvalue())
        self.assertTrue(payload["success"])
        pdf_read_blocks.assert_called_once_with(
            pdf_path,
            pages="0-1",
            include_spans=False,
            include_images=True,
            include_empty=True,
        )
