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

    def _build_text_pdf(self, pdf_path: str, texts) -> None:
        doc = fitz.open()
        for text in texts:
            page = doc.new_page(width=595.0, height=842.0)
            page.insert_text((72, 72), text)
        doc.save(pdf_path)
        doc.close()

    def _pdf_text(self, pdf_path: str) -> str:
        doc = fitz.open(pdf_path)
        try:
            return "\n".join(page.get_text() for page in doc)
        finally:
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
        full_clean_path = self._path("sanitize-full-clean.pdf")
        doc = fitz.open()
        page = doc.new_page(width=595.0, height=842.0)
        page.insert_text((72, 72), "Sanitize Me")
        page.add_highlight_annot(fitz.Rect(70, 60, 150, 84))
        doc.embfile_add("secret.txt", b"secret attachment")
        widget = fitz.Widget()
        widget.field_name = "StudentName"
        widget.field_type = fitz.PDF_WIDGET_TYPE_TEXT
        widget.rect = fitz.Rect(72, 120, 220, 145)
        widget.field_value = "Jacky"
        page.add_widget(widget)
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
        self.assertEqual(inspect_result["embedded_files_count"], 1)
        self.assertTrue(inspect_result["has_forms"])
        self.assertTrue(inspect_result["has_xml_metadata"])

        self.assertTrue(sanitize_result["success"])
        self.assertEqual(sanitize_result["after"]["author"], "benbi")
        self.assertFalse(sanitize_result["has_xml_metadata"])
        self.assertEqual(sanitize_result["sanitization_scope"]["annotations"], "reported_not_removed")
        self.assertGreaterEqual(sanitize_result["after_hidden_data"]["annotations_count"], 1)
        self.assertTrue(any("annotations" in warning for warning in sanitize_result["warnings"]))
        self.assertTrue(os.path.exists(clean_path))

        full_sanitize = self.client.pdf_sanitize(
            pdf_path,
            output_path=full_clean_path,
            clear_metadata=True,
            remove_xml_metadata=True,
            remove_annotations=True,
            remove_embedded_files=True,
            flatten_forms=True,
        )

        self.assertTrue(full_sanitize["success"], full_sanitize)
        self.assertEqual(full_sanitize["sanitization_scope"]["annotations"], "removed")
        self.assertEqual(full_sanitize["sanitization_scope"]["embedded_files"], "removed")
        self.assertEqual(full_sanitize["sanitization_scope"]["forms"], "flattened")
        self.assertGreaterEqual(full_sanitize["actions"]["annotations_removed"], 1)
        self.assertEqual(full_sanitize["actions"]["embedded_files_removed"], 1)
        self.assertEqual(full_sanitize["after_hidden_data"]["annotations_count"], 0)
        self.assertEqual(full_sanitize["after_hidden_data"]["embedded_files_count"], 0)
        self.assertFalse(full_sanitize["after_hidden_data"]["has_forms"])

    def test_pdf_compact_merge_split_and_reorder(self):
        a_path = self._path("a.pdf")
        b_path = self._path("b.pdf")
        compact_path = self._path("a-compact.pdf")
        merge_path = self._path("merged.pdf")
        reorder_path = self._path("reordered.pdf")
        split_dir = self._path("split")
        self._build_text_pdf(a_path, ["Alpha page", "Beta page"])
        self._build_text_pdf(b_path, ["Gamma page"])

        compact = self.client.pdf_compact(a_path, compact_path)
        merge = self.client.pdf_merge([a_path, b_path], merge_path)
        split = self.client.pdf_split(merge_path, split_dir, pages="1-2", prefix="part")
        reorder = self.client.pdf_reorder(merge_path, "2,0,2", output_path=reorder_path)

        self.assertTrue(compact["success"], compact)
        self.assertTrue(compact["compression_requested"])
        self.assertFalse(compact["default_applied"])
        self.assertTrue(os.path.exists(compact_path))

        self.assertTrue(merge["success"], merge)
        self.assertEqual(merge["pages"], 3)
        self.assertEqual([entry["input_page"] for entry in merge["page_map"]], [0, 1, 0])
        self.assertIn("Alpha page", self._pdf_text(merge_path))
        self.assertIn("Gamma page", self._pdf_text(merge_path))

        self.assertTrue(split["success"], split)
        self.assertEqual(split["pages_selected"], 2)
        self.assertEqual(len(split["outputs"]), 2)
        self.assertTrue(all(os.path.exists(entry["output_file"]) for entry in split["outputs"]))

        self.assertTrue(reorder["success"], reorder)
        self.assertEqual(reorder["page_order"], [2, 0, 2])
        reordered_text = self._pdf_text(reorder_path)
        self.assertLess(reordered_text.find("Gamma page"), reordered_text.find("Alpha page"))
        self.assertEqual(reordered_text.count("Gamma page"), 2)

    def test_pdf_add_text_add_image_and_redact(self):
        pdf_path = self._path("edit.pdf")
        stamped_path = self._path("stamped.pdf")
        image_path = self._path("stamp.png")
        image_out = self._path("image-stamped.pdf")
        redacted_path = self._path("redacted.pdf")
        self._build_text_pdf(pdf_path, ["Visible text SECRET token"])
        Image.new("RGB", (32, 32), color="red").save(image_path)

        added_text = self.client.pdf_add_text(
            pdf_path,
            0,
            "STAMPED",
            output_path=stamped_path,
            x=72,
            y=120,
            width=200,
            height=50,
            font_size=14,
        )
        added_image = self.client.pdf_add_image(
            stamped_path,
            0,
            image_path,
            output_path=image_out,
            x=72,
            y=180,
            width=64,
            height=64,
        )
        dry_run = self.client.pdf_redact(image_out, text="SECRET", dry_run=True)
        redacted = self.client.pdf_redact(
            image_out,
            output_path=redacted_path,
            text="SECRET",
        )

        self.assertTrue(added_text["success"], added_text)
        self.assertIn("STAMPED", self._pdf_text(stamped_path))

        self.assertTrue(added_image["success"], added_image)
        img_doc = fitz.open(image_out)
        try:
            self.assertGreaterEqual(len(img_doc[0].get_images(full=True)), 1)
        finally:
            img_doc.close()

        self.assertTrue(dry_run["success"], dry_run)
        self.assertEqual(dry_run["match_count"], 1)

        self.assertTrue(redacted["success"], redacted)
        self.assertEqual(redacted["redactions_applied"], 1)
        redacted_text = self._pdf_text(redacted_path)
        self.assertIn("Visible text", redacted_text)
        self.assertIn("token", redacted_text)
        self.assertNotIn("SECRET", redacted_text)
        self.assertEqual(redacted["verification"]["match_count"], 0)

    def test_pdf_redact_text_exact_match_preserves_surrounding_text_and_case(self):
        pdf_path = self._path("redact-exact.pdf")
        redacted_path = self._path("redact-exact-out.pdf")
        self._build_text_pdf(pdf_path, ["Alpha SECRET beta secret gamma"])

        dry_upper = self.client.pdf_redact(
            pdf_path,
            text="SECRET",
            case_sensitive=True,
            dry_run=True,
        )
        dry_lower = self.client.pdf_redact(
            pdf_path,
            text="secret",
            case_sensitive=True,
            dry_run=True,
        )
        dry_insensitive = self.client.pdf_redact(
            pdf_path,
            text="secret",
            dry_run=True,
        )
        redacted = self.client.pdf_redact(
            pdf_path,
            output_path=redacted_path,
            text="secret",
            case_sensitive=True,
        )

        self.assertTrue(dry_upper["success"], dry_upper)
        self.assertEqual(dry_upper["match_count"], 1)
        self.assertEqual(dry_upper["matches"][0]["text_preview"], "SECRET")

        self.assertTrue(dry_lower["success"], dry_lower)
        self.assertEqual(dry_lower["match_count"], 1)
        self.assertEqual(dry_lower["matches"][0]["text_preview"], "secret")

        self.assertTrue(dry_insensitive["success"], dry_insensitive)
        self.assertEqual(dry_insensitive["match_count"], 2)

        self.assertTrue(redacted["success"], redacted)
        self.assertEqual(redacted["redactions_applied"], 1)
        redacted_text = self._pdf_text(redacted_path)
        self.assertIn("Alpha", redacted_text)
        self.assertIn("SECRET", redacted_text)
        self.assertIn("beta", redacted_text)
        self.assertIn("gamma", redacted_text)
        self.assertNotIn("secret", redacted_text)
        self.assertEqual(redacted["verification"]["match_count"], 0)

    def test_pdf_map_page_and_redact_block_support_visual_selection(self):
        pdf_path = self._path("map-blocks.pdf")
        map_path = self._path("map-blocks.png")
        redacted_path = self._path("map-blocks-redacted.pdf")
        doc = fitz.open()
        page = doc.new_page(width=595.0, height=842.0)
        page.insert_text((72, 72), "KEEP THIS")
        page.insert_text((72, 140), "REMOVE THIS")
        doc.save(pdf_path)
        doc.close()

        page_map = self.client.pdf_map_page(pdf_path, 0, map_path, dpi=100)
        remove_block_id = next(
            block["block_id"]
            for block in page_map["blocks"]
            if "REMOVE THIS" in block.get("text", "")
        )
        dry_run = self.client.pdf_redact_block(pdf_path, remove_block_id, dry_run=True)
        invalid_fill = self.client.pdf_redact_block(
            pdf_path,
            remove_block_id,
            fill="not-a-color",
            dry_run=True,
        )
        redacted = self.client.pdf_redact_block(
            pdf_path,
            remove_block_id,
            output_path=redacted_path,
        )

        self.assertTrue(page_map["success"], page_map)
        self.assertTrue(os.path.exists(map_path))
        self.assertGreaterEqual(page_map["blocks_mapped"], 2)
        self.assertEqual(page_map["format"], "png")

        self.assertTrue(dry_run["success"], dry_run)
        self.assertEqual(dry_run["block_id"], remove_block_id)
        self.assertEqual(dry_run["block"]["text"], "REMOVE THIS")

        self.assertFalse(invalid_fill["success"], invalid_fill)
        self.assertEqual(invalid_fill["error_code"], "usage_error")

        self.assertTrue(redacted["success"], redacted)
        self.assertEqual(redacted["selector"], "block")
        redacted_text = self._pdf_text(redacted_path)
        self.assertIn("KEEP THIS", redacted_text)
        self.assertNotIn("REMOVE THIS", redacted_text)

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

    def test_cli_pdf_mutation_commands_dispatch_via_pdf_handler(self):
        pdf_path = self._path("dispatch_edit.pdf")
        out_path = self._path("dispatch_out.pdf")
        image_path = self._path("dispatch.png")
        second_path = self._path("dispatch_second.pdf")
        out_dir = self._path("split_out")
        with open(pdf_path, "wb") as handle:
            handle.write(b"%PDF-1.4")
        with open(second_path, "wb") as handle:
            handle.write(b"%PDF-1.4")
        Image.new("RGB", (8, 8), color="blue").save(image_path)

        cases = [
            (
                "pdf-compact",
                "pdf_compact",
                [
                    "pdf-compact",
                    pdf_path,
                    out_path,
                    "--garbage",
                    "3",
                    "--linearize",
                    "--json",
                ],
                {
                    "output_path": out_path,
                    "garbage": 3,
                    "deflate": True,
                    "clean": True,
                    "linearize": True,
                },
            ),
            (
                "pdf-merge",
                "pdf_merge",
                [
                    "pdf-merge",
                    pdf_path,
                    second_path,
                    "--output",
                    out_path,
                    "--json",
                ],
                {},
            ),
            (
                "pdf-split",
                "pdf_split",
                [
                    "pdf-split",
                    pdf_path,
                    out_dir,
                    "--pages",
                    "0",
                    "--prefix",
                    "part",
                    "--json",
                ],
                {"pages": "0", "prefix": "part"},
            ),
            (
                "pdf-reorder",
                "pdf_reorder",
                ["pdf-reorder", pdf_path, "0,0", out_path, "--json"],
                {"output_path": out_path},
            ),
            (
                "pdf-add-text",
                "pdf_add_text",
                [
                    "pdf-add-text",
                    pdf_path,
                    "0",
                    "Note",
                    "--output",
                    out_path,
                    "--x",
                    "10",
                    "--y",
                    "20",
                    "--width",
                    "100",
                    "--height",
                    "30",
                    "--font-size",
                    "12",
                    "--font",
                    "helv",
                    "--color",
                    "FF0000",
                    "--rotation",
                    "0",
                    "--json",
                ],
                {
                    "output_path": out_path,
                    "x": 10.0,
                    "y": 20.0,
                    "width": 100.0,
                    "height": 30.0,
                    "font_size": 12.0,
                    "font_name": "helv",
                    "color": "FF0000",
                    "rotation": 0,
                },
            ),
            (
                "pdf-add-image",
                "pdf_add_image",
                [
                    "pdf-add-image",
                    pdf_path,
                    "0",
                    image_path,
                    "--output",
                    out_path,
                    "--x",
                    "10",
                    "--y",
                    "20",
                    "--width",
                    "40",
                    "--height",
                    "50",
                    "--no-keep-proportion",
                    "--json",
                ],
                {
                    "output_path": out_path,
                    "x": 10.0,
                    "y": 20.0,
                    "width": 40.0,
                    "height": 50.0,
                    "keep_proportion": False,
                },
            ),
            (
                "pdf-redact",
                "pdf_redact",
                [
                    "pdf-redact",
                    pdf_path,
                    out_path,
                    "--text",
                    "SECRET",
                    "--pages",
                    "0",
                    "--case-sensitive",
                    "--fill",
                    "000000",
                    "--dry-run",
                    "--json",
                ],
                {
                    "output_path": out_path,
                    "text": "SECRET",
                    "rects": None,
                    "pages": "0",
                    "case_sensitive": True,
                    "fill": "000000",
                    "dry_run": True,
                },
            ),
            (
                "pdf-map-page",
                "pdf_map_page",
                [
                    "pdf-map-page",
                    pdf_path,
                    "0",
                    image_path,
                    "--dpi",
                    "120",
                    "--format",
                    "jpg",
                    "--no-labels",
                    "--no-images",
                    "--json",
                ],
                {
                    "dpi": 120,
                    "fmt": "jpg",
                    "labels": False,
                    "include_images": False,
                },
            ),
            (
                "pdf-redact-block",
                "pdf_redact_block",
                [
                    "pdf-redact-block",
                    pdf_path,
                    "page_0_block_1",
                    out_path,
                    "--fill",
                    "FFFFFF",
                    "--dry-run",
                    "--json",
                ],
                {
                    "output_path": out_path,
                    "fill": "FFFFFF",
                    "dry_run": True,
                },
            ),
        ]

        for command, method_name, argv, expected_kwargs in cases:
            with self.subTest(command=command):
                stdout = io.StringIO()
                with mock.patch.object(
                    cli_module.doc_server,
                    method_name,
                    return_value={"success": True, "command": command},
                ) as operation:
                    with mock.patch("sys.argv", ["cli-anything-onlyoffice", *argv]):
                        with redirect_stdout(stdout):
                            cli_module.main()

                payload = json.loads(stdout.getvalue())
                self.assertTrue(payload["success"])
                if command == "pdf-merge":
                    operation.assert_called_once_with([pdf_path, second_path], out_path)
                elif command == "pdf-split":
                    operation.assert_called_once_with(pdf_path, out_dir, **expected_kwargs)
                elif command == "pdf-reorder":
                    operation.assert_called_once_with(pdf_path, "0,0", **expected_kwargs)
                elif command == "pdf-add-text":
                    operation.assert_called_once_with(pdf_path, 0, "Note", **expected_kwargs)
                elif command == "pdf-add-image":
                    operation.assert_called_once_with(pdf_path, 0, image_path, **expected_kwargs)
                elif command == "pdf-redact":
                    operation.assert_called_once_with(pdf_path, **expected_kwargs)
                elif command == "pdf-map-page":
                    operation.assert_called_once_with(pdf_path, 0, image_path, **expected_kwargs)
                elif command == "pdf-redact-block":
                    operation.assert_called_once_with(
                        pdf_path,
                        "page_0_block_1",
                        **expected_kwargs,
                    )
                else:
                    operation.assert_called_once_with(pdf_path, **expected_kwargs)
