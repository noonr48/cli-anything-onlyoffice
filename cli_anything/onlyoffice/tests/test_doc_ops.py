import os
import json
import tempfile
import unittest
from unittest import mock

from docx import Document
from docx.shared import Pt

from cli_anything.onlyoffice.utils.docserver import get_client


class OnlyOfficeDocOpsTests(unittest.TestCase):
    def setUp(self):
        self.client = get_client()
        self.ops = self.client._doc_ops
        self.tmpdir = tempfile.TemporaryDirectory(prefix="oo-doc-ops-test-")
        self.base = self.tmpdir.name

    def tearDown(self):
        self.tmpdir.cleanup()

    def _path(self, name: str) -> str:
        return os.path.join(self.base, name)

    def test_doc_ops_sanitize_removes_comments_and_metadata(self):
        path = self._path("sanitize_comments.docx")
        self.client.create_document(path, "Title", "Body")
        self.client.set_metadata(
            path,
            author="Original Author",
            title="Assignment Draft",
            subject="Research Methods",
            keywords="draft,metadata",
            comments="Internal note",
            category="Assignments",
        )
        comment_result = self.client.add_comment(path, "Review note", 0)
        self.assertTrue(comment_result["success"])

        before = self.ops.inspect_hidden_data(path)
        self.assertTrue(before["success"])
        self.assertTrue(before["comments_part_present"])
        self.assertGreaterEqual(before["comments_count"], 1)
        self.assertEqual(before["core_properties"]["author"], "Original Author")

        result = self.ops.sanitize_document(
            path,
            remove_comments=True,
            clear_metadata=True,
            author="benbi",
        )

        self.assertTrue(result["success"])
        after = result["after"]
        self.assertFalse(after["comments_part_present"])
        self.assertEqual(after["comments_count"], 0)
        self.assertEqual(after["comment_reference_count"], 0)
        self.assertEqual(after["core_properties"]["author"], "benbi")
        self.assertEqual(after["core_properties"]["title"], "")
        self.assertTrue(after["remove_personal_information"])

    def test_doc_ops_preflight_flags_mixed_fonts_and_metadata(self):
        path = self._path("preflight.docx")
        doc = Document()
        first = doc.add_paragraph()
        first.add_run("Times paragraph")
        second = doc.add_paragraph()
        run = second.add_run("Calibri paragraph")
        run.font.name = "Calibri"
        run.font.size = Pt(12)
        doc.save(path)
        self.client.set_page_layout(path, page_size="A4")
        self.client.set_metadata(path, author="benbi", title="Draft")

        result = self.ops.document_preflight(
            path,
            expected_page_size="A4",
            expected_font_name="Times New Roman",
            expected_font_size=11,
        )

        self.assertTrue(result["success"])
        self.assertEqual(result["overall_status"], "warn")
        check_status = {check["name"]: check["status"] for check in result["checks"]}
        self.assertEqual(check_status["page_size"], "pass")
        self.assertEqual(check_status["metadata"], "warn")
        self.assertEqual(check_status["font_names"], "warn")
        self.assertEqual(check_status["font_sizes"], "warn")
        self.assertEqual(result["font_audit"]["unexpected_font_run_count"], 2)
        self.assertEqual(result["font_audit"]["unexpected_size_run_count"], 2)

    def test_doc_ops_preview_uses_host_doc_pdf_pipeline(self):
        path = self._path("preview.docx")
        self.client.create_document(path, "Title", "Intro")
        output_dir = self._path("previews")

        captured = {}

        def fake_doc_to_pdf(file_path, output_path=None):
            self.assertEqual(file_path, path)
            self.assertIsNotNone(output_path)
            with open(output_path, "wb") as handle:
                handle.write(b"%PDF-1.4 fake")
            captured["pdf_path"] = output_path
            return {
                "success": True,
                "input_file": file_path,
                "output_file": output_path,
                "pages": 2,
            }

        def fake_pdf_page_to_image(file_path, render_dir, pages=None, dpi=150, fmt="png"):
            self.assertEqual(file_path, captured["pdf_path"])
            self.assertTrue(os.path.exists(file_path))
            self.assertEqual(render_dir, output_dir)
            self.assertEqual(pages, "0-1")
            self.assertEqual(dpi, 200)
            self.assertEqual(fmt, "jpg")
            return {
                "success": True,
                "total_pages": 2,
                "pages_rendered": 2,
                "images": [{"page": 0, "file": os.path.join(render_dir, "page_000.jpg")}],
            }

        with mock.patch.object(self.client, "doc_to_pdf", side_effect=fake_doc_to_pdf):
            with mock.patch.object(
                self.client, "pdf_page_to_image", side_effect=fake_pdf_page_to_image
            ):
                result = self.ops.preview_document(
                    path, output_dir, pages="0-1", dpi=200, fmt="jpg"
                )

        self.assertTrue(result["success"])
        self.assertEqual(result["total_pages"], 2)
        self.assertEqual(result["pages_rendered"], 2)
        self.assertEqual(result["format"], "jpg")
        self.assertFalse(os.path.exists(captured["pdf_path"]))

    def test_doc_ops_render_map_uses_host_pdf_block_pipeline(self):
        path = self._path("render_map.docx")
        doc = Document()
        doc.add_paragraph("Executive summary")
        table = doc.add_table(rows=1, cols=1)
        table.cell(0, 0).text = "42"
        doc.save(path)

        def fake_doc_to_pdf(file_path, output_path=None):
            self.assertTrue(file_path.endswith(".docx"))
            self.assertIsNotNone(output_path)
            with open(output_path, "wb") as handle:
                handle.write(b"%PDF-1.4 fake")
            return {"success": True, "output_file": output_path, "pages": 1}

        def fake_pdf_read_blocks(file_path, pages=None, include_spans=True, include_images=False, include_empty=False):
            self.assertTrue(file_path.endswith(".pdf"))
            self.assertTrue(os.path.exists(file_path))
            self.assertTrue(include_spans)
            return {
                "success": True,
                "pages_scanned": 1,
                "total_pages": 1,
                "pages": [
                    {
                        "page_index": 0,
                        "page_number": 1,
                        "blocks": [
                            {
                                "block_id": "page_0_block_0",
                                "type": "text",
                                "bbox": {"left": 10, "top": 10, "right": 200, "bottom": 40, "width": 190, "height": 30},
                                "lines": [
                                    {
                                        "line_id": "page_0_block_0_line_0",
                                        "bbox": {"left": 10, "top": 10, "right": 200, "bottom": 25, "width": 190, "height": 15},
                                        "spans": [
                                            {
                                                "span_id": "page_0_block_0_line_0_span_0",
                                                "text": "SLOANE_P_0001 Executive summary",
                                                "bbox": {"left": 10, "top": 10, "right": 200, "bottom": 25, "width": 190, "height": 15},
                                            }
                                        ],
                                    }
                                ],
                            },
                            {
                                "block_id": "page_0_block_1",
                                "type": "text",
                                "bbox": {"left": 10, "top": 50, "right": 200, "bottom": 80, "width": 190, "height": 30},
                                "lines": [
                                    {
                                        "line_id": "page_0_block_1_line_0",
                                        "bbox": {"left": 10, "top": 50, "right": 200, "bottom": 65, "width": 190, "height": 15},
                                        "spans": [
                                            {
                                                "span_id": "page_0_block_1_line_0_span_0",
                                                "text": "SLOANE_T1R1C1 42",
                                                "bbox": {"left": 10, "top": 50, "right": 200, "bottom": 65, "width": 190, "height": 15},
                                            }
                                        ],
                                    }
                                ],
                            },
                        ],
                    }
                ],
            }

        with mock.patch.object(self.client, "doc_to_pdf", side_effect=fake_doc_to_pdf):
            with mock.patch.object(self.client, "pdf_read_blocks", side_effect=fake_pdf_read_blocks):
                result = self.ops.doc_render_map(path)

        self.assertTrue(result["success"])
        self.assertEqual(result["mapped_paragraph_count"], 1)
        self.assertEqual(result["mapped_table_cell_count"], 1)
        self.assertEqual(result["paragraphs"][0]["page_number"], 1)
        self.assertEqual(result["table_cells"][0]["cell_ref"], "T1R1C1")

    def test_doc_ops_basic_crud_replace_format_and_counts(self):
        path = self._path("basic.docx")

        created = self.ops.create_document(path, "Report Title", "First paragraph")
        self.assertTrue(created["success"])

        appended = self.ops.append_to_document(path, "Second paragraph\nThird paragraph")
        self.assertTrue(appended["success"])

        replaced = self.ops.search_replace_document(path, "Second", "Updated Second")
        self.assertTrue(replaced["success"])
        self.assertEqual(replaced["replacements"], 1)

        formatted = self.ops.format_paragraph(
            path,
            1,
            bold=True,
            italic=True,
            font_name="Times New Roman",
            font_size=12,
            alignment="center",
        )
        self.assertTrue(formatted["success"])

        read_back = self.ops.read_document(path)
        self.assertTrue(read_back["success"])
        self.assertEqual(
            read_back["paragraphs"],
            [
                "Report Title",
                "First paragraph",
                "Updated Second paragraph",
                "Third paragraph",
            ],
        )

        counts = self.ops.word_count(path)
        self.assertTrue(counts["success"])
        self.assertGreaterEqual(counts["words"], 8)
        self.assertEqual(counts["paragraphs"], 4)

        doc = Document(path)
        para = doc.paragraphs[1]
        run = para.runs[0]
        self.assertTrue(run.bold)
        self.assertTrue(run.italic)
        self.assertEqual(run.font.name, "Times New Roman")
        self.assertEqual(round(run.font.size.pt), 12)
        self.assertIsNotNone(para.alignment)

    def test_doc_ops_table_search_insert_delete_and_list_helpers(self):
        path = self._path("table_ops.docx")
        self.ops.create_document(path, "Table Report", "Intro")

        table_result = self.ops.add_table(path, "Name,Value", "alpha,1;beta,2")
        self.assertTrue(table_result["success"])
        self.assertEqual(table_result["rows"], 2)

        tables = self.ops.read_tables(path)
        self.assertTrue(tables["success"])
        self.assertEqual(tables["table_count"], 1)
        self.assertEqual(tables["tables"][0]["headers"], ["Name", "Value"])
        self.assertEqual(tables["tables"][0]["rows"][1], ["beta", "2"])

        search = self.ops.search_document(path, "beta")
        self.assertTrue(search["success"])
        self.assertEqual(search["matches"], 0)
        self.assertEqual(search["table_matches"], 1)

        inserted = self.ops.insert_paragraph(path, "Inserted section", 1, "Heading 1")
        self.assertTrue(inserted["success"])
        deleted = self.ops.delete_paragraph(path, 1)
        self.assertTrue(deleted["success"])
        self.assertEqual(deleted["deleted_text"], "Inserted section")

        listed = self.ops.add_list(path, ["Item one", "Item two"], "bullet")
        self.assertTrue(listed["success"])
        self.assertEqual(listed["items_added"], 2)

        page_break = self.ops.add_page_break(path)
        self.assertTrue(page_break["success"])

        styles = self.ops.list_styles(path)
        self.assertTrue(styles["success"])
        self.assertGreater(styles["paragraph_style_count"], 0)

    def test_doc_ops_metadata_layout_and_references_roundtrip(self):
        path = self._path("refs_layout.docx")
        self.ops.create_document(path, "Findings", "Body text")

        metadata = self.ops.set_metadata(
            path,
            author="benbi",
            title="Assignment",
            subject="Methods",
            keywords="health,readiness",
            comments="internal",
            category="Reports",
        )
        self.assertTrue(metadata["success"])

        metadata_read = self.ops.get_metadata(path)
        self.assertTrue(metadata_read["success"])
        self.assertEqual(metadata_read["author"], "benbi")
        self.assertEqual(metadata_read["title"], "Assignment")

        layout = self.ops.set_page_layout(
            path,
            orientation="landscape",
            page_size="A4",
            header_text="Running header",
            page_numbers=True,
        )
        self.assertTrue(layout["success"])
        self.assertEqual(layout["page_size"], "A4")
        self.assertEqual(layout["orientation"], "landscape")

        formatting = self.ops.get_formatting_info(path)
        self.assertTrue(formatting["success"])
        self.assertEqual(formatting["sections"][0]["orientation"], "landscape")

        ref = {
            "author": "Smith, J.",
            "year": "2024",
            "title": "Work readiness in context",
            "source": "Journal of Health Education",
            "volume": "15",
            "issue": "1",
            "pages": "10-20",
            "doi": "10.1000/example",
            "type": "journal",
        }
        add_ref = self.ops.add_reference(path, json.dumps(ref))
        self.assertTrue(add_ref["success"])
        self.assertEqual(add_ref["action"], "added")

        duplicate_ref = self.ops.add_reference(path, json.dumps(ref))
        self.assertTrue(duplicate_ref["success"])
        self.assertEqual(duplicate_ref["action"], "duplicate_skipped")

        built = self.ops.build_references(path)
        self.assertTrue(built["success"])
        self.assertEqual(built["references_added"], 1)

        read_back = self.ops.read_document(path)
        self.assertTrue(read_back["success"])
        self.assertIn("References", read_back["paragraphs"])
        self.assertIn("Smith, J. (2024). Work readiness in context.", read_back["full_text"])
