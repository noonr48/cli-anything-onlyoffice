#!/usr/bin/env python3
"""DOCX submission/runtime operations for the OnlyOffice CLI."""

from __future__ import annotations

import json
import os
import re
import shutil
import tempfile
import xml.etree.ElementTree as ET
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional
from zipfile import ZipFile

try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor, Mm
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False


OOXML_NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
    "xsi": "http://www.w3.org/2001/XMLSchema-instance",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


class DocumentOperations:
    """Encapsulate DOCX submission, sanitization, and rendering workflows."""

    def __init__(self, host: Any):
        self.host = host

    def create_document(
        self, output_path: str, title: str = "", content: str = ""
    ) -> Dict[str, Any]:
        """Create a new .docx document."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            with self.host._file_lock(output_path):
                backup = self.host._snapshot_backup(output_path)
                doc = Document()
                section = doc.sections[0]
                section.page_width = Mm(210)
                section.page_height = Mm(297)
                section.top_margin = Pt(72)
                section.bottom_margin = Pt(72)
                section.left_margin = Pt(72)
                section.right_margin = Pt(72)
                normal = doc.styles["Normal"]
                normal.font.name = "Calibri"
                normal.font.size = Pt(11)
                normal.paragraph_format.line_spacing_rule = WD_LINE_SPACING.DOUBLE
                normal.paragraph_format.space_after = Pt(0)
                if title:
                    heading = doc.add_heading(title, 0)
                    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if content:
                    for paragraph in content.split("\n"):
                        if paragraph.strip():
                            doc.add_paragraph(paragraph)
                self.host._safe_save(doc, output_path)
            return {
                "success": True,
                "file": output_path,
                "title": title,
                "size": Path(output_path).stat().st_size,
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def read_document(self, file_path: str) -> Dict[str, Any]:
        """Read and extract all text from a .docx document."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            doc = Document(file_path)
            paragraphs = [para.text for para in doc.paragraphs if para.text.strip()]
            return {
                "success": True,
                "file": file_path,
                "paragraphs": paragraphs,
                "paragraph_count": len(paragraphs),
                "full_text": "\n\n".join(paragraphs),
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def append_to_document(self, file_path: str, content: str) -> Dict[str, Any]:
        """Append content to a .docx document."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            with self.host._file_lock(file_path):
                backup = self.host._snapshot_backup(file_path)
                doc = Document(file_path)
                for paragraph in content.split("\n"):
                    if paragraph.strip():
                        doc.add_paragraph(paragraph)
                self.host._safe_save(doc, file_path)
            return {
                "success": True,
                "file": file_path,
                "appended_length": len(content),
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def search_replace_document(
        self, file_path: str, search_text: str, replace_text: str
    ) -> Dict[str, Any]:
        """Find and replace text in a .docx document."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            with self.host._file_lock(file_path):
                backup = self.host._snapshot_backup(file_path)
                doc = Document(file_path)
                replacements = 0
                for para in doc.paragraphs:
                    replacements += self.host._replace_across_runs(
                        para, search_text, replace_text
                    )
                self.host._safe_save(doc, file_path)
            return {
                "success": True,
                "file": file_path,
                "replacements": replacements,
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def format_paragraph(
        self,
        file_path: str,
        paragraph_index: int,
        bold: bool = False,
        italic: bool = False,
        underline: bool = False,
        font_name: str = None,
        font_size: int = None,
        color: str = None,
        alignment: str = None,
    ) -> Dict[str, Any]:
        """Apply formatting to a specific paragraph."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            with self.host._file_lock(file_path):
                backup = self.host._snapshot_backup(file_path)
                doc = Document(file_path)
                if paragraph_index >= len(doc.paragraphs):
                    return {
                        "success": False,
                        "error": f"Paragraph index {paragraph_index} out of range",
                    }
                para = doc.paragraphs[paragraph_index]
                run = para.runs[0] if para.runs else para.add_run("")
                if bold:
                    run.bold = True
                if italic:
                    run.italic = True
                if underline:
                    run.underline = True
                if font_name:
                    run.font.name = font_name
                if font_size:
                    run.font.size = Pt(font_size)
                if color:
                    run.font.color.rgb = RGBColor(*self.host._hex_to_rgb(color))
                if alignment:
                    para.alignment = self.host._get_alignment(alignment)
                self.host._safe_save(doc, file_path)
            return {
                "success": True,
                "file": file_path,
                "paragraph": paragraph_index,
                "formatted": True,
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def highlight_text(
        self, file_path: str, search_text: str, color: str = "yellow"
    ) -> Dict[str, Any]:
        """Highlight all occurrences of text in a document."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            with self.host._file_lock(file_path):
                backup = self.host._snapshot_backup(file_path)
                doc = Document(file_path)
                highlights = 0
                for para in doc.paragraphs:
                    for run in para.runs:
                        if search_text in run.text:
                            rPr = run._element.get_or_add_rPr()
                            highlight = rPr.find(qn("w:highlight"))
                            if highlight is None:
                                highlight = OxmlElement("w:highlight")
                                rPr.append(highlight)
                            valid_colors = {
                                "yellow",
                                "green",
                                "cyan",
                                "magenta",
                                "blue",
                                "red",
                                "darkBlue",
                                "darkCyan",
                                "darkGreen",
                                "darkMagenta",
                                "darkRed",
                                "darkYellow",
                                "darkGray",
                                "lightGray",
                                "black",
                                "white",
                                "none",
                            }
                            color_val = (
                                color if color and color.lower() in valid_colors else "yellow"
                            )
                            highlight.set(qn("w:val"), color_val)
                            highlights += 1
                self.host._safe_save(doc, file_path)
            return {
                "success": True,
                "file": file_path,
                "highlights": highlights,
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def add_table(self, file_path: str, headers_csv: str, data_csv: str) -> Dict[str, Any]:
        """Add a formatted table to a Word document."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            headers = [h.strip() for h in headers_csv.split(",")]
            rows = [
                [c.strip() for c in row.split(",")]
                for row in data_csv.split(";")
                if row.strip()
            ]
            with self.host._file_lock(file_path):
                backup = self.host._snapshot_backup(file_path)
                doc = Document(file_path)
                table = doc.add_table(rows=1 + len(rows), cols=len(headers))
                table.style = "Table Grid"
                for i, h in enumerate(headers):
                    cell = table.rows[0].cells[i]
                    cell.text = h
                    for run in cell.paragraphs[0].runs:
                        run.bold = True
                for ri, row in enumerate(rows, 1):
                    for ci, val in enumerate(row):
                        if ci < len(headers):
                            table.rows[ri].cells[ci].text = val
                self.host._safe_save(doc, file_path)
            return {
                "success": True,
                "file": file_path,
                "rows": len(rows),
                "columns": len(headers),
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def add_comment(
        self,
        file_path: str,
        comment_text: str,
        paragraph_index: int = 0,
        author: str = "SLOANE Agent",
    ) -> Dict[str, Any]:
        """Add a real OOXML comment annotation to a specific paragraph."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            from lxml import etree as LET
            from docx.opc.part import Part
            from docx.opc.packuri import PackURI

            with self.host._file_lock(file_path):
                backup = self.host._snapshot_backup(file_path)
                doc = Document(file_path)
                if paragraph_index >= len(doc.paragraphs):
                    return {
                        "success": False,
                        "error": f"Paragraph index {paragraph_index} out of range (max {len(doc.paragraphs) - 1})",
                    }

                w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                comments_rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
                comments_ct = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"

                part = doc.part
                comments_el = None
                comments_part = None
                for rel in part.rels.values():
                    if "comments" in rel.reltype:
                        comments_part = rel.target_part
                        break

                if comments_part is not None:
                    comments_el = LET.fromstring(comments_part.blob)
                else:
                    comments_el = LET.Element(
                        f"{{{w_ns}}}comments", nsmap={"w": w_ns, "r": r_ns}
                    )
                    blob = LET.tostring(
                        comments_el,
                        xml_declaration=True,
                        encoding="UTF-8",
                        standalone=True,
                    )
                    comments_part = Part(
                        PackURI("/word/comments.xml"), comments_ct, blob, part.package
                    )
                    part.relate_to(comments_part, comments_rel)

                existing_ids = []
                for c in comments_el.iter(f"{{{w_ns}}}comment"):
                    cid = c.get(f"{{{w_ns}}}id")
                    if cid is not None:
                        try:
                            existing_ids.append(int(cid))
                        except ValueError:
                            pass
                comment_id = str(max(existing_ids, default=-1) + 1)

                now_str = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
                comment_elem = LET.SubElement(comments_el, f"{{{w_ns}}}comment")
                comment_elem.set(f"{{{w_ns}}}id", comment_id)
                comment_elem.set(f"{{{w_ns}}}author", author)
                comment_elem.set(f"{{{w_ns}}}date", now_str)
                cp = LET.SubElement(comment_elem, f"{{{w_ns}}}p")
                cr = LET.SubElement(cp, f"{{{w_ns}}}r")
                ct = LET.SubElement(cr, f"{{{w_ns}}}t")
                ct.text = comment_text
                comments_part._blob = LET.tostring(
                    comments_el,
                    xml_declaration=True,
                    encoding="UTF-8",
                    standalone=True,
                )

                para_el = doc.paragraphs[paragraph_index]._element
                range_start = OxmlElement("w:commentRangeStart")
                range_start.set(qn("w:id"), comment_id)
                range_end = OxmlElement("w:commentRangeEnd")
                range_end.set(qn("w:id"), comment_id)
                ref_run = OxmlElement("w:r")
                ref_rpr = OxmlElement("w:rPr")
                ref_style = OxmlElement("w:rStyle")
                ref_style.set(qn("w:val"), "CommentReference")
                ref_rpr.append(ref_style)
                ref_run.append(ref_rpr)
                comment_ref = OxmlElement("w:commentReference")
                comment_ref.set(qn("w:id"), comment_id)
                ref_run.append(comment_ref)
                para_el.insert(0, range_start)
                para_el.append(range_end)
                para_el.append(ref_run)

                self.host._safe_save(doc, file_path)
            return {
                "success": True,
                "file": file_path,
                "comment_id": int(comment_id),
                "comment_text": comment_text,
                "paragraph_index": paragraph_index,
                "author": author,
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def add_reference(self, file_path: str, ref_json: str) -> Dict[str, Any]:
        """Add a reference to the sidecar .refs.json file for later formatting."""
        try:
            ref = json.loads(ref_json)
            required = ["author", "year", "title"]
            missing = [k for k in required if not ref.get(k)]
            if missing:
                return {"success": False, "error": f"Missing required fields: {missing}"}

            refs_path = file_path + ".refs.json"
            refs = []
            if os.path.exists(refs_path):
                with open(refs_path, "r") as f:
                    refs = json.load(f)

            sig = (
                ref["author"].strip().lower(),
                str(ref["year"]).strip(),
                ref["title"].strip().lower(),
            )
            for existing in refs:
                esig = (
                    existing["author"].strip().lower(),
                    str(existing["year"]).strip(),
                    existing["title"].strip().lower(),
                )
                if esig == sig:
                    return {
                        "success": True,
                        "file": refs_path,
                        "action": "duplicate_skipped",
                        "total_refs": len(refs),
                        "note": f"Reference already exists: {ref['author']} ({ref['year']})",
                    }

            ref.setdefault("type", "journal")
            refs.append(ref)
            with open(refs_path, "w") as f:
                json.dump(refs, f, indent=2)

            return {
                "success": True,
                "file": refs_path,
                "action": "added",
                "total_refs": len(refs),
                "in_text_citation": self.host._apa_in_text(ref["author"], ref["year"]),
            }
        except json.JSONDecodeError as e:
            return {"success": False, "error": f"Invalid JSON: {e}"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def build_references(self, file_path: str) -> Dict[str, Any]:
        """Build an APA 7th References section from the sidecar .refs.json file."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        refs_path = file_path + ".refs.json"
        if not os.path.exists(refs_path):
            return {
                "success": False,
                "error": f"No references file found at {refs_path}. Use doc-add-reference first.",
            }
        try:
            with open(refs_path, "r") as f:
                refs = json.load(f)
            if not refs:
                return {"success": False, "error": "References file is empty"}

            refs.sort(key=lambda r: r.get("author", "").split(",")[0].strip().lower())

            with self.host._file_lock(file_path):
                backup = self.host._snapshot_backup(file_path)
                doc = Document(file_path)

                ref_start_idx = None
                for i, para in enumerate(doc.paragraphs):
                    if para.text.strip() == "References" and any(
                        r.bold for r in para.runs if r.bold
                    ):
                        ref_start_idx = i
                        break
                if ref_start_idx is not None:
                    body = doc.element.body
                    paras_to_remove = list(doc.paragraphs[ref_start_idx:])
                    for p in paras_to_remove:
                        body.remove(p._element)
                    if ref_start_idx > 0:
                        prev = (
                            doc.paragraphs[ref_start_idx - 1]
                            if ref_start_idx - 1 < len(doc.paragraphs)
                            else None
                        )
                        if prev and not prev.text.strip():
                            for run_el in prev._element.findall(qn("w:r")):
                                for br in run_el.findall(qn("w:br")):
                                    if br.get(qn("w:type")) == "page":
                                        body.remove(prev._element)
                                        break

                doc.add_page_break()
                heading = doc.add_paragraph()
                heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = heading.add_run("References")
                run.bold = True

                for ref in refs:
                    spans = self.host._format_apa7_reference(ref)
                    para = doc.add_paragraph()
                    for span in spans:
                        run = para.add_run(span["text"])
                        if span.get("italic"):
                            run.italic = True
                    pf = para.paragraph_format
                    pf.first_line_indent = -Inches(0.5)
                    pf.left_indent = Inches(0.5)
                    pf.space_after = Pt(0)
                    pf.space_before = Pt(0)
                    pf.line_spacing_rule = WD_LINE_SPACING.DOUBLE

                self.host._safe_save(doc, file_path)

            return {
                "success": True,
                "file": file_path,
                "references_added": len(refs),
                "refs_file": refs_path,
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def add_image(
        self,
        file_path: str,
        image_path: str,
        width_inches: float = 5.5,
        caption: str = None,
        paragraph_index: int = None,
        position: str = "after",
    ) -> Dict[str, Any]:
        """Add an image to a Word document with optional caption."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        if not os.path.exists(image_path):
            return {"success": False, "error": f"Image not found: {image_path}"}
        position = position.lower()
        if position not in {"before", "after"}:
            return {"success": False, "error": "position must be 'before' or 'after'"}
        try:
            with self.host._file_lock(file_path):
                backup = self.host._snapshot_backup(file_path)
                doc = Document(file_path)
                total = len(doc.paragraphs)
                if paragraph_index is not None and (
                    paragraph_index < 0 or paragraph_index >= total
                ):
                    return {
                        "success": False,
                        "error": f"Paragraph index {paragraph_index} out of range (0..{total - 1})",
                    }
                para = doc.add_paragraph()
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                para.paragraph_format.space_before = Pt(0)
                para.paragraph_format.space_after = Pt(0)
                run = para.add_run()
                run.add_picture(image_path, width=Inches(width_inches))
                cap_para = None
                if caption:
                    para.paragraph_format.keep_with_next = True
                    cap_para = doc.add_paragraph()
                    cap_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    cap_para.paragraph_format.space_before = Pt(0)
                    cap_para.paragraph_format.space_after = Pt(0)
                    cap_para.paragraph_format.keep_together = True
                    cap_run = cap_para.add_run(caption)
                    cap_run.italic = True
                    cap_run.font.size = Pt(10)
                if paragraph_index is not None:
                    body = doc.element.body
                    ref = doc.paragraphs[paragraph_index]._element
                    elements = [para._element]
                    if cap_para is not None:
                        elements.append(cap_para._element)
                    for element in elements:
                        body.remove(element)
                    insert_at = body.index(ref)
                    if position == "after":
                        insert_at += 1
                    for offset, element in enumerate(elements):
                        body.insert(insert_at + offset, element)
                self.host._safe_save(doc, file_path)
            return {
                "success": True,
                "file": file_path,
                "image": image_path,
                "width_inches": width_inches,
                "caption": caption,
                "paragraph_index": paragraph_index,
                "position": "append" if paragraph_index is None else position,
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def set_paragraph_style(
        self, file_path: str, paragraph_index: int, style_name: str
    ) -> Dict[str, Any]:
        """Set a paragraph's Word style."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            with self.host._file_lock(file_path):
                backup = self.host._snapshot_backup(file_path)
                doc = Document(file_path)
                if paragraph_index >= len(doc.paragraphs):
                    return {
                        "success": False,
                        "error": f"Paragraph index {paragraph_index} out of range (max {len(doc.paragraphs) - 1})",
                    }
                para = doc.paragraphs[paragraph_index]
                available_styles = [s.name for s in doc.styles if s.type == 1]
                if style_name not in available_styles:
                    return {
                        "success": False,
                        "error": f"Style '{style_name}' not found. Available: {available_styles[:20]}",
                    }
                para.style = doc.styles[style_name]
                self.host._safe_save(doc, file_path)
            return {
                "success": True,
                "file": file_path,
                "paragraph_index": paragraph_index,
                "style_applied": style_name,
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def set_page_layout(
        self,
        file_path: str,
        orientation: str = None,
        margins: Dict[str, float] = None,
        header_text: str = None,
        page_numbers: bool = False,
        page_size: str = None,
    ) -> Dict[str, Any]:
        """Set page layout (page size, orientation, margins, header, page numbers)."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            from docx.enum.section import WD_ORIENT

            with self.host._file_lock(file_path):
                backup = self.host._snapshot_backup(file_path)
                doc = Document(file_path)
                resolved_page_size = self.host._resolve_page_size(page_size)
                normalized_orientation = (
                    str(orientation).strip().lower() if orientation is not None else None
                )
                if normalized_orientation is not None and normalized_orientation not in {
                    "portrait",
                    "landscape",
                }:
                    return {
                        "success": False,
                        "error": "orientation must be 'portrait' or 'landscape'",
                    }

                for section in doc.sections:
                    effective_orientation = normalized_orientation or self.host._current_orientation(
                        section
                    )
                    if resolved_page_size:
                        _, width, height = resolved_page_size
                        section.page_width = width
                        section.page_height = height

                    if effective_orientation == "landscape":
                        section.orientation = WD_ORIENT.LANDSCAPE
                        if section.page_width < section.page_height:
                            section.page_width, section.page_height = (
                                section.page_height,
                                section.page_width,
                            )
                    else:
                        section.orientation = WD_ORIENT.PORTRAIT
                        if section.page_width > section.page_height:
                            section.page_width, section.page_height = (
                                section.page_height,
                                section.page_width,
                            )

                    if margins:
                        for side, value in margins.items():
                            if hasattr(section, f"{side}_margin"):
                                setattr(section, f"{side}_margin", Inches(value))

                    if header_text is not None:
                        header = section.header
                        header.is_linked_to_previous = False
                        hp = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
                        hp.text = ""
                        run = hp.add_run(header_text)
                        run.font.size = Pt(10)
                        hp.alignment = WD_ALIGN_PARAGRAPH.LEFT

                    if page_numbers:
                        footer = section.footer
                        footer.is_linked_to_previous = False
                        fp = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
                        fp.text = ""
                        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        run = fp.add_run()
                        run.font.size = Pt(10)
                        fld_char_begin = OxmlElement("w:fldChar")
                        fld_char_begin.set(qn("w:fldCharType"), "begin")
                        run._element.append(fld_char_begin)
                        instr_run = fp.add_run()
                        instr_text = OxmlElement("w:instrText")
                        instr_text.set(qn("xml:space"), "preserve")
                        instr_text.text = " PAGE "
                        instr_run._element.append(instr_text)
                        fld_char_end_run = fp.add_run()
                        fld_char_end = OxmlElement("w:fldChar")
                        fld_char_end.set(qn("w:fldCharType"), "end")
                        fld_char_end_run._element.append(fld_char_end)

                self.host._safe_save(doc, file_path)
            return {
                "success": True,
                "file": file_path,
                "layout_updated": True,
                "page_size": resolved_page_size[0] if resolved_page_size else None,
                "orientation": normalized_orientation
                or self.host._current_orientation(doc.sections[0]),
                "header_text": header_text,
                "page_numbers": page_numbers,
                "section_count": len(doc.sections),
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def insert_paragraph(
        self, file_path: str, text: str, index: int, style: str = None
    ) -> Dict[str, Any]:
        """Insert a paragraph at a specific index."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            from docx.text.paragraph import Paragraph

            with self.host._file_lock(file_path):
                backup = self.host._snapshot_backup(file_path)
                doc = Document(file_path)
                total = len(doc.paragraphs)
                if index < 0 or index > total:
                    return {"success": False, "error": f"Index {index} out of range (0..{total})"}
                body = doc.element.body
                new_para = OxmlElement("w:p")
                run = OxmlElement("w:r")
                t = OxmlElement("w:t")
                t.text = text
                t.set(qn("xml:space"), "preserve")
                run.append(t)
                new_para.append(run)
                if index < total:
                    ref = doc.paragraphs[index]._element
                    body.insert(body.index(ref), new_para)
                else:
                    body.append(new_para)
                if style:
                    p = Paragraph(new_para, doc)
                    available = [s.name for s in doc.styles if s.type == 1]
                    if style in available:
                        p.style = doc.styles[style]
                self.host._safe_save(doc, file_path)
            return {
                "success": True,
                "file": file_path,
                "inserted_at": index,
                "text": text[:100],
                "style": style,
                "total_paragraphs": total + 1,
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def delete_paragraph(self, file_path: str, index: int) -> Dict[str, Any]:
        """Delete a paragraph by index."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            with self.host._file_lock(file_path):
                backup = self.host._snapshot_backup(file_path)
                doc = Document(file_path)
                total = len(doc.paragraphs)
                if index < 0 or index >= total:
                    return {
                        "success": False,
                        "error": f"Index {index} out of range (0..{total - 1})",
                    }
                para = doc.paragraphs[index]
                deleted_text = para.text[:200]
                doc.element.body.remove(para._element)
                self.host._safe_save(doc, file_path)
            return {
                "success": True,
                "file": file_path,
                "deleted_index": index,
                "deleted_text": deleted_text,
                "remaining_paragraphs": total - 1,
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def search_document(
        self, file_path: str, search_text: str, case_sensitive: bool = False
    ) -> Dict[str, Any]:
        """Search for text in a document and return all locations."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            doc = Document(file_path)
            results = []
            query = search_text if case_sensitive else search_text.lower()
            for i, para in enumerate(doc.paragraphs):
                text = para.text
                target = text if case_sensitive else text.lower()
                start = 0
                while True:
                    idx = target.find(query, start)
                    if idx == -1:
                        break
                    results.append(
                        {
                            "paragraph_index": i,
                            "char_offset": idx,
                            "context": text[max(0, idx - 30) : idx + len(search_text) + 30],
                            "style": para.style.name if para.style else None,
                        }
                    )
                    start = idx + len(search_text)
            table_results = []
            for ti, table in enumerate(doc.tables):
                for ri, row in enumerate(table.rows):
                    for ci, cell in enumerate(row.cells):
                        text = cell.text
                        target = text if case_sensitive else text.lower()
                        if query in target:
                            table_results.append(
                                {"table_index": ti, "row": ri, "col": ci, "text": text[:200]}
                            )
            return {
                "success": True,
                "file": file_path,
                "search_text": search_text,
                "matches": len(results),
                "table_matches": len(table_results),
                "results": results,
                "table_results": table_results,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def read_tables(self, file_path: str) -> Dict[str, Any]:
        """Read all tables from a document."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            doc = Document(file_path)
            tables = []
            for ti, table in enumerate(doc.tables):
                rows_data = []
                for row in table.rows:
                    rows_data.append([cell.text for cell in row.cells])
                headers = rows_data[0] if rows_data else []
                data = rows_data[1:] if len(rows_data) > 1 else []
                tables.append(
                    {
                        "table_index": ti,
                        "headers": headers,
                        "rows": data,
                        "row_count": len(data),
                        "col_count": len(headers),
                    }
                )
            return {
                "success": True,
                "file": file_path,
                "table_count": len(tables),
                "tables": tables,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def add_hyperlink(
        self, file_path: str, text: str, url: str, paragraph_index: int = -1
    ) -> Dict[str, Any]:
        """Add a hyperlink to the document. paragraph_index=-1 appends a new paragraph."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            with self.host._file_lock(file_path):
                backup = self.host._snapshot_backup(file_path)
                doc = Document(file_path)
                if paragraph_index == -1:
                    para = doc.add_paragraph()
                else:
                    if paragraph_index >= len(doc.paragraphs):
                        return {
                            "success": False,
                            "error": f"Paragraph index {paragraph_index} out of range",
                        }
                    para = doc.paragraphs[paragraph_index]
                part = doc.part
                r_id = part.relate_to(
                    url,
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                    is_external=True,
                )
                hyperlink = OxmlElement("w:hyperlink")
                hyperlink.set(qn("r:id"), r_id)
                run = OxmlElement("w:r")
                rPr = OxmlElement("w:rPr")
                rStyle = OxmlElement("w:rStyle")
                rStyle.set(qn("w:val"), "Hyperlink")
                rPr.append(rStyle)
                color = OxmlElement("w:color")
                color.set(qn("w:val"), "0563C1")
                rPr.append(color)
                u = OxmlElement("w:u")
                u.set(qn("w:val"), "single")
                rPr.append(u)
                run.append(rPr)
                t = OxmlElement("w:t")
                t.text = text
                t.set(qn("xml:space"), "preserve")
                run.append(t)
                hyperlink.append(run)
                para._element.append(hyperlink)
                self.host._safe_save(doc, file_path)
            return {
                "success": True,
                "file": file_path,
                "text": text,
                "url": url,
                "paragraph_index": paragraph_index,
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def add_page_break(self, file_path: str) -> Dict[str, Any]:
        """Add a page break at the end of the document."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            with self.host._file_lock(file_path):
                backup = self.host._snapshot_backup(file_path)
                doc = Document(file_path)
                doc.add_page_break()
                self.host._safe_save(doc, file_path)
            return {
                "success": True,
                "file": file_path,
                "action": "page_break_added",
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def list_styles(self, file_path: str) -> Dict[str, Any]:
        """List all available paragraph styles in the document."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            doc = Document(file_path)
            para_styles = []
            char_styles = []
            for s in doc.styles:
                entry = {"name": s.name, "builtin": s.builtin}
                if s.type == 1:
                    para_styles.append(entry)
                elif s.type == 2:
                    char_styles.append(entry)
            return {
                "success": True,
                "file": file_path,
                "paragraph_styles": para_styles,
                "character_styles": char_styles,
                "paragraph_style_count": len(para_styles),
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def set_metadata(
        self,
        file_path: str,
        author: str = None,
        title: str = None,
        subject: str = None,
        keywords: str = None,
        comments: str = None,
        category: str = None,
    ) -> Dict[str, Any]:
        """Set document core properties (metadata)."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            with self.host._file_lock(file_path):
                backup = self.host._snapshot_backup(file_path)
                doc = Document(file_path)
                cp = doc.core_properties
                if author is not None:
                    cp.author = author
                if title is not None:
                    cp.title = title
                if subject is not None:
                    cp.subject = subject
                if keywords is not None:
                    cp.keywords = keywords
                if comments is not None:
                    cp.comments = comments
                if category is not None:
                    cp.category = category
                self.host._safe_save(doc, file_path)
            return {
                "success": True,
                "file": file_path,
                "metadata_set": {
                    k: v
                    for k, v in {
                        "author": author,
                        "title": title,
                        "subject": subject,
                        "keywords": keywords,
                        "comments": comments,
                        "category": category,
                    }.items()
                    if v is not None
                },
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_metadata(self, file_path: str) -> Dict[str, Any]:
        """Read document core properties (metadata)."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            doc = Document(file_path)
            cp = doc.core_properties
            return {
                "success": True,
                "file": file_path,
                "author": cp.author,
                "title": cp.title,
                "subject": cp.subject,
                "keywords": cp.keywords,
                "comments": cp.comments,
                "category": cp.category,
                "created": cp.created.isoformat() if cp.created else None,
                "modified": cp.modified.isoformat() if cp.modified else None,
                "last_modified_by": cp.last_modified_by,
                "revision": cp.revision,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def add_list(
        self, file_path: str, items: List[str], list_type: str = "bullet"
    ) -> Dict[str, Any]:
        """Add a bulleted or numbered list."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            style_name = "List Bullet" if list_type == "bullet" else "List Number"
            with self.host._file_lock(file_path):
                backup = self.host._snapshot_backup(file_path)
                doc = Document(file_path)
                available = [s.name for s in doc.styles if s.type == 1]
                if style_name not in available:
                    style_name = "Normal"
                for item in items:
                    doc.add_paragraph(item, style=style_name)
                self.host._safe_save(doc, file_path)
            return {
                "success": True,
                "file": file_path,
                "list_type": list_type,
                "items_added": len(items),
                "style_used": style_name,
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def word_count(self, file_path: str) -> Dict[str, Any]:
        """Get word, character, paragraph, and page estimate counts."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            doc = Document(file_path)
            full_text = "\n".join(p.text for p in doc.paragraphs)
            words = len(full_text.split())
            chars = len(full_text)
            chars_no_spaces = len(full_text.replace(" ", "").replace("\n", ""))
            paras = len([p for p in doc.paragraphs if p.text.strip()])
            pages_est = max(1, round(words / 250))
            return {
                "success": True,
                "file": file_path,
                "words": words,
                "characters": chars,
                "characters_no_spaces": chars_no_spaces,
                "paragraphs": paras,
                "pages_estimate": pages_est,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_formatting_info(self, file_path: str) -> Dict[str, Any]:
        """Get detailed formatting information."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            doc = Document(file_path)
            info = {
                "success": True,
                "file": file_path,
                "sections": [],
                "paragraphs": [],
            }
            for section in doc.sections:
                info["sections"].append(
                    {
                        "orientation": "landscape"
                        if section.page_width > section.page_height
                        else "portrait",
                        "margins": {
                            "top": round(section.top_margin.inches, 2),
                            "bottom": round(section.bottom_margin.inches, 2),
                            "left": round(section.left_margin.inches, 2),
                            "right": round(section.right_margin.inches, 2),
                        },
                    }
                )
            for i, para in enumerate(doc.paragraphs[:10]):
                para_info = {
                    "index": i,
                    "alignment": para.alignment,
                    "text_preview": para.text[:50],
                }
                if para.runs:
                    run = para.runs[0]
                    para_info["font"] = {
                        "name": run.font.name,
                        "size": run.font.size.pt if run.font.size else None,
                        "bold": run.bold,
                        "italic": run.italic,
                        "underline": run.underline,
                    }
                info["paragraphs"].append(para_info)
            return info
        except Exception as e:
            return {"success": False, "error": str(e)}

    def inspect_hidden_data(self, file_path: str) -> Dict[str, Any]:
        """Inspect hidden DOCX data relevant to submission prep."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            abs_path = str(Path(file_path).resolve())
            with ZipFile(abs_path, "r") as zin:
                names = zin.namelist()
                story_parts = self.host._docx_story_xml_parts(names)

                comment_reference_count = 0
                tracked_counts = {
                    "insertions": 0,
                    "deletions": 0,
                    "move_to": 0,
                    "move_from": 0,
                    "move_to_range_markers": 0,
                    "move_from_range_markers": 0,
                }
                for part_name in story_parts:
                    root = ET.fromstring(zin.read(part_name))
                    comment_reference_count += len(
                        root.findall(".//w:commentReference", OOXML_NS)
                    )
                    comment_reference_count += len(
                        root.findall(".//w:commentRangeStart", OOXML_NS)
                    )
                    tracked_counts["insertions"] += len(root.findall(".//w:ins", OOXML_NS))
                    tracked_counts["deletions"] += len(root.findall(".//w:del", OOXML_NS))
                    tracked_counts["move_to"] += len(root.findall(".//w:moveTo", OOXML_NS))
                    tracked_counts["move_from"] += len(
                        root.findall(".//w:moveFrom", OOXML_NS)
                    )
                    tracked_counts["move_to_range_markers"] += len(
                        root.findall(".//w:moveToRangeStart", OOXML_NS)
                    ) + len(root.findall(".//w:moveToRangeEnd", OOXML_NS))
                    tracked_counts["move_from_range_markers"] += len(
                        root.findall(".//w:moveFromRangeStart", OOXML_NS)
                    ) + len(root.findall(".//w:moveFromRangeEnd", OOXML_NS))

                comments_count = 0
                if "word/comments.xml" in names:
                    root = ET.fromstring(zin.read("word/comments.xml"))
                    comments_count = len(root.findall(".//w:comment", OOXML_NS))

                settings_flag = False
                if "word/settings.xml" in names:
                    root = ET.fromstring(zin.read("word/settings.xml"))
                    flag = root.find(".//w:removePersonalInformation", OOXML_NS)
                    settings_flag = bool(
                        flag is not None
                        and str(flag.get(f"{{{OOXML_NS['w']}}}val", "true")).lower()
                        != "false"
                    )

                core_props = {}
                if "docProps/core.xml" in names:
                    root = ET.fromstring(zin.read("docProps/core.xml"))
                    core_props = {
                        "author": (
                            root.findtext("dc:creator", default="", namespaces=OOXML_NS)
                            or ""
                        ),
                        "title": (
                            root.findtext("dc:title", default="", namespaces=OOXML_NS)
                            or ""
                        ),
                        "subject": (
                            root.findtext("dc:subject", default="", namespaces=OOXML_NS)
                            or ""
                        ),
                        "keywords": (
                            root.findtext("cp:keywords", default="", namespaces=OOXML_NS)
                            or ""
                        ),
                        "description": (
                            root.findtext(
                                "dc:description", default="", namespaces=OOXML_NS
                            )
                            or ""
                        ),
                        "category": (
                            root.findtext("cp:category", default="", namespaces=OOXML_NS)
                            or ""
                        ),
                        "last_modified_by": (
                            root.findtext(
                                "cp:lastModifiedBy", default="", namespaces=OOXML_NS
                            )
                            or ""
                        ),
                    }

                app_props = {}
                if "docProps/app.xml" in names:
                    root = ET.fromstring(zin.read("docProps/app.xml"))
                    app_props = {
                        "application": root.findtext(
                            "{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}Application",
                            default="",
                        ),
                        "company": root.findtext(
                            "{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}Company",
                            default="",
                        ),
                        "manager": root.findtext(
                            "{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}Manager",
                            default="",
                        ),
                        "hyperlink_base": root.findtext(
                            "{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}HyperlinkBase",
                            default="",
                        ),
                    }

                custom_xml = self.host._custom_xml_summary(names)

            doc = Document(abs_path)
            first_section = doc.sections[0] if doc.sections else None
            page_size = None
            orientation = None
            if first_section is not None:
                page_size = self.host._label_page_size(
                    first_section.page_width, first_section.page_height
                )
                orientation = (
                    "landscape"
                    if first_section.page_width > first_section.page_height
                    else "portrait"
                )

            tracked_total = sum(tracked_counts.values())
            return {
                "success": True,
                "file": abs_path,
                "page_size": page_size,
                "orientation": orientation,
                "section_count": len(doc.sections),
                "comments_part_present": "word/comments.xml" in names,
                "comments_extended_part_present": "word/commentsExtended.xml" in names,
                "people_part_present": "word/people.xml" in names,
                "comments_count": comments_count,
                "comment_reference_count": comment_reference_count,
                "tracked_change_count": tracked_total,
                "tracked_changes_present": tracked_total > 0,
                "tracked_changes": tracked_counts,
                "remove_personal_information": settings_flag,
                "custom_xml_item_count": len(custom_xml["payload_parts"]),
                "custom_xml_part_count": len(custom_xml["payload_parts"]),
                "custom_xml_parts_total_count": len(custom_xml["all_parts"]),
                "custom_xml_support_part_count": len(custom_xml["support_parts"]),
                "custom_xml_parts": custom_xml["payload_parts"],
                "custom_xml_support_parts": custom_xml["support_parts"],
                "core_properties": core_props,
                "app_properties": app_props,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def audit_document_fonts(
        self,
        file_path: str,
        expected_font_name: str = None,
        expected_font_size: float = None,
    ) -> Dict[str, Any]:
        """Audit font consistency across visible text runs."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            doc = Document(file_path)
            runs_payload = self.host._collect_text_runs(doc)
            font_name_counts = Counter(
                run["font_name"] or "<unspecified>" for run in runs_payload
            )
            font_size_counts = Counter(
                run["font_size"] if run["font_size"] is not None else "<unspecified>"
                for run in runs_payload
            )
            dominant_font = (
                font_name_counts.most_common(1)[0][0] if font_name_counts else None
            )
            dominant_size = (
                font_size_counts.most_common(1)[0][0] if font_size_counts else None
            )

            unexpected_font_runs = []
            if expected_font_name:
                target = self.host._normalized_font_name(expected_font_name)
                unexpected_font_runs = [
                    run for run in runs_payload if run["font_name_normalized"] != target
                ]

            unexpected_size_runs = []
            if expected_font_size is not None:
                target_size = round(float(expected_font_size), 2)
                unexpected_size_runs = [
                    run for run in runs_payload if run["font_size"] != target_size
                ]

            mixed_fonts = len(font_name_counts) > 1
            mixed_sizes = len(font_size_counts) > 1
            return {
                "success": True,
                "file": file_path,
                "text_run_count": len(runs_payload),
                "font_name_counts": dict(font_name_counts),
                "font_size_counts": dict(font_size_counts),
                "dominant_font_name": dominant_font,
                "dominant_font_size": dominant_size,
                "mixed_fonts": mixed_fonts,
                "mixed_sizes": mixed_sizes,
                "expected_font_name": expected_font_name,
                "expected_font_size": expected_font_size,
                "unexpected_font_run_count": len(unexpected_font_runs),
                "unexpected_size_run_count": len(unexpected_size_runs),
                "unexpected_font_examples": unexpected_font_runs[:10],
                "unexpected_size_examples": unexpected_size_runs[:10],
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def audit_document_images(self, file_path: str) -> Dict[str, Any]:
        """Audit inline image sizing against available page content area."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            doc = Document(file_path)
            sections = self.host._document_sections_summary(doc)
            max_usable_width = max(
                (s["usable_width_in"] for s in sections), default=0.0
            )
            max_usable_height = max(
                (s["usable_height_in"] for s in sections), default=0.0
            )

            oversized = []
            near_limit = []
            images = []
            for idx, shape in enumerate(doc.inline_shapes):
                width_in = round(shape.width.inches, 3)
                height_in = round(shape.height.inches, 3)
                width_ratio = (
                    round(width_in / max_usable_width, 3) if max_usable_width else None
                )
                height_ratio = (
                    round(height_in / max_usable_height, 3)
                    if max_usable_height
                    else None
                )
                image_info = {
                    "index": idx,
                    "width_in": width_in,
                    "height_in": height_in,
                    "width_ratio_to_max_content": width_ratio,
                    "height_ratio_to_max_content": height_ratio,
                }
                if max_usable_width and width_in > max_usable_width + 0.02:
                    oversized.append(image_info)
                elif max_usable_width and width_ratio is not None and width_ratio >= 0.95:
                    near_limit.append(image_info)
                images.append(image_info)

            return {
                "success": True,
                "file": file_path,
                "inline_image_count": len(images),
                "section_max_usable_width_in": max_usable_width,
                "section_max_usable_height_in": max_usable_height,
                "images": images,
                "oversized_image_count": len(oversized),
                "oversized_images": oversized[:10],
                "near_limit_image_count": len(near_limit),
                "near_limit_images": near_limit[:10],
                "note": "Only inline images are audited. Floating/anchored images are not inspected by python-docx.",
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def document_preflight(
        self,
        file_path: str,
        expected_page_size: str = None,
        expected_font_name: str = None,
        expected_font_size: float = None,
    ) -> Dict[str, Any]:
        """Run a submission-oriented DOCX preflight."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            abs_path = str(Path(file_path).resolve())
            hidden = self.inspect_hidden_data(abs_path)
            if not hidden.get("success"):
                return hidden
            fonts = self.audit_document_fonts(
                abs_path,
                expected_font_name=expected_font_name,
                expected_font_size=expected_font_size,
            )
            if not fonts.get("success"):
                return fonts
            images = self.audit_document_images(abs_path)
            if not images.get("success"):
                return images

            sections = self.host._document_sections_summary(Document(abs_path))
            section_page_sizes = sorted({s["page_size"] for s in sections})
            page_size_consistent = len(section_page_sizes) <= 1

            checks = []

            def add_check(
                name: str, status: str, message: str, details: Dict[str, Any] = None
            ):
                checks.append(
                    {
                        "name": name,
                        "status": status,
                        "message": message,
                        "details": details or {},
                    }
                )

            if expected_page_size:
                matches_expected = (
                    page_size_consistent
                    and len(section_page_sizes) == 1
                    and section_page_sizes[0].lower() == expected_page_size.lower()
                )
                add_check(
                    "page_size",
                    "pass" if matches_expected else "warn",
                    (
                        f"All sections match expected page size {expected_page_size}."
                        if matches_expected
                        else f"Detected page sizes {section_page_sizes}; expected {expected_page_size}."
                    ),
                    {"detected": section_page_sizes, "expected": expected_page_size},
                )
            else:
                add_check(
                    "page_size",
                    "pass" if page_size_consistent else "warn",
                    (
                        f"All sections use {section_page_sizes[0]}."
                        if page_size_consistent and section_page_sizes
                        else f"Mixed section page sizes detected: {section_page_sizes}."
                    ),
                    {"detected": section_page_sizes},
                )

            add_check(
                "comments",
                "pass"
                if hidden["comments_count"] == 0 and hidden["comment_reference_count"] == 0
                else "fail",
                (
                    "No comment parts or references detected."
                    if hidden["comments_count"] == 0
                    and hidden["comment_reference_count"] == 0
                    else f"Detected {hidden['comments_count']} comments and {hidden['comment_reference_count']} comment references."
                ),
                {
                    "comments_count": hidden["comments_count"],
                    "comment_reference_count": hidden["comment_reference_count"],
                },
            )

            add_check(
                "tracked_changes",
                "pass" if not hidden["tracked_changes_present"] else "fail",
                (
                    "No tracked revision markup detected."
                    if not hidden["tracked_changes_present"]
                    else f"Detected {hidden['tracked_change_count']} tracked-change elements."
                ),
                hidden["tracked_changes"],
            )

            metadata_values = hidden["core_properties"].copy()
            metadata_values.update(hidden["app_properties"])
            nonempty_metadata = {
                k: v for k, v in metadata_values.items() if str(v or "").strip()
            }
            add_check(
                "metadata",
                "pass" if not nonempty_metadata else "warn",
                (
                    "No non-empty hidden core/app metadata fields detected."
                    if not nonempty_metadata
                    else f"Detected non-empty metadata fields: {sorted(nonempty_metadata)}."
                ),
                nonempty_metadata,
            )

            add_check(
                "custom_xml",
                "pass" if hidden["custom_xml_item_count"] == 0 else "warn",
                (
                    "No custom XML payload items detected."
                    if hidden["custom_xml_item_count"] == 0
                    else f"Detected {hidden['custom_xml_item_count']} custom XML payload item(s)."
                ),
                {
                    "custom_xml_items": hidden["custom_xml_parts"],
                    "custom_xml_support_parts": hidden["custom_xml_support_parts"],
                },
            )

            add_check(
                "remove_personal_information_flag",
                "pass" if hidden["remove_personal_information"] else "warn",
                (
                    "removePersonalInformation flag is enabled."
                    if hidden["remove_personal_information"]
                    else "removePersonalInformation flag is not enabled."
                ),
            )

            font_status = "pass"
            font_message = "Document font usage looks consistent."
            if expected_font_name and fonts["unexpected_font_run_count"] > 0:
                font_status = "warn"
                font_message = (
                    f"Detected {fonts['unexpected_font_run_count']} run(s) not matching expected font '{expected_font_name}'."
                )
            elif fonts["mixed_fonts"]:
                font_status = "warn"
                font_message = (
                    f"Multiple font families detected; dominant font is '{fonts['dominant_font_name']}'."
                )
            add_check(
                "font_names",
                font_status,
                font_message,
                {
                    "font_name_counts": fonts["font_name_counts"],
                    "dominant_font_name": fonts["dominant_font_name"],
                    "unexpected_examples": fonts["unexpected_font_examples"],
                },
            )

            size_status = "pass"
            size_message = "Document font sizes look consistent."
            if expected_font_size is not None and fonts["unexpected_size_run_count"] > 0:
                size_status = "warn"
                size_message = (
                    f"Detected {fonts['unexpected_size_run_count']} run(s) not matching expected font size {expected_font_size}."
                )
            elif fonts["mixed_sizes"]:
                size_status = "warn"
                size_message = (
                    f"Multiple font sizes detected; dominant size is {fonts['dominant_font_size']}."
                )
            add_check(
                "font_sizes",
                size_status,
                size_message,
                {
                    "font_size_counts": fonts["font_size_counts"],
                    "dominant_font_size": fonts["dominant_font_size"],
                    "unexpected_examples": fonts["unexpected_size_examples"],
                },
            )

            image_status = "pass"
            image_message = "Inline images fit within the maximum content area."
            if images["oversized_image_count"] > 0:
                image_status = "warn"
                image_message = (
                    f"Detected {images['oversized_image_count']} inline image(s) wider than the available content width."
                )
            elif images["near_limit_image_count"] > 0:
                image_status = "warn"
                image_message = (
                    f"Detected {images['near_limit_image_count']} inline image(s) at or near the available content width."
                )
            add_check(
                "images",
                image_status,
                image_message,
                {
                    "oversized_images": images["oversized_images"],
                    "near_limit_images": images["near_limit_images"],
                    "note": images["note"],
                },
            )

            overall = "pass"
            if any(check["status"] == "fail" for check in checks):
                overall = "fail"
            elif any(check["status"] == "warn" for check in checks):
                overall = "warn"

            return {
                "success": True,
                "file": abs_path,
                "overall_status": overall,
                "checks": checks,
                "sections": sections,
                "hidden_data": hidden,
                "font_audit": fonts,
                "image_audit": images,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def sanitize_document(
        self,
        file_path: str,
        *,
        remove_comments: bool = False,
        accept_revisions: bool = False,
        clear_metadata: bool = False,
        remove_custom_xml: bool = False,
        set_remove_personal_information: bool = False,
        author: Optional[str] = None,
        title: Optional[str] = None,
        subject: Optional[str] = None,
        keywords: Optional[str] = None,
        output_path: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Sanitize a DOCX for submission by removing hidden data and normalising metadata."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        if not any(
            [
                remove_comments,
                accept_revisions,
                clear_metadata,
                remove_custom_xml,
                set_remove_personal_information,
                author is not None,
                title is not None,
                subject is not None,
                keywords is not None,
            ]
        ):
            return {"success": False, "error": "No sanitization options provided"}
        try:
            abs_path = str(Path(file_path).resolve())
            out_path = str(Path(output_path).resolve()) if output_path else abs_path
            overwrite = out_path == abs_path
            with self.host._file_lock(abs_path):
                backup_target = abs_path if overwrite else out_path
                backup = (
                    self.host._snapshot_backup(backup_target)
                    if Path(backup_target).exists()
                    else None
                )
                before = self.inspect_hidden_data(abs_path)
                if not before.get("success"):
                    return before

                with ZipFile(abs_path, "r") as zin:
                    files = {name: zin.read(name) for name in zin.namelist()}

                stats: Dict[str, int] = {}
                if remove_comments or accept_revisions:
                    for part_name in self.host._docx_story_xml_parts(list(files.keys())):
                        root = ET.fromstring(files[part_name])
                        self.host._rewrite_story_tree(
                            root,
                            remove_comment_nodes=remove_comments,
                            accept_revisions=accept_revisions,
                            stats=stats,
                        )
                        files[part_name] = ET.tostring(
                            root, encoding="utf-8", xml_declaration=True
                        )

                if remove_comments:
                    removed_parts = []
                    for part_name in [
                        "word/comments.xml",
                        "word/commentsExtended.xml",
                        "word/people.xml",
                    ]:
                        if part_name in files:
                            removed_parts.append(part_name)
                            del files[part_name]

                    self.host._strip_docx_relationship_targets(
                        files,
                        lambda target, rel_type: "comments" in rel_type
                        or "comments" in target.lower()
                        or target.lower().endswith("people.xml"),
                    )
                    self.host._strip_docx_content_types(
                        files,
                        lambda child: "comments" in child.get("PartName", "").lower()
                        or child.get("PartName", "").lower().endswith("/people.xml"),
                    )
                    if removed_parts:
                        stats["comment_parts_removed"] = len(removed_parts)

                if remove_custom_xml:
                    custom_xml = self.host._custom_xml_summary(list(files.keys()))
                    removed_custom = list(custom_xml["all_parts"])
                    for name in removed_custom:
                        del files[name]
                    self.host._strip_docx_relationship_targets(
                        files,
                        lambda target, rel_type: target.startswith("customXml/")
                        or "/customXml" in rel_type,
                    )
                    self.host._strip_docx_content_types(
                        files,
                        lambda child: child.get("PartName", "").startswith("/customXml/"),
                    )
                    stats["custom_xml_items_removed"] = len(custom_xml["payload_parts"])
                    stats["custom_xml_support_parts_removed"] = len(
                        custom_xml["support_parts"]
                    )
                    stats["custom_xml_parts_removed"] = len(removed_custom)

                if clear_metadata or any(
                    opt is not None for opt in [author, title, subject, keywords]
                ):
                    core_name = "docProps/core.xml"
                    if core_name in files:
                        root = ET.fromstring(files[core_name])
                        field_map = {
                            "title": ("dc", "title", title),
                            "subject": ("dc", "subject", subject),
                            "creator": ("dc", "creator", author),
                            "keywords": ("cp", "keywords", keywords),
                            "description": ("dc", "description", ""),
                            "category": ("cp", "category", ""),
                            "contentStatus": ("cp", "contentStatus", ""),
                            "identifier": ("dc", "identifier", ""),
                            "language": ("dc", "language", ""),
                            "lastModifiedBy": ("cp", "lastModifiedBy", ""),
                            "version": ("cp", "version", ""),
                        }
                        for _, (prefix, local, explicit_value) in field_map.items():
                            node = root.find(f"{prefix}:{local}", OOXML_NS)
                            if node is None:
                                continue
                            if explicit_value is not None:
                                node.text = explicit_value
                            elif clear_metadata:
                                node.text = ""
                        revision = root.find("cp:revision", OOXML_NS)
                        if revision is not None and clear_metadata:
                            revision.text = "1"
                        files[core_name] = ET.tostring(
                            root, encoding="utf-8", xml_declaration=True
                        )

                    app_name = "docProps/app.xml"
                    if app_name in files and clear_metadata:
                        root = ET.fromstring(files[app_name])
                        for tag in [
                            "{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}Manager",
                            "{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}Company",
                            "{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}HyperlinkBase",
                        ]:
                            node = root.find(tag)
                            if node is not None:
                                node.text = ""
                        files[app_name] = ET.tostring(
                            root, encoding="utf-8", xml_declaration=True
                        )

                if set_remove_personal_information or clear_metadata:
                    settings_name = "word/settings.xml"
                    if settings_name in files:
                        root = ET.fromstring(files[settings_name])
                        flag = root.find("w:removePersonalInformation", OOXML_NS)
                        if flag is None:
                            flag = ET.Element(
                                f"{{{OOXML_NS['w']}}}removePersonalInformation"
                            )
                            flag.set(f"{{{OOXML_NS['w']}}}val", "true")
                            insert_at = 1 if len(root) >= 1 else 0
                            root.insert(insert_at, flag)
                        else:
                            flag.set(f"{{{OOXML_NS['w']}}}val", "true")
                        files[settings_name] = ET.tostring(
                            root, encoding="utf-8", xml_declaration=True
                        )

                self.host._atomic_zip_write(files, out_path)

                after = self.inspect_hidden_data(out_path)
                if not after.get("success"):
                    return after

            return {
                "success": True,
                "file": out_path,
                "input_file": abs_path,
                "output_file": out_path,
                "remove_comments": remove_comments,
                "accept_revisions": accept_revisions,
                "clear_metadata": clear_metadata,
                "remove_custom_xml": remove_custom_xml,
                "set_remove_personal_information": set_remove_personal_information
                or clear_metadata,
                "stats": stats,
                "before": before,
                "after": after,
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def extract_images_from_docx(
        self,
        file_path: str,
        output_dir: str,
        fmt: str = "png",
        prefix: str = "image",
    ) -> Dict[str, Any]:
        """Extract all embedded images from a .docx file."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            from PIL import Image as PILImage
            import io

            os.makedirs(output_dir, exist_ok=True)
            doc = Document(file_path)
            extracted = []
            idx = 0
            for rel in doc.part.rels.values():
                if "image" not in rel.reltype:
                    continue
                try:
                    image_part = rel.target_part
                    image_bytes = image_part.blob
                    content_type = image_part.content_type
                    ext_map = {
                        "image/png": "png",
                        "image/jpeg": "jpg",
                        "image/gif": "gif",
                        "image/bmp": "bmp",
                        "image/tiff": "tiff",
                        "image/x-emf": "emf",
                        "image/x-wmf": "wmf",
                    }
                    src_ext = ext_map.get(content_type, "png")
                    out_name = f"{prefix}_{idx:03d}.{fmt}"
                    out_path = os.path.join(output_dir, out_name)
                    if src_ext in ("emf", "wmf"):
                        raw_name = f"{prefix}_{idx:03d}.{src_ext}"
                        raw_path = os.path.join(output_dir, raw_name)
                        with open(raw_path, "wb") as f:
                            f.write(image_bytes)
                        extracted.append(
                            {
                                "index": idx,
                                "file": raw_path,
                                "format": src_ext,
                                "size_bytes": len(image_bytes),
                                "note": f"Vector format saved as .{src_ext} (cannot convert to {fmt})",
                            }
                        )
                    else:
                        img = PILImage.open(io.BytesIO(image_bytes))
                        if fmt.lower() == "jpg":
                            img = img.convert("RGB")
                            img.save(out_path, "JPEG", quality=90)
                        else:
                            img.save(out_path, fmt.upper())
                        extracted.append(
                            {
                                "index": idx,
                                "file": out_path,
                                "format": fmt,
                                "width": img.width,
                                "height": img.height,
                                "size_bytes": os.path.getsize(out_path),
                            }
                        )
                    idx += 1
                except Exception as img_err:
                    extracted.append({"index": idx, "error": str(img_err)})
                    idx += 1
            return {
                "success": True,
                "file": file_path,
                "output_dir": output_dir,
                "images_extracted": len(extracted),
                "images": extracted,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def doc_render_map(self, file_path: str) -> Dict[str, Any]:
        """Build deterministic paragraph/table-cell to rendered PDF page/bbox mappings."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            abs_path = str(Path(file_path).resolve())
            if not os.path.exists(abs_path):
                return {"success": False, "error": f"File not found: {file_path}"}

            with tempfile.TemporaryDirectory(prefix="onlyoffice-doc-render-map-") as tmpdir:
                temp_docx = os.path.join(tmpdir, Path(abs_path).name)
                temp_pdf = os.path.join(tmpdir, f"{Path(abs_path).stem}.pdf")
                shutil.copy2(abs_path, temp_docx)

                doc = Document(temp_docx)
                paragraph_anchors = []
                table_cell_anchors = []

                paragraph_index = 0
                for paragraph in doc.paragraphs:
                    original_text = re.sub(r"\s+", " ", paragraph.text or "").strip()
                    if not original_text:
                        continue
                    paragraph_index += 1
                    anchor_id = f"SLOANE_P_{paragraph_index:04d}"
                    paragraph.text = f"{anchor_id} {original_text}"
                    paragraph_anchors.append(
                        {
                            "anchor_id": anchor_id,
                            "paragraph_index": paragraph_index,
                            "text_preview": original_text[:220],
                        }
                    )

                for table_number, table in enumerate(doc.tables, start=1):
                    for row_number, row in enumerate(table.rows, start=1):
                        for column_number, cell in enumerate(row.cells, start=1):
                            cell_text = re.sub(
                                r"\s+",
                                " ",
                                " ".join(
                                    paragraph.text or "" for paragraph in cell.paragraphs
                                ),
                            ).strip()
                            if not cell_text:
                                continue
                            anchor_id = (
                                f"SLOANE_T{table_number}R{row_number}C{column_number}"
                            )
                            cell.text = f"{anchor_id} {cell_text}"
                            table_cell_anchors.append(
                                {
                                    "anchor_id": anchor_id,
                                    "table_number": table_number,
                                    "row_number": row_number,
                                    "column_number": column_number,
                                    "cell_ref": f"T{table_number}R{row_number}C{column_number}",
                                    "text_preview": cell_text[:220],
                                }
                            )

                doc.save(temp_docx)

                pdf_result = self.host.doc_to_pdf(temp_docx, output_path=temp_pdf)
                if not pdf_result.get("success"):
                    return pdf_result

                block_payload = self.host.pdf_read_blocks(
                    temp_pdf,
                    include_spans=True,
                    include_images=False,
                    include_empty=False,
                )
                if not block_payload.get("success"):
                    return block_payload

                marker_hits = {}
                marker_pattern = re.compile(r"SLOANE_(?:P_\d{4}|T\d+R\d+C\d+)")

                for page in block_payload.get("pages", []):
                    for block in page.get("blocks", []):
                        if block.get("type") != "text":
                            continue
                        for line in block.get("lines", []):
                            for span in line.get("spans", []):
                                span_text = str(span.get("text", "") or "")
                                for match in marker_pattern.finditer(span_text):
                                    anchor_id = match.group(0)
                                    if anchor_id in marker_hits:
                                        continue
                                    marker_hits[anchor_id] = {
                                        "page_index": page.get("page_index"),
                                        "page_number": page.get("page_number"),
                                        "block_id": block.get("block_id"),
                                        "line_id": line.get("line_id"),
                                        "span_id": span.get("span_id"),
                                        "bbox": span.get("bbox") or block.get("bbox"),
                                        "block_bbox": block.get("bbox"),
                                        "text": block.get("text"),
                                    }

                paragraph_mappings = []
                table_cell_mappings = []
                unresolved_anchor_ids = []

                for anchor in paragraph_anchors:
                    hit = marker_hits.get(anchor["anchor_id"])
                    if not hit:
                        unresolved_anchor_ids.append(anchor["anchor_id"])
                        continue
                    paragraph_mappings.append({**anchor, **hit})

                for anchor in table_cell_anchors:
                    hit = marker_hits.get(anchor["anchor_id"])
                    if not hit:
                        unresolved_anchor_ids.append(anchor["anchor_id"])
                        continue
                    table_cell_mappings.append({**anchor, **hit})

                return {
                    "success": True,
                    "file": abs_path,
                    "paragraph_count": len(paragraph_anchors),
                    "table_cell_count": len(table_cell_anchors),
                    "mapped_paragraph_count": len(paragraph_mappings),
                    "mapped_table_cell_count": len(table_cell_mappings),
                    "pages": block_payload.get("pages_scanned", 0),
                    "paragraphs": paragraph_mappings,
                    "table_cells": table_cell_mappings,
                    "unresolved_anchor_ids": unresolved_anchor_ids,
                }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def doc_to_pdf(self, file_path: str, output_path: str = None) -> Dict[str, Any]:
        """Convert a .docx file to PDF via OnlyOffice x2t."""
        return self.host._office_to_pdf(file_path, output_path=output_path)

    def preview_document(
        self,
        file_path: str,
        output_dir: str,
        pages: str = None,
        dpi: int = 150,
        fmt: str = "png",
    ) -> Dict[str, Any]:
        """Render DOCX pages as images via OnlyOffice conversion + PyMuPDF."""
        return self.host._preview_via_pdf(
            file_path, output_dir, self.host.doc_to_pdf, pages=pages, dpi=dpi, fmt=fmt
        )
