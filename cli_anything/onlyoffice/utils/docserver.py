#!/usr/bin/env python3
"""
OnlyOffice Document Server API Client - Advanced Formatting Edition

Provides full programmatic control with advanced formatting:
- Documents (.docx) - Create, read, edit, format, highlight, comments
- Spreadsheets (.xlsx) - Create, read, edit, formulas, styling
"""

import requests
import json
import os
import shutil
import tempfile
import fcntl
import re
import hashlib
import threading
from pathlib import Path
from datetime import datetime, timezone
from typing import Dict, Any, List
from contextlib import contextmanager

# Module-level per-path thread locks — serialises same-process threads before
# flock takes over for cross-process serialisation.
_thread_lock_map: dict[str, threading.Lock] = {}
_thread_lock_map_mu: threading.Lock = threading.Lock()

# Import document libraries with advanced formatting
try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor, Mm
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
    from docx.enum.section import WD_ORIENT
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.chart import (
        BarChart,
        LineChart,
        PieChart,
        ScatterChart,
        Reference,
    )
    from openpyxl.chart.label import DataLabelList
    from openpyxl.drawing.image import Image as XLImage
    import openpyxl.utils

    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.enum.text import PP_ALIGN
    from pptx.dml.color import RGBColor
    from pptx.util import Inches, Pt

    PPTX_AVAILABLE = True
except ImportError:
    PPTX_AVAILABLE = False

try:
    from scipy import stats as scipy_stats

    SCIPY_AVAILABLE = True
except ImportError:
    SCIPY_AVAILABLE = False


class DocumentServerClient:
    """Enhanced client for full document and spreadsheet control"""

    def __init__(
        self, server_url="http://localhost:8080", secret="sloane-os-secret-key-2026"
    ):
        self.server_url = server_url.rstrip("/")
        self.secret = secret
        self.headers = {"Authorization": f"Bearer {secret}"}
        self.backup_dir = Path(
            os.environ.get("ONLYOFFICE_BACKUP_DIR", "/tmp/sloane-onlyoffice-backups")
        )
        self.backup_dir.mkdir(parents=True, exist_ok=True)
        self._lock_timeout_seconds = int(
            os.environ.get("ONLYOFFICE_LOCK_TIMEOUT", "15")
        )
        self.supported_formula_functions = {
            "SUM",
            "AVERAGE",
            "AVG",
            "MIN",
            "MAX",
            "ABS",
            "ROUND",
            "COUNT",
            "COUNTA",
        }

    def _file_key(self, file_path: str) -> str:
        abs_path = str(Path(file_path).resolve())
        return hashlib.sha1(abs_path.encode("utf-8")).hexdigest()[:10]

    @contextmanager
    def _file_lock(self, file_path: str):
        """Two-layer file lock: threading.Lock (intra-process) + fcntl.flock (cross-process).

        fcntl.flock is per-process on Linux — threads within the same process bypass it.
        The threading.Lock layer serialises concurrent threads before flock serialises
        concurrent processes.
        """
        import time
        abs_path = str(Path(file_path).resolve())
        # Retrieve or create the per-path thread lock
        with _thread_lock_map_mu:
            if abs_path not in _thread_lock_map:
                _thread_lock_map[abs_path] = threading.Lock()
            tlock = _thread_lock_map[abs_path]

        with tlock:  # serialise threads within this process
            lock_path = f"{file_path}.lock"
            lock_file = open(lock_path, "w")
            deadline = time.monotonic() + self._lock_timeout_seconds
            acquired = False
            try:
                while True:
                    try:
                        fcntl.flock(lock_file.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                        acquired = True
                        break
                    except BlockingIOError:
                        if time.monotonic() >= deadline:
                            raise TimeoutError(
                                f"Could not acquire lock on {file_path} within "
                                f"{self._lock_timeout_seconds}s"
                            )
                        time.sleep(0.05)
                yield
            finally:
                if acquired:
                    try:
                        fcntl.flock(lock_file.fileno(), fcntl.LOCK_UN)
                    except Exception:
                        pass
                lock_file.close()
                try:
                    os.unlink(lock_path)
                except OSError:
                    pass

    def _get_sheet(self, wb, sheet_name: str):
        """Get a worksheet with a friendly error if not found."""
        if sheet_name in wb.sheetnames:
            return wb[sheet_name]
        raise KeyError(
            f"Sheet '{sheet_name}' not found. Available sheets: {wb.sheetnames}. "
            f"Use --sheet {wb.sheetnames[0]} to target the first sheet."
        )

    def _snapshot_backup(self, file_path: str) -> str:
        """Create a point-in-time backup snapshot before mutating existing files."""
        src = Path(file_path)
        if not src.exists():
            return ""
        stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S%fZ")
        backup_name = f"{src.name}.{self._file_key(file_path)}.{stamp}.bak{src.suffix}"
        dst = self.backup_dir / backup_name
        shutil.copy2(src, dst)
        return str(dst)

    def _safe_save(self, wb_or_doc, file_path: str):
        """Atomic save to reduce risk of partial writes."""
        target = Path(file_path)
        fd, tmp_path = tempfile.mkstemp(
            prefix=f".{target.name}.", dir=str(target.parent)
        )
        os.close(fd)
        try:
            wb_or_doc.save(tmp_path)
            os.replace(tmp_path, str(target))
        finally:
            if os.path.exists(tmp_path):
                os.unlink(tmp_path)

    def _validate_tabular_rows(
        self,
        headers: List[Any],
        rows: List[List[Any]],
        coerce_rows: bool = False,
    ):
        expected_cols = len(headers)
        if expected_cols == 0:
            return False, "Headers cannot be empty", []

        normalized_rows = []
        for idx, row in enumerate(rows, 1):
            row_len = len(row)
            if row_len != expected_cols:
                if not coerce_rows:
                    return (
                        False,
                        f"Row {idx} has {row_len} columns, expected {expected_cols}. Use --coerce-rows to pad/truncate.",
                        [],
                    )
                if row_len < expected_cols:
                    row = row + [""] * (expected_cols - row_len)
                else:
                    row = row[:expected_cols]
            normalized_rows.append(row)

        return True, None, normalized_rows

    def _backup_candidates(self, file_path: str):
        p = Path(file_path)
        key = self._file_key(file_path)
        suffix = p.suffix
        candidates = []
        for bf in self.backup_dir.glob(f"{p.name}*.bak{suffix}"):
            name = bf.name
            if f".{key}." in name or re.match(
                rf"^{re.escape(p.name)}\.\d{{8}}T\d{{12}}Z\.bak{re.escape(suffix)}$",
                name,
            ):
                candidates.append(bf)
        return sorted(candidates, key=lambda x: x.stat().st_mtime, reverse=True)

    def list_backups(self, file_path: str, limit: int = 20) -> Dict[str, Any]:
        try:
            items = []
            for bf in self._backup_candidates(file_path)[: max(1, int(limit))]:
                st = bf.stat()
                items.append(
                    {
                        "backup": str(bf),
                        "name": bf.name,
                        "size": st.st_size,
                        "modified": datetime.fromtimestamp(
                            st.st_mtime, timezone.utc
                        ).isoformat(),
                    }
                )
            return {
                "success": True,
                "file": file_path,
                "count": len(items),
                "backups": items,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def prune_backups(
        self,
        file_path: str = None,
        keep: int = 20,
        older_than_days: int = None,
    ) -> Dict[str, Any]:
        try:
            keep = max(0, int(keep))
            now = datetime.now(timezone.utc).timestamp()
            if file_path:
                groups = {str(Path(file_path)): self._backup_candidates(file_path)}
            else:
                groups = {}
                for bf in self.backup_dir.glob("*.bak*"):
                    name = bf.name
                    parts = name.split(".")
                    if len(parts) < 4:
                        group = name
                    elif re.fullmatch(r"[a-f0-9]{10}", parts[-4]):
                        group = f"{'.'.join(parts[:-4])}.{parts[-4]}.{parts[-1]}"
                    else:
                        group = f"{'.'.join(parts[:-3])}.{parts[-1]}"
                    groups.setdefault(group, []).append(bf)

            deleted = []
            scanned = 0
            for files in groups.values():
                files = sorted(files, key=lambda x: x.stat().st_mtime, reverse=True)
                scanned += len(files)
                for idx, bf in enumerate(files):
                    st = bf.stat()
                    too_old = False
                    if older_than_days is not None:
                        too_old = (now - st.st_mtime) > (int(older_than_days) * 86400)
                    over_keep = idx >= keep
                    if too_old or over_keep:
                        try:
                            bf.unlink()
                            deleted.append(str(bf))
                        except Exception:
                            pass

            return {
                "success": True,
                "scanned": scanned,
                "deleted": len(deleted),
                "deleted_files": deleted,
                "keep": keep,
                "older_than_days": older_than_days,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def restore_backup(
        self,
        file_path: str,
        backup: str = None,
        latest: bool = False,
        dry_run: bool = False,
    ) -> Dict[str, Any]:
        try:
            target = Path(file_path)
            if backup:
                chosen = Path(backup)
                if not chosen.is_absolute():
                    chosen = self.backup_dir / backup
                if not chosen.exists():
                    return {"success": False, "error": f"Backup not found: {backup}"}
            else:
                backups = self._backup_candidates(file_path)
                if not backups:
                    return {"success": False, "error": "No backups found for file"}
                chosen = backups[0] if latest or not backup else None

            if dry_run:
                return {
                    "success": True,
                    "file": file_path,
                    "restore_from": str(chosen),
                    "dry_run": True,
                }

            with self._file_lock(file_path):
                pre_restore_backup = (
                    self._snapshot_backup(file_path) if target.exists() else ""
                )
                fd, tmp_path = tempfile.mkstemp(
                    prefix=f".{target.name}.restore.", dir=str(target.parent)
                )
                os.close(fd)
                try:
                    shutil.copy2(chosen, tmp_path)
                    os.replace(tmp_path, str(target))
                finally:
                    if os.path.exists(tmp_path):
                        os.unlink(tmp_path)

            return {
                "success": True,
                "file": file_path,
                "restored_from": str(chosen),
                "pre_restore_backup": pre_restore_backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ==================== BASIC DOCUMENT OPERATIONS ====================

    def create_document(
        self, output_path: str, title: str = "", content: str = ""
    ) -> Dict[str, Any]:
        """Create a new .docx document"""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            with self._file_lock(output_path):
                backup = self._snapshot_backup(output_path)
                doc = Document()
                # Page layout: A4, 1" margins all sides
                section = doc.sections[0]
                section.page_width = Mm(210)
                section.page_height = Mm(297)
                section.top_margin = Pt(72)     # 1 inch = 72pt
                section.bottom_margin = Pt(72)
                section.left_margin = Pt(72)
                section.right_margin = Pt(72)
                # Normal style: Calibri 11pt, double spacing, 0pt space-after
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
                self._safe_save(doc, output_path)
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
        """Read and extract all text from a .docx document"""
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
        """Append content to a .docx document"""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
                doc = Document(file_path)
                for paragraph in content.split("\n"):
                    if paragraph.strip():
                        doc.add_paragraph(paragraph)
                self._safe_save(doc, file_path)
            return {
                "success": True,
                "file": file_path,
                "appended_length": len(content),
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def _replace_across_runs(self, paragraph, search_text: str, replace_text: str) -> int:
        """Replace text that may span multiple runs while preserving run-level formatting.
        Uses runs-only text universe (ignores tracked changes, hyperlinks, field codes).
        Collects all match positions first, then applies back-to-front to avoid infinite loops."""
        runs = paragraph.runs
        if not runs:
            return 0
        # Build text and char_map from runs only (not paragraph.text which includes non-run nodes)
        full = "".join(r.text for r in runs)
        if search_text not in full:
            return 0

        # Collect ALL match positions on the original string first (avoids infinite loop
        # when replace_text contains search_text)
        matches = []
        start = 0
        while True:
            idx = full.find(search_text, start)
            if idx == -1:
                break
            matches.append(idx)
            start = idx + len(search_text)  # non-overlapping

        if not matches:
            return 0

        # Apply replacements back-to-front so earlier positions stay valid
        for idx in reversed(matches):
            # Rebuild char_map fresh for current run state
            char_map = []
            for ri, run in enumerate(runs):
                for ci in range(len(run.text)):
                    char_map.append((ri, ci))

            end_idx = idx + len(search_text)
            if end_idx > len(char_map) or idx >= len(char_map):
                continue  # skip: match extends beyond run-text (tracked change / hyperlink region)

            first_run, first_offset = char_map[idx]
            last_run, last_offset = char_map[end_idx - 1]

            if first_run == last_run:
                r = runs[first_run]
                r.text = r.text[:first_offset] + replace_text + r.text[last_offset + 1:]
            else:
                fr = runs[first_run]
                fr.text = fr.text[:first_offset] + replace_text
                for mi in range(first_run + 1, last_run):
                    runs[mi].text = ""
                lr = runs[last_run]
                lr.text = lr.text[last_offset + 1:]

        return len(matches)

    def search_replace_document(
        self, file_path: str, search_text: str, replace_text: str
    ) -> Dict[str, Any]:
        """Find and replace text in a .docx document (handles cross-run text splits)"""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
                doc = Document(file_path)
                replacements = 0
                for para in doc.paragraphs:
                    replacements += self._replace_across_runs(para, search_text, replace_text)
                self._safe_save(doc, file_path)
            return {
                "success": True,
                "file": file_path,
                "replacements": replacements,
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ==================== ADVANCED FORMATTING ====================

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
        """Apply formatting to a specific paragraph"""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
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
                    run.font.color.rgb = RGBColor(*self._hex_to_rgb(color))
                if alignment:
                    para.alignment = self._get_alignment(alignment)
                self._safe_save(doc, file_path)
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
        """Highlight all occurrences of text in a document"""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
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
                            valid_colors = {"yellow", "green", "cyan", "magenta", "blue", "red", "darkBlue", "darkCyan", "darkGreen", "darkMagenta", "darkRed", "darkYellow", "darkGray", "lightGray", "black", "white", "none"}
                            color_val = color if color and color.lower() in valid_colors else "yellow"
                            highlight.set(qn("w:val"), color_val)
                            highlights += 1
                self._safe_save(doc, file_path)
            return {
                "success": True,
                "file": file_path,
                "highlights": highlights,
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def add_table(
        self, file_path: str, headers_csv: str, data_csv: str
    ) -> Dict[str, Any]:
        """Add a formatted table to a Word document.
        headers: comma-separated. data: rows separated by ';', columns by ','."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            headers = [h.strip() for h in headers_csv.split(",")]
            rows = [
                [c.strip() for c in row.split(",")]
                for row in data_csv.split(";")
                if row.strip()
            ]
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
                doc = Document(file_path)
                table = doc.add_table(rows=1 + len(rows), cols=len(headers))
                table.style = "Table Grid"
                # Header row
                for i, h in enumerate(headers):
                    cell = table.rows[0].cells[i]
                    cell.text = h
                    for run in cell.paragraphs[0].runs:
                        run.bold = True
                # Data rows
                for ri, row in enumerate(rows, 1):
                    for ci, val in enumerate(row):
                        if ci < len(headers):
                            table.rows[ri].cells[ci].text = val
                self._safe_save(doc, file_path)
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
        self, file_path: str, comment_text: str, paragraph_index: int = 0,
        author: str = "SLOANE Agent"
    ) -> Dict[str, Any]:
        """Add a real OOXML comment annotation to a specific paragraph.
        The comment will be visible in Word/OnlyOffice as a margin annotation."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            from lxml import etree as ET
            from docx.opc.part import Part
            from docx.opc.packuri import PackURI

            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
                doc = Document(file_path)
                if paragraph_index >= len(doc.paragraphs):
                    return {
                        "success": False,
                        "error": f"Paragraph index {paragraph_index} out of range (max {len(doc.paragraphs) - 1})",
                    }

                W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                COMMENTS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
                COMMENTS_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"

                part = doc.part
                # Find existing comments part or create new one
                comments_el = None
                comments_part = None
                for rel in part.rels.values():
                    if "comments" in rel.reltype:
                        comments_part = rel.target_part
                        break

                if comments_part is not None:
                    # Parse existing comments XML
                    comments_el = ET.fromstring(comments_part.blob)
                else:
                    # Create fresh comments XML
                    comments_el = ET.Element(
                        f"{{{W_NS}}}comments",
                        nsmap={"w": W_NS, "r": R_NS}
                    )
                    blob = ET.tostring(comments_el, xml_declaration=True, encoding="UTF-8", standalone=True)
                    comments_part = Part(
                        PackURI("/word/comments.xml"), COMMENTS_CT, blob, part.package
                    )
                    part.relate_to(comments_part, COMMENTS_REL)

                # Determine next comment ID (scan ALL w:comment elements including nested)
                existing_ids = []
                for c in comments_el.iter(f"{{{W_NS}}}comment"):
                    cid = c.get(f"{{{W_NS}}}id")
                    if cid is not None:
                        try:
                            existing_ids.append(int(cid))
                        except ValueError:
                            pass
                comment_id = str(max(existing_ids, default=-1) + 1)

                # Build <w:comment> element
                now_str = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
                comment_elem = ET.SubElement(comments_el, f"{{{W_NS}}}comment")
                comment_elem.set(f"{{{W_NS}}}id", comment_id)
                comment_elem.set(f"{{{W_NS}}}author", author)
                comment_elem.set(f"{{{W_NS}}}date", now_str)
                cp = ET.SubElement(comment_elem, f"{{{W_NS}}}p")
                cr = ET.SubElement(cp, f"{{{W_NS}}}r")
                ct = ET.SubElement(cr, f"{{{W_NS}}}t")
                ct.text = comment_text

                # Serialize back to the part
                comments_part._blob = ET.tostring(
                    comments_el, xml_declaration=True, encoding="UTF-8", standalone=True
                )

                # Annotate the target paragraph with commentRangeStart/End + reference
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

                self._safe_save(doc, file_path)
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

    # ==================== APA REFERENCE MANAGEMENT ====================

    def _apa_in_text(self, author_str: str, year: str) -> str:
        """Generate APA 7th in-text citation from author string.
        Handles: single author, two authors (& separator), 3+ authors (et al.)"""
        # Detect multiple authors by common separators
        # Forms: "Smith, J., Jones, A." or "Smith, J., & Jones, A." or "Smith and Jones"
        parts = author_str.strip()
        # Split by " & " or " and " first
        if " & " in parts or " and " in parts:
            authors = [a.strip() for a in parts.replace(" and ", " & ").split(" & ")]
            surnames = [a.split(",")[0].strip() for a in authors]
            if len(surnames) == 2:
                return f"({surnames[0]} & {surnames[1]}, {year})"
            elif len(surnames) >= 3:
                return f"({surnames[0]} et al., {year})"
        # Check for comma-separated authors: "Smith, J., Jones, A."
        # Pattern: if there are 3+ commas, likely multiple authors (surname, initial, surname, initial)
        commas = parts.count(",")
        if commas >= 3:
            # Multiple authors in "Surname, I., Surname, I." format
            # Split by ", " but keep pairs together
            tokens = [t.strip() for t in parts.split(",")]
            surnames = [tokens[i] for i in range(0, len(tokens), 2) if i < len(tokens)]
            if len(surnames) == 2:
                return f"({surnames[0]} & {surnames[1]}, {year})"
            elif len(surnames) >= 3:
                return f"({surnames[0]} et al., {year})"
        # Single author
        surname = parts.split(",")[0].strip()
        return f"({surname}, {year})"

    def add_reference(
        self, file_path: str, ref_json: str
    ) -> Dict[str, Any]:
        """Add a reference to the sidecar .refs.json file for later formatting.
        ref_json is a JSON string with keys: author, year, title, source, volume, issue, pages, doi, url, type.
        type: journal|book|chapter|website|report (default: journal)"""
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

            # Deduplicate by author+year+title
            sig = (ref["author"].strip().lower(), str(ref["year"]).strip(), ref["title"].strip().lower())
            for existing in refs:
                esig = (existing["author"].strip().lower(), str(existing["year"]).strip(), existing["title"].strip().lower())
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
                "in_text_citation": self._apa_in_text(ref['author'], ref['year']),
            }
        except json.JSONDecodeError as e:
            return {"success": False, "error": f"Invalid JSON: {e}"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def _format_apa7_reference(self, ref: dict):
        """Format a single reference in APA 7th edition style.
        Returns a list of {"text": str, "italic": bool} spans for proper Word rendering."""
        rtype = ref.get("type", "journal")
        author = ref.get("author", "Unknown")
        year = ref.get("year", "n.d.")
        title = ref.get("title", "Untitled")
        S = lambda t, i=False: {"text": t, "italic": i}  # span helper

        if rtype == "journal":
            source = ref.get("source", "")
            vol = ref.get("volume", "")
            issue = ref.get("issue", "")
            pages = ref.get("pages", "")
            doi = ref.get("doi", "")
            spans = [S(f"{author} ({year}). {title}. ")]
            if source:
                spans.append(S(source, True))  # journal name italic
                if vol:
                    spans.append(S(", "))
                    spans.append(S(vol, True))  # volume number italic
                    if issue:
                        spans.append(S(f"({issue})"))
                elif issue:
                    spans.append(S(f"({issue})"))
                if pages:
                    spans.append(S(f", {pages}"))
                spans.append(S(". "))
            if doi:
                url = f"https://doi.org/{doi}" if not doi.startswith("http") else doi
                spans.append(S(url))
            return spans

        elif rtype == "book":
            publisher = ref.get("source", ref.get("publisher", ""))
            doi = ref.get("doi", "")
            spans = [S(f"{author} ({year}). "), S(title, True), S(". ")]
            if publisher:
                spans.append(S(f"{publisher}. "))
            if doi:
                url = f"https://doi.org/{doi}" if not doi.startswith("http") else doi
                spans.append(S(url))
            return spans

        elif rtype == "chapter":
            editor = ref.get("editor", "")
            book_title = ref.get("source", "")
            pages = ref.get("pages", "")
            publisher = ref.get("publisher", "")
            spans = [S(f"{author} ({year}). {title}. ")]
            if editor:
                spans.append(S(f"In {editor} (Ed.), "))
            if book_title:
                spans.append(S(book_title, True))  # book title italic
            if pages:
                spans.append(S(f" (pp. {pages})"))
            spans.append(S(". "))
            if publisher:
                spans.append(S(f"{publisher}."))
            return spans

        elif rtype == "website":
            url = ref.get("url", ref.get("doi", ""))
            source = ref.get("source", "")
            spans = [S(f"{author} ({year}). "), S(title, True), S(". ")]
            if source:
                spans.append(S(f"{source}. "))
            if url:
                spans.append(S(url))
            return spans

        elif rtype == "report":
            source = ref.get("source", "")
            url = ref.get("url", "")
            spans = [S(f"{author} ({year}). "), S(title, True), S(". ")]
            if source:
                spans.append(S(f"{source}. "))
            if url:
                spans.append(S(url))
            return spans

        else:
            # Fallback: generic
            return [S(f"{author} ({year}). {title}.")]

    def build_references(self, file_path: str) -> Dict[str, Any]:
        """Read the sidecar .refs.json, format all references in APA 7th edition,
        and append a 'References' section to the Word document."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        refs_path = file_path + ".refs.json"
        if not os.path.exists(refs_path):
            return {"success": False, "error": f"No references file found at {refs_path}. Use doc-add-reference first."}
        try:
            with open(refs_path, "r") as f:
                refs = json.load(f)
            if not refs:
                return {"success": False, "error": "References file is empty"}

            # Sort alphabetically by author surname (APA requirement)
            refs.sort(key=lambda r: r.get("author", "").split(",")[0].strip().lower())

            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
                doc = Document(file_path)

                # Remove existing References section if present (idempotent)
                ref_start_idx = None
                for i, para in enumerate(doc.paragraphs):
                    if para.text.strip() == "References" and any(r.bold for r in para.runs if r.bold):
                        ref_start_idx = i
                        break
                if ref_start_idx is not None:
                    # Delete from the References heading to end of document
                    body = doc.element.body
                    paras_to_remove = list(doc.paragraphs[ref_start_idx:])
                    for p in paras_to_remove:
                        body.remove(p._element)
                    # Also remove the page break before it (last element before References)
                    # It's typically a paragraph with a <w:br w:type="page"/> run
                    if ref_start_idx > 0:
                        prev = doc.paragraphs[ref_start_idx - 1] if ref_start_idx - 1 < len(doc.paragraphs) else None
                        if prev and not prev.text.strip():
                            # Check if it's a page break paragraph
                            for run_el in prev._element.findall(qn("w:r")):
                                for br in run_el.findall(qn("w:br")):
                                    if br.get(qn("w:type")) == "page":
                                        body.remove(prev._element)
                                        break

                # Add a page break before references
                doc.add_page_break()

                # "References" heading — centered, bold
                heading = doc.add_paragraph()
                heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = heading.add_run("References")
                run.bold = True

                # Each reference as a hanging-indent paragraph with italic spans
                from docx.enum.text import WD_LINE_SPACING
                for ref in refs:
                    spans = self._format_apa7_reference(ref)
                    para = doc.add_paragraph()
                    for span in spans:
                        run = para.add_run(span["text"])
                        if span.get("italic"):
                            run.italic = True
                    # APA hanging indent: first line 0, rest 0.5 inches
                    pf = para.paragraph_format
                    pf.first_line_indent = -Inches(0.5)
                    pf.left_indent = Inches(0.5)
                    pf.space_after = Pt(0)
                    pf.space_before = Pt(0)
                    pf.line_spacing_rule = WD_LINE_SPACING.DOUBLE

                self._safe_save(doc, file_path)

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
        self, file_path: str, image_path: str,
        width_inches: float = 5.5, caption: str = None
    ) -> Dict[str, Any]:
        """Add an image to a Word document with optional caption.
        Supports PNG, JPG, GIF, BMP, TIFF."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        if not os.path.exists(image_path):
            return {"success": False, "error": f"Image not found: {image_path}"}
        try:
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
                doc = Document(file_path)
                # Add image paragraph, centered
                para = doc.add_paragraph()
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = para.add_run()
                run.add_picture(image_path, width=Inches(width_inches))
                # Add caption below image if provided
                if caption:
                    cap_para = doc.add_paragraph()
                    cap_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    cap_run = cap_para.add_run(caption)
                    cap_run.italic = True
                    cap_run.font.size = Pt(10)
                self._safe_save(doc, file_path)
            return {
                "success": True,
                "file": file_path,
                "image": image_path,
                "width_inches": width_inches,
                "caption": caption,
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def set_paragraph_style(
        self, file_path: str, paragraph_index: int, style_name: str
    ) -> Dict[str, Any]:
        """Set a paragraph's Word style (Heading 1, Heading 2, Normal, Title, etc.)"""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
                doc = Document(file_path)
                if paragraph_index >= len(doc.paragraphs):
                    return {
                        "success": False,
                        "error": f"Paragraph index {paragraph_index} out of range (max {len(doc.paragraphs) - 1})",
                    }
                para = doc.paragraphs[paragraph_index]
                available_styles = [s.name for s in doc.styles if s.type == 1]  # paragraph styles
                if style_name not in available_styles:
                    return {
                        "success": False,
                        "error": f"Style '{style_name}' not found. Available: {available_styles[:20]}",
                    }
                para.style = doc.styles[style_name]
                self._safe_save(doc, file_path)
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
        orientation: str = "portrait",
        margins: Dict[str, float] = None,
        header_text: str = None,
        page_numbers: bool = False,
    ) -> Dict[str, Any]:
        """Set page layout (orientation, margins, running header, page numbers).
        header_text: appears top-left on every page.
        page_numbers: adds page number to bottom-center."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
                doc = Document(file_path)
                section = doc.sections[0]
                if orientation.lower() == "landscape":
                    section.orientation = WD_ORIENT.LANDSCAPE
                    # python-docx does NOT auto-swap dimensions — must do it explicitly
                    if section.page_width < section.page_height:
                        section.page_width, section.page_height = (
                            section.page_height, section.page_width
                        )
                else:
                    section.orientation = WD_ORIENT.PORTRAIT
                    # Ensure portrait has height > width
                    if section.page_width > section.page_height:
                        section.page_width, section.page_height = (
                            section.page_height, section.page_width
                        )
                if margins:
                    for side, value in margins.items():
                        if hasattr(section, f"{side}_margin"):
                            setattr(section, f"{side}_margin", Inches(value))

                # Running header
                if header_text is not None:
                    header = section.header
                    header.is_linked_to_previous = False
                    if header.paragraphs:
                        hp = header.paragraphs[0]
                    else:
                        hp = header.add_paragraph()
                    hp.text = ""
                    run = hp.add_run(header_text)
                    run.font.size = Pt(10)
                    hp.alignment = WD_ALIGN_PARAGRAPH.LEFT

                # Page numbers (bottom center)
                if page_numbers:
                    footer = section.footer
                    footer.is_linked_to_previous = False
                    if footer.paragraphs:
                        fp = footer.paragraphs[0]
                    else:
                        fp = footer.add_paragraph()
                    fp.text = ""
                    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    # Insert PAGE field code via OOXML
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

                self._safe_save(doc, file_path)
            return {
                "success": True,
                "file": file_path,
                "layout_updated": True,
                "header_text": header_text,
                "page_numbers": page_numbers,
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def _parse_range(self, range_str: str):
        from openpyxl.utils.cell import range_boundaries

        if not range_str or ":" not in range_str:
            raise ValueError(
                f"Invalid range '{range_str}'. Expected format like A1:C10"
            )
        return range_boundaries(range_str)

    def _range_has_data(self, ws, range_str: str) -> bool:
        min_col, min_row, max_col, max_row = self._parse_range(range_str)
        for row in ws.iter_rows(
            min_row=min_row,
            max_row=max_row,
            min_col=min_col,
            max_col=max_col,
            values_only=True,
        ):
            for cell in row:
                if cell is not None and str(cell).strip() != "":
                    return True
        return False

    def _resolve_formula_value(
        self, ws, expr: str, row_hint: int = None, _depth: int = 0
    ):
        if _depth > 5:
            return None

        def cell_numeric(cell_ref: str):
            val = ws[cell_ref].value
            if isinstance(val, (int, float)):
                return float(val)
            if isinstance(val, str):
                if val.startswith("="):
                    return self._resolve_formula_value(
                        ws, val[1:], row_hint=row_hint, _depth=_depth + 1
                    )
                try:
                    return float(val)
                except Exception:
                    return None
            return None

        fn_pat = re.compile(
            r"\b(SUM|AVERAGE|AVG|MIN|MAX)\(([A-Z]+\d+:[A-Z]+\d+)\)",
            re.IGNORECASE,
        )

        def fn_repl(match):
            fn = match.group(1).upper()
            rng = match.group(2)
            min_col, min_row, max_col, max_row = self._parse_range(rng)
            vals = []
            for row in ws.iter_rows(
                min_row=min_row,
                max_row=max_row,
                min_col=min_col,
                max_col=max_col,
                values_only=False,
            ):
                for c in row:
                    v = c.value
                    if isinstance(v, (int, float)):
                        vals.append(float(v))
                    elif isinstance(v, str):
                        if v.startswith("="):
                            fv = self._resolve_formula_value(
                                ws, v[1:], row_hint=c.row, _depth=_depth + 1
                            )
                            if fv is not None:
                                vals.append(float(fv))
                        else:
                            try:
                                vals.append(float(v))
                            except Exception:
                                pass
            if not vals:
                return "0"
            if fn == "SUM":
                return str(sum(vals))
            if fn in ("AVERAGE", "AVG"):
                return str(sum(vals) / len(vals))
            if fn == "MIN":
                return str(min(vals))
            if fn == "MAX":
                return str(max(vals))
            return "0"

        expr = fn_pat.sub(fn_repl, expr)
        cell_pat = re.compile(r"\b([A-Z]+\d+)\b")
        expr = cell_pat.sub(lambda m: str(cell_numeric(m.group(1)) or 0), expr)

        # Safety guard: arithmetic-only after substitutions
        if not re.fullmatch(r"[0-9\s\+\-\*/\(\)\.]+", expr):
            return None

        try:
            return float(eval(expr, {"__builtins__": {}}, {}))
        except Exception:
            return None

    def _extract_formula_functions(self, formula: str):
        if not formula:
            return set()
        expr = formula[1:] if formula.startswith("=") else formula
        return set(re.findall(r"\b([A-Z][A-Z0-9_]*)\s*\(", expr.upper()))

    def _formula_depth(self, formula: str) -> int:
        expr = (
            formula[1:]
            if isinstance(formula, str) and formula.startswith("=")
            else str(formula or "")
        )
        depth = 0
        max_depth = 0
        for ch in expr:
            if ch == "(":
                depth += 1
                max_depth = max(max_depth, depth)
            elif ch == ")":
                depth = max(0, depth - 1)
        return max_depth

    def audit_spreadsheet_formulas(
        self,
        file_path: str,
        sheet_name: str = None,
        max_examples: int = 30,
    ) -> Dict[str, Any]:
        """Audit workbook formulas for unsupported patterns and risk signals."""
        if not OPENPYXL_AVAILABLE:
            return {"success": False, "error": "openpyxl not installed"}
        try:
            wb = load_workbook(file_path, data_only=False)
            sheets = (
                [sheet_name]
                if sheet_name and sheet_name in wb.sheetnames
                else wb.sheetnames
            )

            has_vba = getattr(wb, "vba_archive", None) is not None
            external_link_count = len(getattr(wb, "_external_links", []) or [])

            formula_count = 0
            unsupported_functions = {}
            complex_formulas = []
            external_ref_formulas = []
            function_usage = {}

            for sn in sheets:
                ws = wb[sn]
                for row in ws.iter_rows(
                    min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column
                ):
                    for cell in row:
                        val = cell.value
                        if not (isinstance(val, str) and val.startswith("=")):
                            continue
                        formula_count += 1
                        funcs = self._extract_formula_functions(val)
                        for f in funcs:
                            function_usage[f] = function_usage.get(f, 0) + 1
                            if f not in self.supported_formula_functions:
                                unsupported_functions[f] = (
                                    unsupported_functions.get(f, 0) + 1
                                )
                        depth = self._formula_depth(val)
                        if depth > 3 and len(complex_formulas) < max_examples:
                            complex_formulas.append(
                                {
                                    "sheet": sn,
                                    "cell": cell.coordinate,
                                    "depth": depth,
                                    "formula": val[:220],
                                }
                            )
                        if (
                            "[" in val
                            or "http://" in val.lower()
                            or "https://" in val.lower()
                        ) and len(external_ref_formulas) < max_examples:
                            external_ref_formulas.append(
                                {
                                    "sheet": sn,
                                    "cell": cell.coordinate,
                                    "formula": val[:220],
                                }
                            )

            wb.close()

            risks = []
            if has_vba:
                risks.append(
                    "Workbook contains VBA/macros; CLI does not execute macros"
                )
            if external_link_count > 0 or external_ref_formulas:
                risks.append(
                    "Workbook contains external links/references that may not resolve in CLI"
                )
            if unsupported_functions:
                risks.append(
                    "Workbook uses formula functions outside CLI evaluator support"
                )
            if complex_formulas:
                risks.append(
                    "Workbook contains deeply nested formulas that can be error-prone"
                )

            safe = not risks
            return {
                "success": True,
                "file": file_path,
                "sheet_scope": sheets,
                "formula_count": formula_count,
                "has_vba": has_vba,
                "external_link_count": external_link_count,
                "function_usage": function_usage,
                "unsupported_functions": unsupported_functions,
                "complex_formula_examples": complex_formulas,
                "external_reference_examples": external_ref_formulas,
                "safe_for_cli_formula_eval": safe,
                "risk_level": "low"
                if safe
                else ("high" if has_vba or unsupported_functions else "medium"),
                "risks": risks,
                "supported_functions": sorted(self.supported_formula_functions),
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_formatting_info(self, file_path: str) -> Dict[str, Any]:
        """Get detailed formatting information"""
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
            for i, section in enumerate(doc.sections):
                info["sections"].append(
                    {
                        "orientation": "landscape"
                        if section.orientation == WD_ORIENT.LANDSCAPE
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

    def get_document_info(self, file_path: str) -> Dict[str, Any]:
        """Get file information for any supported format"""
        try:
            from pathlib import Path
            import os

            path = Path(file_path)
            if not path.exists():
                return {"success": False, "error": f"File not found: {file_path}"}

            stat = os.stat(file_path)
            ext = path.suffix.lower()

            info = {
                "success": True,
                "file": file_path,
                "name": path.name,
                "size": stat.st_size,
                "modified": datetime.fromtimestamp(stat.st_mtime).isoformat(),
                "created": datetime.fromtimestamp(stat.st_ctime).isoformat(),
                "extension": ext,
            }

            if ext == ".docx" and DOCX_AVAILABLE:
                doc = Document(file_path)
                info["type"] = "document"
                info["paragraph_count"] = len(doc.paragraphs)
                info["word_count"] = sum(
                    len(para.text.split()) for para in doc.paragraphs
                )
            elif ext == ".xlsx" and OPENPYXL_AVAILABLE:
                wb = load_workbook(file_path, read_only=True)
                info["type"] = "spreadsheet"
                info["sheets"] = wb.sheetnames
                info["sheet_count"] = len(wb.sheetnames)
                wb.close()
            elif ext == ".pptx" and PPTX_AVAILABLE:
                prs = Presentation(file_path)
                info["type"] = "presentation"
                info["slide_count"] = len(prs.slides)

            return info
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ==================== SPREADSHEET OPERATIONS ====================

    def create_spreadsheet(
        self, output_path: str, sheet_name: str = "Sheet1"
    ) -> Dict[str, Any]:
        """Create a new .xlsx spreadsheet"""
        if not OPENPYXL_AVAILABLE:
            return {"success": False, "error": "openpyxl not installed"}
        try:
            with self._file_lock(output_path):
                backup = self._snapshot_backup(output_path)
                wb = Workbook()
                ws = wb.active
                ws.title = sheet_name
                ws.page_setup.paperSize = 9  # A4
                self._safe_save(wb, output_path)
            return {
                "success": True,
                "file": output_path,
                "sheets": [sheet_name],
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def write_spreadsheet(
        self,
        output_path: str,
        headers: List[str],
        data: List[List[Any]],
        sheet_name: str = "Sheet1",
        overwrite_workbook: bool = False,
        coerce_rows: bool = False,
        text_columns: List[str] = None,
    ) -> Dict[str, Any]:
        """Write headers/data to spreadsheet. Non-destructive by default; full overwrite only when explicitly requested.
        text_columns: list of header names whose values should NOT be coerced to numbers (preserves leading zeros)."""
        if not OPENPYXL_AVAILABLE:
            return {"success": False, "error": "openpyxl not installed"}
        try:
            ok, err, data = self._validate_tabular_rows(
                headers, data, coerce_rows=coerce_rows
            )
            if not ok:
                return {"success": False, "error": err}

            with self._file_lock(output_path):
                backup = self._snapshot_backup(output_path)
                if Path(output_path).exists() and not overwrite_workbook:
                    wb = load_workbook(output_path)
                    ws = (
                        wb[sheet_name]
                        if sheet_name in wb.sheetnames
                        else wb.create_sheet(sheet_name)
                    )
                    ws.delete_rows(1, ws.max_row)
                else:
                    wb = Workbook()
                    ws = wb.active
                    ws.title = sheet_name

                for col, header in enumerate(headers, 1):
                    ws.cell(row=1, column=col, value=header)
                # Build set of column indices that should stay as text
                text_col_indices = set()
                if text_columns:
                    tc_lower = {tc.strip().lower() for tc in text_columns}
                    for ci, h in enumerate(headers):
                        if h.strip().lower() in tc_lower:
                            text_col_indices.add(ci + 1)  # 1-indexed

                for row_idx, row_data in enumerate(data, 2):
                    for col_idx, value in enumerate(row_data, 1):
                        if col_idx in text_col_indices:
                            # Text column: write as string, preserve leading zeros
                            cell = ws.cell(row=row_idx, column=col_idx, value=str(value))
                            cell.data_type = "s"
                        elif isinstance(value, str) and value.startswith("="):
                            ws.cell(row=row_idx, column=col_idx, value=value)
                        else:
                            try:
                                if isinstance(value, str):
                                    if "." in value:
                                        value = float(value)
                                    else:
                                        value = int(value)
                            except (ValueError, TypeError):
                                pass
                            ws.cell(row=row_idx, column=col_idx, value=value)
                # Auto-fit column widths based on content (min 12, max 50 chars)
                for col in ws.columns:
                    max_len = max((len(str(cell.value or "")) for cell in col), default=8)
                    col_letter = openpyxl.utils.get_column_letter(col[0].column)
                    ws.column_dimensions[col_letter].width = max(12, min(max_len + 2, 50))
                # A4 paper size for printing
                ws.page_setup.paperSize = 9  # 9 = A4
                self._safe_save(wb, output_path)
            return {
                "success": True,
                "file": output_path,
                "rows_written": len(data),
                "columns": len(headers),
                "sheet": sheet_name,
                "mode": "overwrite_workbook" if overwrite_workbook else "update_sheet",
                "coerce_rows": bool(coerce_rows),
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def read_spreadsheet(
        self, file_path: str, sheet_name: str = None
    ) -> Dict[str, Any]:
        """Read all data from a spreadsheet"""
        if not OPENPYXL_AVAILABLE:
            return {"success": False, "error": "openpyxl not installed"}
        try:
            wb = load_workbook(file_path, read_only=True)
            if sheet_name and sheet_name not in wb.sheetnames:
                wb.close()
                return {
                    "success": False,
                    "error": f"Sheet '{sheet_name}' not found. Available: {list(wb.sheetnames)}",
                }
            sheets_to_read = [sheet_name] if sheet_name else list(wb.sheetnames)
            result = {
                "success": True,
                "file": file_path,
                "sheets": wb.sheetnames,
                "data": {},
            }
            for sn in sheets_to_read:
                ws = wb[sn]
                rows_data = []
                headers = None
                for row in ws.iter_rows(values_only=True):
                    if all(cell is None for cell in row):
                        continue
                    if headers is None:
                        headers = [str(cell) if cell else "" for cell in row]
                    else:
                        rows_data.append([str(cell) if cell else "" for cell in row])
                result["data"][sn] = {
                    "headers": headers,
                    "rows": rows_data,
                    "row_count": len(rows_data),
                }
            wb.close()
            return result
        except Exception as e:
            return {"success": False, "error": str(e)}

    def calculate_column(
        self,
        file_path: str,
        column_letter: str,
        operation: str,
        sheet_name: str = "Sheet1",
        include_formulas: bool = False,
        strict_formula_safety: bool = False,
    ) -> Dict[str, Any]:
        """Calculate column statistics (sum, avg, min, max)"""
        if not OPENPYXL_AVAILABLE:
            return {"success": False, "error": "openpyxl not installed"}
        try:
            # Must use read_only=False for column access
            wb = load_workbook(file_path)
            ws = self._get_sheet(wb, sheet_name)

            # Get column index from letter
            col_idx = openpyxl.utils.column_index_from_string(column_letter.upper())

            # Extract numeric values from the column (skip header row)
            values = []
            formula_rows_total = 0
            formula_rows_evaluated = 0
            formula_rows_failed = 0
            formula_failure_examples = []
            for row in ws.iter_rows(min_row=2, values_only=True):
                if col_idx <= len(row):
                    val = row[col_idx - 1]
                    if val is not None and isinstance(val, (int, float)):
                        values.append(val)
                    elif (
                        include_formulas
                        and isinstance(val, str)
                        and val.startswith("=")
                    ):
                        formula_rows_total += 1
                        funcs = self._extract_formula_functions(val)
                        unsupported = sorted(
                            [
                                f
                                for f in funcs
                                if f not in self.supported_formula_functions
                            ]
                        )
                        if unsupported:
                            formula_rows_failed += 1
                            if len(formula_failure_examples) < 8:
                                formula_failure_examples.append(
                                    {
                                        "formula": val[:220],
                                        "reason": f"unsupported_functions:{','.join(unsupported)}",
                                    }
                                )
                            continue
                        resolved = self._resolve_formula_value(ws, val[1:])
                        if isinstance(resolved, (int, float)):
                            values.append(float(resolved))
                            formula_rows_evaluated += 1
                        else:
                            formula_rows_failed += 1
                            if len(formula_failure_examples) < 8:
                                formula_failure_examples.append(
                                    {
                                        "formula": val[:220],
                                        "reason": "evaluation_failed",
                                    }
                                )

            wb.close()

            if not values:
                return {
                    "success": False,
                    "error": f"No numeric values in column {column_letter}",
                }

            op_name = (operation or "all").lower()
            if op_name not in {"sum", "avg", "average", "min", "max", "all"}:
                return {
                    "success": False,
                    "error": f"Unsupported operation '{operation}'. Use sum|avg|min|max|all",
                }

            result = {
                "success": True,
                "column": column_letter,
                "sheet": sheet_name,
                "count": len(values),
                "sum": sum(values),
                "average": sum(values) / len(values),
                "min": min(values),
                "max": max(values),
                "formula_mode": bool(include_formulas),
                "operation": op_name,
                "formula_rows_total": formula_rows_total,
                "formula_rows_evaluated": formula_rows_evaluated,
                "formula_rows_failed": formula_rows_failed,
                "formula_failure_examples": formula_failure_examples,
            }
            if op_name == "sum":
                result["value"] = result["sum"]
            elif op_name in {"avg", "average"}:
                result["value"] = result["average"]
            elif op_name == "min":
                result["value"] = result["min"]
            elif op_name == "max":
                result["value"] = result["max"]

            if include_formulas:
                formula_eval_rate = (
                    (formula_rows_evaluated / formula_rows_total)
                    if formula_rows_total > 0
                    else 1.0
                )
                result["formula_eval_rate"] = formula_eval_rate
                result["formula_reliability"] = (
                    "high"
                    if formula_eval_rate >= 0.95
                    else ("medium" if formula_eval_rate >= 0.7 else "low")
                )
                if strict_formula_safety and formula_rows_failed > 0:
                    return {
                        "success": False,
                        "error": "Strict formula safety failed: unresolved/unsupported formulas present",
                        "details": {
                            "formula_rows_total": formula_rows_total,
                            "formula_rows_evaluated": formula_rows_evaluated,
                            "formula_rows_failed": formula_rows_failed,
                            "examples": formula_failure_examples,
                        },
                    }
            return result
        except Exception as e:
            return {"success": False, "error": str(e)}

    def _cell_to_float(self, value):
        if value is None:
            return None
        if isinstance(value, bool):
            return None
        if isinstance(value, (int, float)):
            return float(value)
        if isinstance(value, str):
            text = value.strip()
            if not text or text.startswith("="):
                return None
            try:
                return float(text)
            except Exception:
                return None
        return None

    def _value_matches_group(self, value, target) -> bool:
        if value is None:
            return False
        t = str(target).strip()
        v = str(value).strip()

        try:
            return float(v) == float(t)
        except Exception:
            return v.lower() == t.lower()

    def _category_value(self, value):
        if value is None:
            return None
        if isinstance(value, str):
            txt = value.strip()
            if not txt:
                return None
            try:
                n = float(txt)
                return int(n) if n.is_integer() else n
            except Exception:
                return txt
        if isinstance(value, float) and value.is_integer():
            return int(value)
        return value

    def frequencies(
        self,
        file_path: str,
        column_letter: str,
        sheet_name: str = "Sheet1",
        allowed_values: List[str] = None,
    ) -> Dict[str, Any]:
        """Frequency table for one column (counts + percentages)."""
        if not OPENPYXL_AVAILABLE:
            return {"success": False, "error": "openpyxl not installed"}
        try:
            wb = load_workbook(file_path, read_only=True)
            ws = self._get_sheet(wb, sheet_name)
            col_idx = openpyxl.utils.column_index_from_string(column_letter.upper())

            counts = {}
            missing = 0
            excluded = 0

            def _allowed(cat):
                if not allowed_values:
                    return True
                return any(self._value_matches_group(cat, t) for t in allowed_values)

            for row in ws.iter_rows(min_row=2, values_only=True):
                if col_idx > len(row):
                    missing += 1
                    continue
                cat = self._category_value(row[col_idx - 1])
                if cat is None:
                    missing += 1
                    continue
                if not _allowed(cat):
                    excluded += 1
                    continue
                counts[cat] = counts.get(cat, 0) + 1

            wb.close()

            total_valid = sum(counts.values())
            total_rows = total_valid + missing + excluded
            freq_rows = []
            for k in sorted(counts.keys(), key=lambda x: (str(type(x)), str(x))):
                c = counts[k]
                freq_rows.append(
                    {
                        "category": k,
                        "count": c,
                        "percent_valid": (c / total_valid * 100)
                        if total_valid
                        else 0.0,
                        "percent_total": (c / total_rows * 100) if total_rows else 0.0,
                    }
                )

            return {
                "success": True,
                "file": file_path,
                "sheet": sheet_name,
                "column": column_letter.upper(),
                "valid_n": total_valid,
                "missing_n": missing,
                "excluded_n": excluded,
                "total_n": total_rows,
                "allowed_values": allowed_values or [],
                "frequencies": freq_rows,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def correlation_test(
        self,
        file_path: str,
        x_column: str,
        y_column: str,
        sheet_name: str = "Sheet1",
        method: str = "pearson",
    ) -> Dict[str, Any]:
        """Correlation between two numeric columns."""
        if not OPENPYXL_AVAILABLE:
            return {"success": False, "error": "openpyxl not installed"}
        if not SCIPY_AVAILABLE:
            return {"success": False, "error": "scipy not installed"}
        try:
            wb = load_workbook(file_path, read_only=True)
            ws = self._get_sheet(wb, sheet_name)
            x_idx = openpyxl.utils.column_index_from_string(x_column.upper())
            y_idx = openpyxl.utils.column_index_from_string(y_column.upper())

            x_vals, y_vals = [], []
            dropped = 0
            for row in ws.iter_rows(min_row=2, values_only=True):
                xv = row[x_idx - 1] if x_idx <= len(row) else None
                yv = row[y_idx - 1] if y_idx <= len(row) else None
                x_num = self._cell_to_float(xv)
                y_num = self._cell_to_float(yv)
                if x_num is None or y_num is None:
                    dropped += 1
                    continue
                x_vals.append(x_num)
                y_vals.append(y_num)
            wb.close()

            if len(x_vals) < 3:
                return {
                    "success": False,
                    "error": "Need at least 3 paired numeric observations",
                }

            m = (method or "pearson").lower()
            if m == "pearson":
                stat, p = scipy_stats.pearsonr(x_vals, y_vals)
            elif m == "spearman":
                stat, p = scipy_stats.spearmanr(x_vals, y_vals)
            else:
                return {
                    "success": False,
                    "error": "Unsupported method. Use pearson|spearman",
                }

            abs_r = abs(float(stat))
            strength = (
                "negligible"
                if abs_r < 0.1
                else (
                    "small" if abs_r < 0.3 else ("moderate" if abs_r < 0.5 else "large")
                )
            )
            significant = float(p) < 0.05
            direction = (
                "positive"
                if float(stat) > 0
                else ("negative" if float(stat) < 0 else "none")
            )

            return {
                "success": True,
                "file": file_path,
                "sheet": sheet_name,
                "method": m,
                "x_column": x_column.upper(),
                "y_column": y_column.upper(),
                "n": len(x_vals),
                "dropped_rows": dropped,
                "statistic": float(stat),
                "p_value": float(p),
                "interpretation": {
                    "alpha": 0.05,
                    "significant": significant,
                    "direction": direction,
                    "strength": strength,
                },
                "apa": f"{m.title()} correlation between {x_column.upper()} and {y_column.upper()} was {'ρ' if m == 'spearman' else 'r'} = {float(stat):.3f}, p = {float(p):.4g}, n = {len(x_vals)}.",
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def ttest_independent(
        self,
        file_path: str,
        value_column: str,
        group_column: str,
        group_a: str,
        group_b: str,
        sheet_name: str = "Sheet1",
        equal_var: bool = False,
    ) -> Dict[str, Any]:
        """Independent samples t-test (Welch default)."""
        if not OPENPYXL_AVAILABLE:
            return {"success": False, "error": "openpyxl not installed"}
        if not SCIPY_AVAILABLE:
            return {"success": False, "error": "scipy not installed"}
        try:
            wb = load_workbook(file_path, read_only=True)
            ws = self._get_sheet(wb, sheet_name)
            v_idx = openpyxl.utils.column_index_from_string(value_column.upper())
            g_idx = openpyxl.utils.column_index_from_string(group_column.upper())

            group_a_vals = []
            group_b_vals = []
            dropped = 0
            for row in ws.iter_rows(min_row=2, values_only=True):
                gv = row[g_idx - 1] if g_idx <= len(row) else None
                vv = row[v_idx - 1] if v_idx <= len(row) else None
                v_num = self._cell_to_float(vv)
                if v_num is None:
                    dropped += 1
                    continue
                if self._value_matches_group(gv, group_a):
                    group_a_vals.append(v_num)
                elif self._value_matches_group(gv, group_b):
                    group_b_vals.append(v_num)
            wb.close()

            if len(group_a_vals) < 2 or len(group_b_vals) < 2:
                return {
                    "success": False,
                    "error": "Need at least 2 numeric observations in each group",
                }

            t_stat, p_val = scipy_stats.ttest_ind(
                group_a_vals,
                group_b_vals,
                equal_var=bool(equal_var),
                nan_policy="omit",
            )

            mean_a = float(sum(group_a_vals) / len(group_a_vals))
            mean_b = float(sum(group_b_vals) / len(group_b_vals))
            sd_a = float(scipy_stats.tstd(group_a_vals))
            sd_b = float(scipy_stats.tstd(group_b_vals))

            # Degrees of freedom
            n_a, n_b = len(group_a_vals), len(group_b_vals)
            if equal_var:
                df = n_a + n_b - 2
            else:
                # Welch-Satterthwaite df
                va, vb = sd_a**2 / n_a, sd_b**2 / n_b
                df = (va + vb)**2 / (va**2 / (n_a - 1) + vb**2 / (n_b - 1)) if (va + vb) > 0 else n_a + n_b - 2

            # Cohen's d (pooled SD)
            pooled_var = (
                ((n_a - 1) * (sd_a**2))
                + ((n_b - 1) * (sd_b**2))
            ) / (n_a + n_b - 2)
            pooled_sd = pooled_var**0.5 if pooled_var > 0 else 0.0
            cohens_d = (mean_a - mean_b) / pooled_sd if pooled_sd else 0.0

            abs_d = abs(float(cohens_d))
            d_mag = (
                "negligible"
                if abs_d < 0.2
                else (
                    "small" if abs_d < 0.5 else ("medium" if abs_d < 0.8 else "large")
                )
            )
            significant = float(p_val) < 0.05

            return {
                "success": True,
                "file": file_path,
                "sheet": sheet_name,
                "value_column": value_column.upper(),
                "group_column": group_column.upper(),
                "group_a": str(group_a),
                "group_b": str(group_b),
                "n_a": n_a,
                "n_b": n_b,
                "mean_a": mean_a,
                "mean_b": mean_b,
                "sd_a": sd_a,
                "sd_b": sd_b,
                "df": round(float(df), 2),
                "difference": mean_a - mean_b,
                "equal_var": bool(equal_var),
                "statistic": float(t_stat),
                "p_value": float(p_val),
                "cohens_d": float(cohens_d),
                "dropped_rows": dropped,
                "interpretation": {
                    "alpha": 0.05,
                    "significant": significant,
                    "effect_size_magnitude": d_mag,
                    "higher_group": str(group_a)
                    if mean_a > mean_b
                    else (str(group_b) if mean_b > mean_a else "equal"),
                },
                "apa": f"{'Welch' if not equal_var else 'Student'}'s independent-samples t-test on {value_column.upper()} by {group_column.upper()} ({group_a} vs {group_b}) found t({df:.1f}) = {float(t_stat):.3f}, p = {float(p_val):.4g}, d = {float(cohens_d):.3f}; M{group_a} = {mean_a:.3f}, M{group_b} = {mean_b:.3f}.",
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def mann_whitney_test(
        self,
        file_path: str,
        value_column: str,
        group_column: str,
        group_a: str,
        group_b: str,
        sheet_name: str = "Sheet1",
    ) -> Dict[str, Any]:
        """Mann-Whitney U test — non-parametric alternative to independent t-test.
        Appropriate for ordinal (Likert) data."""
        if not OPENPYXL_AVAILABLE:
            return {"success": False, "error": "openpyxl not installed"}
        if not SCIPY_AVAILABLE:
            return {"success": False, "error": "scipy not installed"}
        try:
            wb = load_workbook(file_path, read_only=True)
            ws = self._get_sheet(wb, sheet_name)
            val_idx = openpyxl.utils.column_index_from_string(value_column.upper())
            grp_idx = openpyxl.utils.column_index_from_string(group_column.upper())
            group_a_vals, group_b_vals = [], []
            dropped = 0
            for row in ws.iter_rows(min_row=2, values_only=True):
                val = row[val_idx - 1] if val_idx - 1 < len(row) else None
                grp = row[grp_idx - 1] if grp_idx - 1 < len(row) else None
                if val is None or grp is None:
                    dropped += 1
                    continue
                try:
                    v = float(val)
                except (ValueError, TypeError):
                    dropped += 1
                    continue
                g = str(grp).strip()
                if g == str(group_a).strip():
                    group_a_vals.append(v)
                elif g == str(group_b).strip():
                    group_b_vals.append(v)
            wb.close()
            if len(group_a_vals) < 2 or len(group_b_vals) < 2:
                return {"success": False, "error": f"Insufficient data: group {group_a} n={len(group_a_vals)}, group {group_b} n={len(group_b_vals)}. Need >= 2 per group."}
            u_stat, p_val = scipy_stats.mannwhitneyu(
                group_a_vals, group_b_vals, alternative="two-sided"
            )
            # Rank-biserial r as effect size
            n1, n2 = len(group_a_vals), len(group_b_vals)
            r_rb = 1 - (2 * float(u_stat)) / (n1 * n2)
            abs_r = abs(r_rb)
            r_mag = "negligible" if abs_r < 0.1 else ("small" if abs_r < 0.3 else ("medium" if abs_r < 0.5 else "large"))
            significant = float(p_val) < 0.05
            import numpy as np
            median_a = float(np.median(group_a_vals))
            median_b = float(np.median(group_b_vals))
            return {
                "success": True,
                "file": file_path,
                "sheet": sheet_name,
                "test": "Mann-Whitney U",
                "value_column": value_column.upper(),
                "group_column": group_column.upper(),
                "group_a": str(group_a),
                "group_b": str(group_b),
                "n_a": n1,
                "n_b": n2,
                "median_a": median_a,
                "median_b": median_b,
                "statistic": float(u_stat),
                "p_value": float(p_val),
                "rank_biserial_r": float(r_rb),
                "dropped_rows": dropped,
                "interpretation": {
                    "alpha": 0.05,
                    "significant": significant,
                    "effect_size_magnitude": r_mag,
                    "higher_group": str(group_a) if median_a > median_b else (str(group_b) if median_b > median_a else "equal"),
                },
                "apa": f"Mann-Whitney U test on {value_column.upper()} by {group_column.upper()} ({group_a} vs {group_b}) found U = {float(u_stat):.1f}, p = {float(p_val):.4g}, r = {float(r_rb):.3f}; Mdn{group_a} = {median_a:.1f}, Mdn{group_b} = {median_b:.1f}.",
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def chi_square_test(
        self,
        file_path: str,
        row_column: str,
        col_column: str,
        sheet_name: str = "Sheet1",
        row_allowed_values: List[str] = None,
        col_allowed_values: List[str] = None,
    ) -> Dict[str, Any]:
        """Chi-square test of independence for two categorical columns."""
        if not OPENPYXL_AVAILABLE:
            return {"success": False, "error": "openpyxl not installed"}
        if not SCIPY_AVAILABLE:
            return {"success": False, "error": "scipy not installed"}
        try:
            wb = load_workbook(file_path, read_only=True)
            ws = self._get_sheet(wb, sheet_name)
            r_idx = openpyxl.utils.column_index_from_string(row_column.upper())
            c_idx = openpyxl.utils.column_index_from_string(col_column.upper())

            pairs = []
            excluded = 0

            def _allowed(cat, allowed):
                if not allowed:
                    return True
                return any(self._value_matches_group(cat, t) for t in allowed)

            for row in ws.iter_rows(min_row=2, values_only=True):
                rv = row[r_idx - 1] if r_idx <= len(row) else None
                cv = row[c_idx - 1] if c_idx <= len(row) else None
                r_cat = self._category_value(rv)
                c_cat = self._category_value(cv)
                if r_cat is None or c_cat is None:
                    continue
                if not _allowed(r_cat, row_allowed_values) or not _allowed(
                    c_cat, col_allowed_values
                ):
                    excluded += 1
                    continue
                pairs.append((r_cat, c_cat))
            wb.close()

            if len(pairs) < 2:
                return {"success": False, "error": "Not enough valid categorical pairs"}

            row_levels = sorted(
                {p[0] for p in pairs}, key=lambda x: (str(type(x)), str(x))
            )
            col_levels = sorted(
                {p[1] for p in pairs}, key=lambda x: (str(type(x)), str(x))
            )

            counts = {r: {c: 0 for c in col_levels} for r in row_levels}
            for r, c in pairs:
                counts[r][c] += 1

            observed = [[counts[r][c] for c in col_levels] for r in row_levels]
            chi2, p, dof, expected = scipy_stats.chi2_contingency(observed)

            n = len(pairs)
            min_dim = min(len(row_levels) - 1, len(col_levels) - 1)
            if min_dim == 0:
                return {
                    "success": False,
                    "error": f"Cramer's V undefined: one variable has only 1 level (rows={len(row_levels)}, cols={len(col_levels)}). Check --row-valid / --col-valid filters.",
                }
            cramers_v = ((chi2 / (n * min_dim)) ** 0.5)
            abs_v = abs(float(cramers_v))
            v_mag = (
                "negligible"
                if abs_v < 0.1
                else (
                    "small" if abs_v < 0.3 else ("medium" if abs_v < 0.5 else "large")
                )
            )
            significant = float(p) < 0.05

            return {
                "success": True,
                "file": file_path,
                "sheet": sheet_name,
                "row_column": row_column.upper(),
                "col_column": col_column.upper(),
                "n": n,
                "excluded_n": excluded,
                "row_allowed_values": row_allowed_values or [],
                "col_allowed_values": col_allowed_values or [],
                "rows": row_levels,
                "cols": col_levels,
                "observed": observed,
                "expected": [list(map(float, r)) for r in expected.tolist()],
                "degrees_of_freedom": int(dof),
                "statistic": float(chi2),
                "p_value": float(p),
                "cramers_v": float(cramers_v),
                "interpretation": {
                    "alpha": 0.05,
                    "significant": significant,
                    "association_strength": v_mag,
                },
                "apa": f"Chi-square test of {row_column.upper()} by {col_column.upper()} was χ²({int(dof)}) = {float(chi2):.3f}, p = {float(p):.4g}, Cramer's V = {float(cramers_v):.3f}, n = {n}.",
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def open_text_extract(
        self,
        file_path: str,
        column_letter: str,
        sheet_name: str = "Sheet1",
        limit: int = 20,
        min_length: int = 20,
    ) -> Dict[str, Any]:
        """Extract non-empty open-text responses from a column for qualitative analysis."""
        if not OPENPYXL_AVAILABLE:
            return {"success": False, "error": "openpyxl not installed"}
        try:
            wb = load_workbook(file_path, read_only=True)
            ws = self._get_sheet(wb, sheet_name)
            col_idx = openpyxl.utils.column_index_from_string(column_letter.upper())

            responses = []
            all_non_empty = 0
            for r in range(2, (ws.max_row or 0) + 1):
                val = ws.cell(r, col_idx).value
                if val is None:
                    continue
                text = str(val).strip()
                if not text:
                    continue
                all_non_empty += 1
                if len(text) < int(min_length):
                    continue
                responses.append(
                    {
                        "row": r,
                        "cell": f"{column_letter.upper()}{r}",
                        "text": text,
                        "length": len(text),
                    }
                )

            wb.close()
            responses = responses[: max(1, int(limit))]
            return {
                "success": True,
                "file": file_path,
                "sheet": sheet_name,
                "column": column_letter.upper(),
                "total_non_empty": all_non_empty,
                "returned": len(responses),
                "limit": int(limit),
                "min_length": int(min_length),
                "responses": responses,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def open_text_keywords(
        self,
        file_path: str,
        column_letter: str,
        sheet_name: str = "Sheet1",
        top_n: int = 15,
        min_word_length: int = 4,
    ) -> Dict[str, Any]:
        """Get keyword frequency summary from an open-text column for theme seeding."""
        if not OPENPYXL_AVAILABLE:
            return {"success": False, "error": "openpyxl not installed"}
        try:
            wb = load_workbook(file_path, read_only=True)
            ws = self._get_sheet(wb, sheet_name)
            col_idx = openpyxl.utils.column_index_from_string(column_letter.upper())

            stopwords = {
                "that",
                "this",
                "with",
                "have",
                "from",
                "more",
                "would",
                "their",
                "they",
                "what",
                "before",
                "into",
                "about",
                "being",
                "because",
                "through",
                "there",
                "where",
                "when",
                "which",
                "just",
                "also",
                "will",
                "than",
                "then",
                "your",
                "them",
                "some",
                "very",
                "much",
                "need",
                "skills",
                "health",
                "experience",
                "placement",
                "clinical",
                "students",
                "student",
                "graduate",
                "university",
                "degree",
                "work",
                "working",
                "field",
                "practical",
                "able",
                "feel",
                "think",
                "would",
                "could",
                "like",
                "more",
                "been",
                "have",
                "being",
                "really",
                "well",
            }

            token_counts = {}
            response_count = 0
            for r in range(2, (ws.max_row or 0) + 1):
                val = ws.cell(r, col_idx).value
                if val is None:
                    continue
                text = str(val).strip().lower()
                if not text:
                    continue
                response_count += 1
                tokens = re.findall(r"[a-zA-Z][a-zA-Z\-']+", text)
                for t in tokens:
                    if len(t) < int(min_word_length):
                        continue
                    if t in stopwords:
                        continue
                    token_counts[t] = token_counts.get(t, 0) + 1
            wb.close()

            top = sorted(token_counts.items(), key=lambda kv: kv[1], reverse=True)[
                : max(1, int(top_n))
            ]
            return {
                "success": True,
                "file": file_path,
                "sheet": sheet_name,
                "column": column_letter.upper(),
                "response_count": response_count,
                "top_n": int(top_n),
                "min_word_length": int(min_word_length),
                "keywords": [
                    {
                        "keyword": k,
                        "count": c,
                        "percent_responses": (c / response_count * 100)
                        if response_count
                        else 0.0,
                    }
                    for k, c in top
                ],
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def research_analysis_pack(
        self,
        file_path: str,
        sheet_name: str = "Sheet0",
        profile: str = "hlth3112",
        require_formula_safe: bool = False,
    ) -> Dict[str, Any]:
        """Run a standardized spreadsheet analysis bundle for research assignments."""
        profile_key = (profile or "hlth3112").lower()
        if profile_key != "hlth3112":
            return {
                "success": False,
                "error": "Unsupported profile. Use: hlth3112",
            }

        result = {
            "success": True,
            "file": file_path,
            "sheet": sheet_name,
            "profile": profile_key,
            "steps": {
                "formula_audit": [],
                "frequencies": [],
                "descriptives": [],
                "correlations": [],
                "ttests": [],
                "chi2": [],
                "qualitative": [],
            },
            "summary": {},
        }

        # Profile configuration tuned for HLTH3112/HLTH3011 survey structure.
        freq_specs = [
            {"column": "A", "valid": ["1", "2"], "label": "Gender"},
            {
                "column": "F",
                "valid": ["1", "2", "3", "4", "5"],
                "label": "Professional_placement",
            },
            {
                "column": "AL",
                "valid": ["1", "2", "3", "4", "5"],
                "label": "More_experience",
            },
        ]
        descriptive_specs = [
            {"column": "M", "op": "all", "label": "Confident_job"},
            {"column": "R", "op": "all", "label": "Overall_prepared"},
            {"column": "AL", "op": "all", "label": "More_experience"},
            {"column": "AM", "op": "all", "label": "More_training"},
        ]
        corr_specs = [
            {
                "x": "M",
                "y": "R",
                "method": "spearman",
                "label": "Confident_job_vs_Overall_prepared",
            },
            {
                "x": "Y",
                "y": "M",
                "method": "spearman",
                "label": "Interpret_statistics_vs_Confident_job",
            },
        ]
        # Mann-Whitney U (non-parametric) for ordinal Likert data
        mannwhitney_specs = [
            {
                "value": "M",
                "group": "F",
                "a": "1",
                "b": "4",
                "label": "Confident_job_by_placement_1_vs_4",
            },
            {
                "value": "R",
                "group": "F",
                "a": "1",
                "b": "4",
                "label": "Overall_prepared_by_placement_1_vs_4",
            },
        ]
        # Parametric t-test (kept for reference/comparison — agent can run separately)
        ttest_specs = []
        chi_specs = [
            {
                "row": "A",
                "col": "AL",
                "row_valid": ["1", "2"],
                "col_valid": ["1", "2", "3", "4", "5"],
                "label": "Gender_vs_More_experience",
            }
        ]
        qualitative_specs = [
            {"column": "AN", "label": "Strongest_skill"},
            {"column": "AO", "label": "Underprepared_skill"},
            {"column": "AP", "label": "Employers_value"},
            {"column": "AQ", "label": "Improve_employability"},
        ]

        def _capture(bucket: str, payload: Dict[str, Any], meta: Dict[str, Any]):
            entry = {"meta": meta, "result": payload}
            result["steps"][bucket].append(entry)

        audit = self.audit_spreadsheet_formulas(
            file_path=file_path,
            sheet_name=sheet_name,
            max_examples=20,
        )
        _capture("formula_audit", audit, {"sheet": sheet_name})

        if require_formula_safe:
            if not audit.get("success"):
                return {
                    "success": False,
                    "file": file_path,
                    "sheet": sheet_name,
                    "profile": profile_key,
                    "steps": result["steps"],
                    "summary": {
                        "total_analyses": 1,
                        "succeeded": 0,
                        "failed": 1,
                        "completion_rate": 0.0,
                        "formula_audit_safe": None,
                    },
                    "error": "Formula safety audit failed and strict policy is enabled",
                }
            if not audit.get("safe_for_cli_formula_eval", False):
                return {
                    "success": False,
                    "file": file_path,
                    "sheet": sheet_name,
                    "profile": profile_key,
                    "steps": result["steps"],
                    "summary": {
                        "total_analyses": 1,
                        "succeeded": 0,
                        "failed": 1,
                        "completion_rate": 0.0,
                        "formula_audit_safe": False,
                    },
                    "error": "Formula safety policy blocked execution: workbook is not safe_for_cli_formula_eval",
                    "formula_audit": audit,
                }

        for spec in freq_specs:
            payload = self.frequencies(
                file_path=file_path,
                column_letter=spec["column"],
                sheet_name=sheet_name,
                allowed_values=spec.get("valid"),
            )
            _capture("frequencies", payload, spec)

        for spec in descriptive_specs:
            payload = self.calculate_column(
                file_path=file_path,
                column_letter=spec["column"],
                operation=spec.get("op", "all"),
                sheet_name=sheet_name,
                include_formulas=False,
            )
            _capture("descriptives", payload, spec)

        for spec in corr_specs:
            payload = self.correlation_test(
                file_path=file_path,
                x_column=spec["x"],
                y_column=spec["y"],
                sheet_name=sheet_name,
                method=spec.get("method", "pearson"),
            )
            _capture("correlations", payload, spec)

        for spec in ttest_specs:
            payload = self.ttest_independent(
                file_path=file_path,
                value_column=spec["value"],
                group_column=spec["group"],
                group_a=spec["a"],
                group_b=spec["b"],
                sheet_name=sheet_name,
                equal_var=False,
            )
            _capture("ttests", payload, spec)

        # Mann-Whitney U tests (non-parametric, appropriate for ordinal/Likert)
        if "mannwhitney" not in result["steps"]:
            result["steps"]["mannwhitney"] = []
        for spec in mannwhitney_specs:
            payload = self.mann_whitney_test(
                file_path=file_path,
                value_column=spec["value"],
                group_column=spec["group"],
                group_a=spec["a"],
                group_b=spec["b"],
                sheet_name=sheet_name,
            )
            _capture("mannwhitney", payload, spec)

        for spec in chi_specs:
            payload = self.chi_square_test(
                file_path=file_path,
                row_column=spec["row"],
                col_column=spec["col"],
                sheet_name=sheet_name,
                row_allowed_values=spec.get("row_valid"),
                col_allowed_values=spec.get("col_valid"),
            )
            _capture("chi2", payload, spec)

        for spec in qualitative_specs:
            kw = self.open_text_keywords(
                file_path=file_path,
                column_letter=spec["column"],
                sheet_name=sheet_name,
                top_n=12,
                min_word_length=4,
            )
            ex = self.open_text_extract(
                file_path=file_path,
                column_letter=spec["column"],
                sheet_name=sheet_name,
                limit=6,
                min_length=35,
            )
            payload = {
                "success": bool(kw.get("success") and ex.get("success")),
                "keywords": kw,
                "quotes": ex,
            }
            if not payload["success"]:
                payload["error"] = kw.get("error") or ex.get("error")
            _capture("qualitative", payload, spec)

        all_results = []
        for bucket in result["steps"].values():
            all_results.extend(bucket)
        ok = sum(1 for x in all_results if x["result"].get("success"))
        fail = len(all_results) - ok

        sig_corr = 0
        for item in result["steps"].get("correlations", []):
            if item["result"].get("interpretation", {}).get("significant"):
                sig_corr += 1
        sig_t = 0
        for item in result["steps"].get("ttests", []):
            if item["result"].get("interpretation", {}).get("significant"):
                sig_t += 1
        sig_chi = 0
        for item in result["steps"].get("chi2", []):
            if item["result"].get("interpretation", {}).get("significant"):
                sig_chi += 1

        result["summary"] = {
            "total_analyses": len(all_results),
            "succeeded": ok,
            "failed": fail,
            "completion_rate": (ok / len(all_results)) if all_results else 0.0,
            "significant": {
                "correlations": sig_corr,
                "ttests": sig_t,
                "chi2": sig_chi,
            },
            "formula_audit_safe": bool(audit.get("safe_for_cli_formula_eval"))
            if audit.get("success")
            else None,
            "require_formula_safe": bool(require_formula_safe),
        }
        if ok == 0:
            result["success"] = False
            result["error"] = (
                "All analyses failed. Check sheet name and dataset structure."
            )

        return result

    def append_to_spreadsheet(
        self, file_path: str, row_data: List[Any], sheet_name: str = None
    ) -> Dict[str, Any]:
        """Append a row to an existing spreadsheet"""
        if not OPENPYXL_AVAILABLE:
            return {"success": False, "error": "openpyxl not installed"}
        try:
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
                wb = load_workbook(file_path)
                ws = self._get_sheet(wb, sheet_name) if sheet_name else wb.active

                # Find the next empty row
                next_row = ws.max_row + 1

                for col_idx, value in enumerate(row_data, 1):
                    try:
                        if (
                            isinstance(value, str)
                            and value.replace(".", "", 1).replace("-", "", 1).isdigit()
                        ):
                            value = float(value) if "." in value else int(value)
                    except Exception:
                        pass
                    ws.cell(row=next_row, column=col_idx, value=value)

                self._safe_save(wb, file_path)
                wb.close()
            return {
                "success": True,
                "file": file_path,
                "row_added": next_row,
                "columns": len(row_data),
                "sheet": ws.title,
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def search_spreadsheet(
        self, file_path: str, search_text: str, sheet_name: str = None
    ) -> Dict[str, Any]:
        """Search for text in a spreadsheet"""
        if not OPENPYXL_AVAILABLE:
            return {"success": False, "error": "openpyxl not installed"}
        try:
            wb = load_workbook(file_path, read_only=True)
            sheets_to_search = (
                [sheet_name]
                if sheet_name and sheet_name in wb.sheetnames
                else wb.sheetnames
            )

            results = []
            for sn in sheets_to_search:
                ws = wb[sn]
                for row_idx, row in enumerate(ws.iter_rows(values_only=False), 1):
                    for col_idx, cell in enumerate(row, 1):
                        if (
                            cell.value is not None
                            and search_text.lower() in str(cell.value).lower()
                        ):
                            col_letter = openpyxl.utils.get_column_letter(col_idx)
                            results.append(
                                {
                                    "sheet": sn,
                                    "cell": f"{col_letter}{row_idx}",
                                    "value": str(cell.value),
                                    "row": row_idx,
                                    "column": col_letter,
                                }
                            )

            wb.close()
            return {
                "success": True,
                "file": file_path,
                "search_text": search_text,
                "matches": len(results),
                "results": results,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def add_formula(
        self, file_path: str, cell: str, formula: str, sheet_name: str = "Sheet1"
    ) -> Dict[str, Any]:
        """Add formula to a specific cell"""
        if not OPENPYXL_AVAILABLE:
            return {"success": False, "error": "openpyxl not installed"}
        try:
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
                wb = load_workbook(file_path)
                ws = self._get_sheet(wb, sheet_name)
                ws[cell] = formula
                self._safe_save(wb, file_path)
                wb.close()
            return {
                "success": True,
                "file": file_path,
                "cell": cell,
                "formula": formula,
                "sheet": sheet_name,
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ==================== CHART OPERATIONS ====================

    def create_chart(
        self,
        file_path: str,
        chart_type: str,
        data_range: str,
        categories_range: str = None,
        title: str = "Chart",
        sheet_name: str = None,
        output_sheet: str = None,
        x_label: str = None,
        y_label: str = None,
        show_data_labels: bool = False,
        show_legend: bool = True,
        legend_pos: str = "right",
        colors: List[str] = None,
    ) -> Dict[str, Any]:
        """
        Create a chart in a spreadsheet.

        Args:
            file_path: Path to the Excel file
            chart_type: Type of chart ('bar', 'line', 'pie', 'scatter')
            data_range: Cell range for data values (e.g., 'B2:D5')
            categories_range: Cell range for category labels (e.g., 'A2:A5')
            title: Chart title
            sheet_name: Source sheet name (default: active sheet)
            output_sheet: Sheet to place chart (default: same as source)
            x_label: X-axis label
            y_label: Y-axis label
            show_data_labels: Show data labels on chart
            show_legend: Show legend
            legend_pos: Legend position ('right', 'top', 'bottom', 'left')
            colors: List of hex colors for chart series

        Returns:
            Dict with success status and chart details
        """
        if not OPENPYXL_AVAILABLE:
            return {"success": False, "error": "openpyxl not installed"}

        try:
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
                wb = load_workbook(file_path)
                ws = self._get_sheet(wb, sheet_name) if sheet_name else wb.active

                # Validate ranges before creating chart
                _ = self._parse_range(data_range)
                if categories_range:
                    _ = self._parse_range(categories_range)
                if not self._range_has_data(ws, data_range):
                    return {
                        "success": False,
                        "error": f"Data range {data_range} contains no usable values",
                    }
                if categories_range and not self._range_has_data(ws, categories_range):
                    return {
                        "success": False,
                        "error": f"Category range {categories_range} contains no usable values",
                    }

                # Create the appropriate chart type
                chart = self._create_chart_object(chart_type)

                # Parse ranges and create references
                if categories_range:
                    min_col, min_row, max_col, max_row = self._parse_range(
                        categories_range
                    )
                    cat_ref = Reference(
                        ws,
                        min_col=min_col,
                        min_row=min_row,
                        max_col=max_col,
                        max_row=max_row,
                    )
                else:
                    # Default: use first column of data range
                    dmin_col, dmin_row, dmax_col, dmax_row = self._parse_range(
                        data_range
                    )
                    cat_ref = Reference(
                        ws,
                        min_col=dmin_col,
                        min_row=dmin_row + 1,
                        max_col=dmin_col,
                        max_row=dmax_row,
                    )

                # Data values (Y-axis)
                min_col, min_row, max_col, max_row = self._parse_range(data_range)
                data_ref = Reference(
                    ws,
                    min_col=min_col,
                    min_row=min_row,
                    max_col=max_col,
                    max_row=max_row,
                )

                chart.add_data(data_ref, titles_from_data=True)
                chart.set_categories(cat_ref)

                chart.title = title

                if not show_legend:
                    chart.has_legend = False
                else:
                    chart.legend.pos = legend_pos

                if x_label:
                    chart.x_axis.title = x_label
                if y_label:
                    chart.y_axis.title = y_label

                if show_data_labels:
                    chart.dataLabels = DataLabelList()
                    chart.dataLabels.show_val = True

                if colors:
                    self._apply_chart_colors(chart, colors)

                if output_sheet and output_sheet != ws.title:
                    if output_sheet not in wb.sheetnames:
                        out_ws = wb.create_sheet(output_sheet)
                    else:
                        out_ws = wb[output_sheet]
                    out_ws.add_chart(chart, "A2")
                else:
                    ws.add_chart(chart, "A10")

                self._safe_save(wb, file_path)
                wb.close()

            return {
                "success": True,
                "file": file_path,
                "chart_type": chart_type,
                "title": title,
                "sheet": ws.title,
                "data_range": data_range,
                "categories_range": categories_range,
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def _create_chart_object(self, chart_type: str):
        """Create the appropriate chart object based on type"""
        chart_type = chart_type.lower()

        if chart_type == "bar" or chart_type == "column":
            chart = BarChart()
            chart.type = "col"  # Column chart (vertical bars)
        elif chart_type == "bar_horizontal":
            chart = BarChart()
            chart.type = "bar"  # Bar chart (horizontal bars)
        elif chart_type == "line":
            chart = LineChart()
        elif chart_type == "pie":
            chart = PieChart()
        elif chart_type == "scatter":
            chart = ScatterChart()
        else:
            raise ValueError(
                f"Unknown chart type: {chart_type}. Use: bar, column, bar_horizontal, line, pie, scatter"
            )

        return chart

    def _apply_chart_colors(self, chart, colors: List[str]):
        """Apply custom colors to chart series"""
        for i, ser in enumerate(chart.series):
            if i < len(colors):
                color = colors[i].lstrip("#")
                ser.graphicalProperties.solidFill.rgb = color

    def create_comparison_chart(
        self,
        file_path: str,
        chart_type: str,
        sheet_name: str = None,
        title: str = "Comparison Chart",
        start_row: int = 1,
        start_col: int = 1,
        num_categories: int = None,
        num_series: int = None,
        category_col: int = 1,
        value_cols: List[int] = None,
        output_cell: str = "A10",
        show_data_labels: bool = False,
        show_legend: bool = True,
    ) -> Dict[str, Any]:
        """
        Create a chart from structured data in a spreadsheet.

        Args:
            file_path: Path to the Excel file
            chart_type: Type of chart ('bar', 'line', 'pie', 'scatter')
            sheet_name: Source sheet name
            title: Chart title
            start_row: Starting row of data (1-indexed, usually 2 to skip header)
            start_col: Starting column of data (1-indexed)
            num_categories: Number of category rows (default: auto-detect)
            num_series: Number of data series (columns)
            category_col: Column containing category labels (1-indexed)
            value_cols: List of columns containing values (1-indexed)
            output_cell: Cell where chart should be placed
            show_data_labels: Show data labels
            show_legend: Show legend

        Returns:
            Dict with success status and chart details
        """
        if not OPENPYXL_AVAILABLE:
            return {"success": False, "error": "openpyxl not installed"}

        try:
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
                wb = load_workbook(file_path)
                ws = self._get_sheet(wb, sheet_name) if sheet_name else wb.active

                # Auto-detect ranges if not specified
                if num_categories is None:
                    num_categories = ws.max_row - start_row + 1
                if num_categories <= 0:
                    return {
                        "success": False,
                        "error": "No category rows found for chart",
                    }

                if value_cols is None:
                    value_cols = list(range(start_col + 1, ws.max_column + 1))
                if not value_cols:
                    return {
                        "success": False,
                        "error": "No value columns found for chart",
                    }

                if num_series is None:
                    num_series = len(value_cols)

                chart = self._create_chart_object(chart_type)

                cat_min_row = start_row
                cat_max_row = start_row + num_categories - 1
                cat_ref = Reference(
                    ws,
                    min_col=category_col,
                    min_row=cat_min_row,
                    max_col=category_col,
                    max_row=cat_max_row,
                )

                for col in value_cols:
                    data_ref = Reference(
                        ws,
                        min_col=col,
                        min_row=cat_min_row,
                        max_col=col,
                        max_row=cat_max_row,
                    )
                    chart.add_data(data_ref, titles_from_data=True)

                chart.set_categories(cat_ref)
                chart.title = title
                chart.has_legend = show_legend

                if show_data_labels:
                    chart.dataLabels = DataLabelList()
                    chart.dataLabels.show_val = True

                ws.add_chart(chart, output_cell)
                self._safe_save(wb, file_path)
                wb.close()

            return {
                "success": True,
                "file": file_path,
                "chart_type": chart_type,
                "title": title,
                "sheet": ws.title,
                "categories": num_categories,
                "series": num_series,
                "output_cell": output_cell,
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def create_grade_distribution_chart(
        self,
        file_path: str,
        sheet_name: str = None,
        grade_column: str = "B",
        title: str = "Grade Distribution",
        output_cell: str = "F2",
    ) -> Dict[str, Any]:
        """
        Create a pie chart showing grade distribution.

        Args:
            file_path: Path to the Excel file
            sheet_name: Sheet name
            grade_column: Column letter containing grades
            title: Chart title
            output_cell: Cell where chart should be placed

        Returns:
            Dict with success status and chart details
        """
        if not OPENPYXL_AVAILABLE:
            return {"success": False, "error": "openpyxl not installed"}

        try:
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
                wb = load_workbook(file_path)
                ws = self._get_sheet(wb, sheet_name) if sheet_name else wb.active

                col_idx = openpyxl.utils.column_index_from_string(grade_column)
                grade_counts = {}
                total_rows = 0

                for row in ws.iter_rows(min_row=2, values_only=True):
                    if col_idx <= len(row) and row[col_idx - 1]:
                        grade = str(row[col_idx - 1]).strip().upper()
                        grade_counts[grade] = grade_counts.get(grade, 0) + 1
                        total_rows += 1

                if not grade_counts:
                    return {
                        "success": False,
                        "error": "No grades found in column " + grade_column,
                    }

                chart = PieChart()
                chart.title = title
                # Write chart helper data to a separate sheet to avoid corrupting the data sheet
                helper_name = "_chart_data"
                if helper_name not in wb.sheetnames:
                    helper_ws = wb.create_sheet(helper_name)
                else:
                    helper_ws = wb[helper_name]
                # Find next free row in helper sheet
                base_row = (helper_ws.max_row or 0) + 2
                helper_ws.cell(row=base_row, column=1, value="Grade")
                helper_ws.cell(row=base_row, column=2, value="Count")
                for i, (grade, count) in enumerate(sorted(grade_counts.items()), 1):
                    helper_ws.cell(row=base_row + i, column=1, value=grade)
                    helper_ws.cell(row=base_row + i, column=2, value=count)

                cat_ref = Reference(
                    helper_ws,
                    min_col=1,
                    min_row=base_row + 1,
                    max_col=1,
                    max_row=base_row + len(grade_counts),
                )
                data_ref = Reference(
                    helper_ws,
                    min_col=2,
                    min_row=base_row + 1,
                    max_col=2,
                    max_row=base_row + len(grade_counts),
                )

                chart.add_data(data_ref, titles_from_data=False)
                chart.set_categories(cat_ref)
                chart.dataLabels = DataLabelList()
                chart.dataLabels.show_val = True
                chart.dataLabels.show_pct = True

                ws.add_chart(chart, output_cell)
                self._safe_save(wb, file_path)
                wb.close()

            return {
                "success": True,
                "file": file_path,
                "chart_type": "pie",
                "title": title,
                "sheet": ws.title,
                "total_grades": total_rows,
                "distribution": grade_counts,
                "output_cell": output_cell,
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def create_progress_chart(
        self,
        file_path: str,
        student_column: str = "A",
        grade_column: str = "B",
        sheet_name: str = None,
        title: str = "Student Progress",
        output_cell: str = "D2",
        show_data_labels: bool = True,
    ) -> Dict[str, Any]:
        """
        Create a bar chart showing individual student grades.

        Args:
            file_path: Path to the Excel file
            student_column: Column letter with student names
            grade_column: Column letter with grades
            sheet_name: Sheet name
            title: Chart title
            output_cell: Cell where chart should be placed
            show_data_labels: Show data labels on bars

        Returns:
            Dict with success status and chart details
        """
        if not OPENPYXL_AVAILABLE:
            return {"success": False, "error": "openpyxl not installed"}

        try:
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
                wb = load_workbook(file_path)
                ws = self._get_sheet(wb, sheet_name) if sheet_name else wb.active

                student_col_idx = openpyxl.utils.column_index_from_string(
                    student_column
                )
                grade_col_idx = openpyxl.utils.column_index_from_string(grade_column)
                num_rows = ws.max_row - 1

                if num_rows <= 0:
                    return {"success": False, "error": "No data found"}

                chart = BarChart()
                chart.type = "bar"
                chart.title = title

                cat_ref = Reference(
                    ws,
                    min_col=student_col_idx,
                    min_row=2,
                    max_col=student_col_idx,
                    max_row=num_rows + 1,
                )

                data_ref = Reference(
                    ws,
                    min_col=grade_col_idx,
                    min_row=2,
                    max_col=grade_col_idx,
                    max_row=num_rows + 1,
                )

                chart.add_data(data_ref, titles_from_data=False)
                chart.set_categories(cat_ref)
                chart.x_axis.title = "Grade"
                chart.y_axis.title = "Student"

                if show_data_labels:
                    chart.dataLabels = DataLabelList()
                    chart.dataLabels.show_val = True

                ws.add_chart(chart, output_cell)
                self._safe_save(wb, file_path)
                wb.close()

            return {
                "success": True,
                "file": file_path,
                "chart_type": "bar_horizontal",
                "title": title,
                "sheet": ws.title,
                "students": num_rows,
                "output_cell": output_cell,
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ==================== PRESENTATION OPERATIONS ====================

    def create_presentation(
        self, output_path: str, title: str = "", subtitle: str = ""
    ) -> Dict[str, Any]:
        """Create a new .pptx presentation with title slide"""
        if not PPTX_AVAILABLE:
            return {"success": False, "error": "python-pptx not installed"}
        try:
            with self._file_lock(output_path):
                backup = self._snapshot_backup(output_path)
                prs = Presentation()
                prs.slide_width = Inches(13.333)   # 16:9 widescreen
                prs.slide_height = Inches(7.5)
                slide_layout = prs.slide_layouts[0]  # Title slide
                slide = prs.slides.add_slide(slide_layout)
                if title:
                    slide.shapes.title.text = title
                if subtitle:
                    slide.placeholders[1].text = subtitle
                self._safe_save(prs, output_path)
            return {
                "success": True,
                "file": output_path,
                "title": title,
                "slides": 1,
                "size": Path(output_path).stat().st_size,
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def add_slide(
        self, file_path: str, title: str, content: str = "", layout: str = "content"
    ) -> Dict[str, Any]:
        """Add a slide. Layouts: title_only, content, blank, two_content"""
        if not PPTX_AVAILABLE:
            return {"success": False, "error": "python-pptx not installed"}
        try:
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
                prs = Presentation(file_path)
                layout_map = {
                    "title_only": 5,
                    "content": 1,
                    "blank": 6,
                    "two_content": 3,
                    "comparison": 4,
                }
                layout_idx = layout_map.get(layout.lower(), 1)
                slide_layout = prs.slide_layouts[layout_idx]
                slide = prs.slides.add_slide(slide_layout)

                if slide.shapes.title:
                    slide.shapes.title.text = title

                if content and len(slide.placeholders) > 1:
                    slide.placeholders[1].text = content

                self._safe_save(prs, file_path)
            return {
                "success": True,
                "file": file_path,
                "title": title,
                "total_slides": len(prs.slides),
                "layout": layout,
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def add_bullet_slide(
        self, file_path: str, title: str, bullets: str
    ) -> Dict[str, Any]:
        """Add a slide with bullet points. Bullets separated by newlines."""
        if not PPTX_AVAILABLE:
            return {"success": False, "error": "python-pptx not installed"}
        try:
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
                prs = Presentation(file_path)
                slide_layout = prs.slide_layouts[1]  # Title and Content
                slide = prs.slides.add_slide(slide_layout)

                if slide.shapes.title:
                    slide.shapes.title.text = title

                if bullets.strip():
                    tf = slide.placeholders[1].text_frame
                    tf.clear()
                    for i, bullet in enumerate(bullets.split("\n")):
                        bullet = bullet.strip()
                        if not bullet:
                            continue
                        if i == 0:
                            p = tf.paragraphs[0]
                        else:
                            p = tf.add_paragraph()
                        p.text = bullet
                        p.level = 0

                self._safe_save(prs, file_path)
            return {
                "success": True,
                "file": file_path,
                "title": title,
                "total_slides": len(prs.slides),
                "bullets": len([b for b in bullets.split("\n") if b.strip()]),
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def read_presentation(self, file_path: str) -> Dict[str, Any]:
        """Read all slides and content from a presentation"""
        if not PPTX_AVAILABLE:
            return {"success": False, "error": "python-pptx not installed"}
        try:
            prs = Presentation(file_path)
            slides_data = []
            for i, slide in enumerate(prs.slides, 1):
                slide_info = {"slide_number": i, "title": "", "content": []}
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        text = shape.text_frame.text.strip()
                        if text:
                            is_title = shape == slide.shapes.title
                            if not is_title:
                                try:
                                    pf = shape.placeholder_format
                                    if pf is not None and pf.idx == 0:
                                        is_title = True
                                except Exception:
                                    pass
                            if is_title:
                                slide_info["title"] = text
                            else:
                                slide_info["content"].append(text)
                slides_data.append(slide_info)

            return {
                "success": True,
                "file": file_path,
                "slide_count": len(slides_data),
                "slides": slides_data,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def add_image_slide(
        self, file_path: str, title: str, image_path: str
    ) -> Dict[str, Any]:
        """Add a slide with an image"""
        if not PPTX_AVAILABLE:
            return {"success": False, "error": "python-pptx not installed"}
        try:
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
                prs = Presentation(file_path)
                slide_layout = prs.slide_layouts[5]  # Title Only
                slide = prs.slides.add_slide(slide_layout)

                if slide.shapes.title:
                    slide.shapes.title.text = title

                slide.shapes.add_picture(
                    image_path, Inches(1), Inches(1.5), width=Inches(8)
                )

                self._safe_save(prs, file_path)
            return {
                "success": True,
                "file": file_path,
                "title": title,
                "total_slides": len(prs.slides),
                "image": image_path,
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def add_table_slide(
        self,
        file_path: str,
        title: str,
        headers_csv: str,
        data_csv: str,
        coerce_rows: bool = False,
    ) -> Dict[str, Any]:
        """Add a slide with a table. Headers comma-separated, rows semicolon-separated."""
        if not PPTX_AVAILABLE:
            return {"success": False, "error": "python-pptx not installed"}
        try:
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
                prs = Presentation(file_path)
                slide_layout = prs.slide_layouts[5]  # Title Only
                slide = prs.slides.add_slide(slide_layout)

                if slide.shapes.title:
                    slide.shapes.title.text = title

                headers = [h.strip() for h in headers_csv.split(",")]
                rows = []
                for row_str in data_csv.split(";"):
                    if row_str.strip():
                        rows.append([c.strip() for c in row_str.split(",")])

                ok, err, rows = self._validate_tabular_rows(
                    headers, rows, coerce_rows=coerce_rows
                )
                if not ok:
                    return {"success": False, "error": err}

                num_rows = len(rows) + 1
                num_cols = len(headers)

                table_shape = slide.shapes.add_table(
                    num_rows, num_cols, Inches(0.5), Inches(1.5), Inches(9), Inches(5)
                )
                table = table_shape.table

                for i, header in enumerate(headers):
                    table.cell(0, i).text = header

                for row_idx, row_data in enumerate(rows, 1):
                    for col_idx, value in enumerate(row_data):
                        table.cell(row_idx, col_idx).text = value

                self._safe_save(prs, file_path)
            return {
                "success": True,
                "file": file_path,
                "title": title,
                "total_slides": len(prs.slides),
                "rows": len(rows),
                "columns": num_cols,
                "coerce_rows": bool(coerce_rows),
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    # Helper methods
    def _hex_to_rgb(self, hex_color: str) -> tuple:
        hex_color = hex_color.lstrip("#")
        return tuple(int(hex_color[i : i + 2], 16) for i in (0, 2, 4))

    def _get_alignment(self, alignment: str) -> int:
        alignments = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
        }
        return alignments.get(alignment.lower(), WD_ALIGN_PARAGRAPH.LEFT)

    # ==================== EXTENDED DOCUMENT OPERATIONS ====================

    def insert_paragraph(
        self, file_path: str, text: str, index: int, style: str = None
    ) -> Dict[str, Any]:
        """Insert a paragraph at a specific index (0 = before first paragraph)."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
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
                # Apply style if requested
                if style:
                    from docx.text.paragraph import Paragraph
                    p = Paragraph(new_para, doc)
                    available = [s.name for s in doc.styles if s.type == 1]
                    if style in available:
                        p.style = doc.styles[style]
                self._safe_save(doc, file_path)
            return {
                "success": True, "file": file_path,
                "inserted_at": index, "text": text[:100],
                "style": style, "total_paragraphs": total + 1,
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def delete_paragraph(
        self, file_path: str, index: int
    ) -> Dict[str, Any]:
        """Delete a paragraph by index."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
                doc = Document(file_path)
                total = len(doc.paragraphs)
                if index < 0 or index >= total:
                    return {"success": False, "error": f"Index {index} out of range (0..{total - 1})"}
                para = doc.paragraphs[index]
                deleted_text = para.text[:200]
                body = doc.element.body
                body.remove(para._element)
                self._safe_save(doc, file_path)
            return {
                "success": True, "file": file_path,
                "deleted_index": index, "deleted_text": deleted_text,
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
            # Define query once — used in both paragraph and table search
            query = search_text if case_sensitive else search_text.lower()
            for i, para in enumerate(doc.paragraphs):
                text = para.text
                target = text if case_sensitive else text.lower()
                start = 0
                while True:
                    idx = target.find(query, start)
                    if idx == -1:
                        break
                    results.append({
                        "paragraph_index": i,
                        "char_offset": idx,
                        "context": text[max(0, idx - 30):idx + len(search_text) + 30],
                        "style": para.style.name if para.style else None,
                    })
                    start = idx + len(search_text)
            # Also search in tables
            table_results = []
            for ti, table in enumerate(doc.tables):
                for ri, row in enumerate(table.rows):
                    for ci, cell in enumerate(row.cells):
                        text = cell.text
                        target = text if case_sensitive else text.lower()
                        if query in target:
                            table_results.append({
                                "table_index": ti,
                                "row": ri, "col": ci,
                                "text": text[:200],
                            })
            return {
                "success": True, "file": file_path,
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
                for ri, row in enumerate(table.rows):
                    row_data = [cell.text for cell in row.cells]
                    rows_data.append(row_data)
                headers = rows_data[0] if rows_data else []
                data = rows_data[1:] if len(rows_data) > 1 else []
                tables.append({
                    "table_index": ti,
                    "headers": headers,
                    "rows": data,
                    "row_count": len(data),
                    "col_count": len(headers),
                })
            return {
                "success": True, "file": file_path,
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
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
                doc = Document(file_path)
                if paragraph_index == -1:
                    para = doc.add_paragraph()
                else:
                    if paragraph_index >= len(doc.paragraphs):
                        return {"success": False, "error": f"Paragraph index {paragraph_index} out of range"}
                    para = doc.paragraphs[paragraph_index]
                # Build hyperlink via OOXML
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
                self._safe_save(doc, file_path)
            return {
                "success": True, "file": file_path,
                "text": text, "url": url,
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
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
                doc = Document(file_path)
                doc.add_page_break()
                self._safe_save(doc, file_path)
            return {"success": True, "file": file_path, "action": "page_break_added", "backup": backup or None}
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
                if s.type == 1:  # paragraph
                    para_styles.append(entry)
                elif s.type == 2:  # character
                    char_styles.append(entry)
            return {
                "success": True, "file": file_path,
                "paragraph_styles": para_styles,
                "character_styles": char_styles,
                "paragraph_style_count": len(para_styles),
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def set_metadata(
        self, file_path: str, author: str = None, title: str = None,
        subject: str = None, keywords: str = None, comments: str = None,
        category: str = None,
    ) -> Dict[str, Any]:
        """Set document core properties (metadata)."""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
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
                self._safe_save(doc, file_path)
            return {
                "success": True, "file": file_path,
                "metadata_set": {k: v for k, v in {
                    "author": author, "title": title, "subject": subject,
                    "keywords": keywords, "comments": comments, "category": category,
                }.items() if v is not None},
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
                "success": True, "file": file_path,
                "author": cp.author, "title": cp.title,
                "subject": cp.subject, "keywords": cp.keywords,
                "comments": cp.comments, "category": cp.category,
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
        """Add a bulleted or numbered list. list_type: bullet|number"""
        if not DOCX_AVAILABLE:
            return {"success": False, "error": "python-docx not installed"}
        try:
            style_name = "List Bullet" if list_type == "bullet" else "List Number"
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
                doc = Document(file_path)
                available = [s.name for s in doc.styles if s.type == 1]
                if style_name not in available:
                    style_name = "Normal"
                for item in items:
                    para = doc.add_paragraph(item, style=style_name)
                self._safe_save(doc, file_path)
            return {
                "success": True, "file": file_path,
                "list_type": list_type, "items_added": len(items),
                "style_used": style_name, "backup": backup or None,
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
            # Rough page estimate (250 words per page)
            pages_est = max(1, round(words / 250))
            return {
                "success": True, "file": file_path,
                "words": words, "characters": chars,
                "characters_no_spaces": chars_no_spaces,
                "paragraphs": paras, "pages_estimate": pages_est,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ==================== EXTENDED SPREADSHEET OPERATIONS ====================

    def cell_read(
        self, file_path: str, cell_ref: str, sheet_name: str = None
    ) -> Dict[str, Any]:
        """Read a single cell value by reference (e.g., 'B5')."""
        if not OPENPYXL_AVAILABLE:
            return {"success": False, "error": "openpyxl not installed"}
        try:
            wb = load_workbook(file_path, data_only=True)
            ws = self._get_sheet(wb, sheet_name) if sheet_name else wb.active
            cell = ws[cell_ref]
            value = cell.value
            wb.close()
            return {
                "success": True, "file": file_path,
                "sheet": ws.title, "cell": cell_ref,
                "value": value, "type": type(value).__name__,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def cell_write(
        self, file_path: str, cell_ref: str, value: str,
        sheet_name: str = None, as_text: bool = False,
    ) -> Dict[str, Any]:
        """Write a value to a single cell."""
        if not OPENPYXL_AVAILABLE:
            return {"success": False, "error": "openpyxl not installed"}
        try:
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
                wb = load_workbook(file_path)
                ws = self._get_sheet(wb, sheet_name) if sheet_name else wb.active
                if as_text:
                    c = ws[cell_ref]
                    c.value = str(value)
                    c.data_type = "s"
                elif isinstance(value, str) and value.startswith("="):
                    ws[cell_ref] = value
                else:
                    try:
                        if "." in str(value):
                            value = float(value)
                        else:
                            value = int(value)
                    except (ValueError, TypeError):
                        pass
                    ws[cell_ref] = value
                self._safe_save(wb, file_path)
                wb.close()
            return {
                "success": True, "file": file_path,
                "sheet": ws.title, "cell": cell_ref,
                "value": value, "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def range_read(
        self, file_path: str, range_ref: str, sheet_name: str = None
    ) -> Dict[str, Any]:
        """Read a range of cells (e.g., 'A1:D10')."""
        if not OPENPYXL_AVAILABLE:
            return {"success": False, "error": "openpyxl not installed"}
        try:
            wb = load_workbook(file_path, data_only=True)
            ws = self._get_sheet(wb, sheet_name) if sheet_name else wb.active
            min_col, min_row, max_col, max_row = self._parse_range(range_ref)
            rows = []
            for row in ws.iter_rows(min_row=min_row, max_row=max_row,
                                     min_col=min_col, max_col=max_col, values_only=True):
                rows.append([str(c) if c is not None else "" for c in row])
            wb.close()
            return {
                "success": True, "file": file_path,
                "sheet": ws.title, "range": range_ref,
                "rows": rows, "row_count": len(rows),
                "col_count": max_col - min_col + 1,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def delete_rows(
        self, file_path: str, start_row: int, count: int = 1, sheet_name: str = None
    ) -> Dict[str, Any]:
        """Delete rows from spreadsheet. start_row is 1-indexed."""
        if not OPENPYXL_AVAILABLE:
            return {"success": False, "error": "openpyxl not installed"}
        try:
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
                wb = load_workbook(file_path)
                ws = self._get_sheet(wb, sheet_name) if sheet_name else wb.active
                ws.delete_rows(start_row, count)
                self._safe_save(wb, file_path)
                wb.close()
            return {
                "success": True, "file": file_path,
                "sheet": ws.title, "deleted_from_row": start_row,
                "rows_deleted": count, "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def delete_columns(
        self, file_path: str, start_col: int, count: int = 1, sheet_name: str = None
    ) -> Dict[str, Any]:
        """Delete columns from spreadsheet. start_col is 1-indexed."""
        if not OPENPYXL_AVAILABLE:
            return {"success": False, "error": "openpyxl not installed"}
        try:
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
                wb = load_workbook(file_path)
                ws = self._get_sheet(wb, sheet_name) if sheet_name else wb.active
                ws.delete_cols(start_col, count)
                self._safe_save(wb, file_path)
                wb.close()
            return {
                "success": True, "file": file_path,
                "sheet": ws.title, "deleted_from_col": start_col,
                "cols_deleted": count, "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def sort_sheet(
        self, file_path: str, column_letter: str, sheet_name: str = None,
        descending: bool = False, numeric: bool = False,
    ) -> Dict[str, Any]:
        """Sort sheet data by a column. Preserves header row."""
        if not OPENPYXL_AVAILABLE:
            return {"success": False, "error": "openpyxl not installed"}
        try:
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
                wb = load_workbook(file_path)
                ws = self._get_sheet(wb, sheet_name) if sheet_name else wb.active
                col_idx = openpyxl.utils.column_index_from_string(column_letter.upper()) - 1
                # Read all data
                all_rows = list(ws.iter_rows(values_only=True))
                if len(all_rows) < 2:
                    wb.close()
                    return {"success": False, "error": "Not enough data to sort (need header + data)"}
                header = all_rows[0]
                data = all_rows[1:]

                def sort_key(row):
                    val = row[col_idx] if col_idx < len(row) else None
                    if val is None:
                        return (1, "")
                    if numeric:
                        try:
                            return (0, float(val))
                        except (ValueError, TypeError):
                            return (1, str(val).lower())
                    return (0, str(val).lower())

                data.sort(key=sort_key, reverse=descending)
                # Clear and rewrite
                for row_cells in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
                    for cell in row_cells:
                        cell.value = None
                for ci, h in enumerate(header, 1):
                    ws.cell(row=1, column=ci, value=h)
                for ri, row in enumerate(data, 2):
                    for ci, val in enumerate(row, 1):
                        ws.cell(row=ri, column=ci, value=val)
                self._safe_save(wb, file_path)
                wb.close()
            return {
                "success": True, "file": file_path,
                "sheet": ws.title, "sorted_by": column_letter.upper(),
                "descending": descending, "rows_sorted": len(data),
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def filter_rows(
        self, file_path: str, column_letter: str, operator: str, value: str,
        sheet_name: str = None,
    ) -> Dict[str, Any]:
        """Filter/query rows. Returns matching rows without modifying the file.
        operator: eq|ne|gt|lt|ge|le|contains|startswith|endswith"""
        if not OPENPYXL_AVAILABLE:
            return {"success": False, "error": "openpyxl not installed"}
        valid_operators = {"eq", "ne", "gt", "lt", "ge", "le", "contains", "startswith", "endswith"}
        if operator not in valid_operators:
            return {
                "success": False,
                "error": f"Unknown operator '{operator}'. Valid: {' | '.join(sorted(valid_operators))}",
            }
        try:
            wb = load_workbook(file_path, read_only=True)
            ws = self._get_sheet(wb, sheet_name) if sheet_name else wb.active
            col_idx = openpyxl.utils.column_index_from_string(column_letter.upper()) - 1
            all_rows = list(ws.iter_rows(values_only=True))
            wb.close()
            if not all_rows:
                return {"success": False, "error": "Empty sheet"}
            header = [str(h) if h else "" for h in all_rows[0]]
            data = all_rows[1:]

            def compare(cell_val):
                if cell_val is None:
                    return False
                cv = str(cell_val)
                if operator in ("gt", "lt", "ge", "le"):
                    try:
                        nv = float(cv)
                        tv = float(value)
                        if operator == "gt": return nv > tv
                        if operator == "lt": return nv < tv
                        if operator == "ge": return nv >= tv
                        if operator == "le": return nv <= tv
                    except (ValueError, TypeError):
                        return False
                if operator == "eq":
                    try:
                        return float(cv) == float(value)
                    except (ValueError, TypeError):
                        return cv.lower() == value.lower()
                if operator == "ne":
                    try:
                        return float(cv) != float(value)
                    except (ValueError, TypeError):
                        return cv.lower() != value.lower()
                if operator == "contains": return value.lower() in cv.lower()
                if operator == "startswith": return cv.lower().startswith(value.lower())
                if operator == "endswith": return cv.lower().endswith(value.lower())
                return False

            matched = []
            for ri, row in enumerate(data, 2):
                cell_val = row[col_idx] if col_idx < len(row) else None
                if compare(cell_val):
                    matched.append({
                        "row_number": ri,
                        "values": [str(c) if c is not None else "" for c in row],
                    })
            return {
                "success": True, "file": file_path,
                "sheet": ws.title, "filter": f"{column_letter.upper()} {operator} {value}",
                "headers": header, "total_rows": len(data),
                "matched_rows": len(matched),
                "rows": matched[:500],
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def sheet_list(self, file_path: str) -> Dict[str, Any]:
        """List all sheets in a workbook with row/column counts."""
        if not OPENPYXL_AVAILABLE:
            return {"success": False, "error": "openpyxl not installed"}
        try:
            wb = load_workbook(file_path, read_only=True)
            sheets = []
            for name in wb.sheetnames:
                ws = wb[name]
                sheets.append({
                    "name": name,
                    "max_row": ws.max_row,
                    "max_column": ws.max_column,
                })
            wb.close()
            return {
                "success": True, "file": file_path,
                "sheet_count": len(sheets), "sheets": sheets,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def sheet_add(
        self, file_path: str, sheet_name: str, position: int = None
    ) -> Dict[str, Any]:
        """Add a new sheet to the workbook."""
        if not OPENPYXL_AVAILABLE:
            return {"success": False, "error": "openpyxl not installed"}
        try:
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
                wb = load_workbook(file_path)
                if sheet_name in wb.sheetnames:
                    wb.close()
                    return {"success": False, "error": f"Sheet '{sheet_name}' already exists"}
                if position is not None:
                    wb.create_sheet(sheet_name, position)
                else:
                    wb.create_sheet(sheet_name)
                self._safe_save(wb, file_path)
                wb.close()
            return {
                "success": True, "file": file_path,
                "sheet_added": sheet_name,
                "sheets": wb.sheetnames if hasattr(wb, 'sheetnames') else [sheet_name],
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def sheet_delete(self, file_path: str, sheet_name: str) -> Dict[str, Any]:
        """Delete a sheet from the workbook."""
        if not OPENPYXL_AVAILABLE:
            return {"success": False, "error": "openpyxl not installed"}
        try:
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
                wb = load_workbook(file_path)
                if sheet_name not in wb.sheetnames:
                    wb.close()
                    return {"success": False, "error": f"Sheet '{sheet_name}' not found. Available: {wb.sheetnames}"}
                if len(wb.sheetnames) == 1:
                    wb.close()
                    return {"success": False, "error": "Cannot delete the only sheet in the workbook"}
                del wb[sheet_name]
                self._safe_save(wb, file_path)
                remaining = wb.sheetnames
                wb.close()
            return {
                "success": True, "file": file_path,
                "sheet_deleted": sheet_name,
                "remaining_sheets": remaining,
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def sheet_rename(
        self, file_path: str, old_name: str, new_name: str
    ) -> Dict[str, Any]:
        """Rename a sheet."""
        if not OPENPYXL_AVAILABLE:
            return {"success": False, "error": "openpyxl not installed"}
        try:
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
                wb = load_workbook(file_path)
                if old_name not in wb.sheetnames:
                    wb.close()
                    return {"success": False, "error": f"Sheet '{old_name}' not found. Available: {wb.sheetnames}"}
                wb[old_name].title = new_name
                self._safe_save(wb, file_path)
                wb.close()
            return {
                "success": True, "file": file_path,
                "old_name": old_name, "new_name": new_name,
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def merge_cells(
        self, file_path: str, range_ref: str, sheet_name: str = None
    ) -> Dict[str, Any]:
        """Merge a range of cells."""
        if not OPENPYXL_AVAILABLE:
            return {"success": False, "error": "openpyxl not installed"}
        try:
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
                wb = load_workbook(file_path)
                ws = self._get_sheet(wb, sheet_name) if sheet_name else wb.active
                ws.merge_cells(range_ref)
                self._safe_save(wb, file_path)
                wb.close()
            return {
                "success": True, "file": file_path,
                "sheet": ws.title, "merged_range": range_ref,
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def unmerge_cells(
        self, file_path: str, range_ref: str, sheet_name: str = None
    ) -> Dict[str, Any]:
        """Unmerge a range of cells."""
        if not OPENPYXL_AVAILABLE:
            return {"success": False, "error": "openpyxl not installed"}
        try:
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
                wb = load_workbook(file_path)
                ws = self._get_sheet(wb, sheet_name) if sheet_name else wb.active
                ws.unmerge_cells(range_ref)
                self._safe_save(wb, file_path)
                wb.close()
            return {
                "success": True, "file": file_path,
                "sheet": ws.title, "unmerged_range": range_ref,
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def format_cells(
        self, file_path: str, range_ref: str, sheet_name: str = None,
        bold: bool = False, italic: bool = False, font_name: str = None,
        font_size: int = None, color: str = None, bg_color: str = None,
        number_format: str = None, alignment: str = None, wrap_text: bool = False,
    ) -> Dict[str, Any]:
        """Apply formatting to a range of cells."""
        if not OPENPYXL_AVAILABLE:
            return {"success": False, "error": "openpyxl not installed"}
        try:
            from openpyxl.styles import Font, PatternFill, Alignment, numbers
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
                wb = load_workbook(file_path)
                ws = self._get_sheet(wb, sheet_name) if sheet_name else wb.active
                min_col, min_row, max_col, max_row = self._parse_range(range_ref)
                count = 0
                for row in ws.iter_rows(min_row=min_row, max_row=max_row,
                                         min_col=min_col, max_col=max_col):
                    for cell in row:
                        if bold or italic or font_name or font_size or color:
                            f = cell.font.copy(
                                bold=bold if bold else cell.font.bold,
                                italic=italic if italic else cell.font.italic,
                                name=font_name if font_name else cell.font.name,
                                size=font_size if font_size else cell.font.size,
                                color=color.lstrip("#") if color else cell.font.color,
                            )
                            cell.font = f
                        if bg_color:
                            cell.fill = PatternFill(
                                start_color=bg_color.lstrip("#"),
                                end_color=bg_color.lstrip("#"),
                                fill_type="solid",
                            )
                        if number_format:
                            cell.number_format = number_format
                        if alignment or wrap_text:
                            h_align = alignment if alignment else None
                            cell.alignment = Alignment(
                                horizontal=h_align,
                                wrap_text=wrap_text,
                            )
                        count += 1
                self._safe_save(wb, file_path)
                wb.close()
            return {
                "success": True, "file": file_path,
                "sheet": ws.title, "range": range_ref,
                "cells_formatted": count, "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def csv_import(
        self, file_path: str, csv_path: str, sheet_name: str = "Sheet1",
        delimiter: str = ",",
    ) -> Dict[str, Any]:
        """Import a CSV file into an xlsx spreadsheet."""
        if not OPENPYXL_AVAILABLE:
            return {"success": False, "error": "openpyxl not installed"}
        try:
            import csv as csv_module
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
                if Path(file_path).exists():
                    wb = load_workbook(file_path)
                    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.create_sheet(sheet_name)
                    ws.delete_rows(1, ws.max_row)
                else:
                    wb = Workbook()
                    ws = wb.active
                    ws.title = sheet_name
                with open(csv_path, "r", newline="", encoding="utf-8-sig") as f:
                    reader = csv_module.reader(f, delimiter=delimiter)
                    row_count = 0
                    for ri, row in enumerate(reader, 1):
                        for ci, val in enumerate(row, 1):
                            try:
                                if "." in val:
                                    val = float(val)
                                else:
                                    val = int(val)
                            except (ValueError, TypeError):
                                pass
                            ws.cell(row=ri, column=ci, value=val)
                        row_count += 1
                self._safe_save(wb, file_path)
                wb.close()
            return {
                "success": True, "file": file_path,
                "csv_source": csv_path, "sheet": sheet_name,
                "rows_imported": row_count, "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def csv_export(
        self, file_path: str, csv_path: str, sheet_name: str = None,
        delimiter: str = ",",
    ) -> Dict[str, Any]:
        """Export a spreadsheet sheet to CSV."""
        if not OPENPYXL_AVAILABLE:
            return {"success": False, "error": "openpyxl not installed"}
        try:
            import csv as csv_module
            wb = load_workbook(file_path, read_only=True)
            ws = self._get_sheet(wb, sheet_name) if sheet_name else wb.active
            row_count = 0
            with open(csv_path, "w", newline="", encoding="utf-8") as f:
                writer = csv_module.writer(f, delimiter=delimiter)
                for row in ws.iter_rows(values_only=True):
                    writer.writerow([str(c) if c is not None else "" for c in row])
                    row_count += 1
            wb.close()
            return {
                "success": True, "file": file_path,
                "csv_output": csv_path, "sheet": ws.title,
                "rows_exported": row_count,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ==================== EXTENDED PRESENTATION OPERATIONS ====================

    def delete_slide(self, file_path: str, slide_index: int) -> Dict[str, Any]:
        """Delete a slide by index (0-based)."""
        if not PPTX_AVAILABLE:
            return {"success": False, "error": "python-pptx not installed"}
        try:
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
                prs = Presentation(file_path)
                total = len(prs.slides)
                if slide_index < 0 or slide_index >= total:
                    return {"success": False, "error": f"Slide index {slide_index} out of range (0..{total - 1})"}
                if total <= 1:
                    return {"success": False, "error": "Cannot delete the only slide"}
                rId = prs.slides._sldIdLst[slide_index].get(qn("r:id"))
                prs.part.drop_rel(rId)
                del prs.slides._sldIdLst[slide_index]
                self._safe_save(prs, file_path)
            return {
                "success": True, "file": file_path,
                "deleted_slide": slide_index,
                "remaining_slides": total - 1,
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def speaker_notes(
        self, file_path: str, slide_index: int, notes_text: str = None
    ) -> Dict[str, Any]:
        """Read or set speaker notes for a slide. If notes_text is None, reads notes."""
        if not PPTX_AVAILABLE:
            return {"success": False, "error": "python-pptx not installed"}
        try:
            if notes_text is None:
                # Read mode
                prs = Presentation(file_path)
                total = len(prs.slides)
                if slide_index < 0 or slide_index >= total:
                    return {"success": False, "error": f"Slide index {slide_index} out of range"}
                slide = prs.slides[slide_index]
                if slide.has_notes_slide:
                    text = slide.notes_slide.notes_text_frame.text
                else:
                    text = ""
                return {
                    "success": True, "file": file_path,
                    "slide_index": slide_index, "has_notes": slide.has_notes_slide,
                    "notes": text, "mode": "read",
                }
            else:
                # Write mode
                with self._file_lock(file_path):
                    backup = self._snapshot_backup(file_path)
                    prs = Presentation(file_path)
                    total = len(prs.slides)
                    if slide_index < 0 or slide_index >= total:
                        return {"success": False, "error": f"Slide index {slide_index} out of range"}
                    slide = prs.slides[slide_index]
                    notes_slide = slide.notes_slide
                    notes_slide.notes_text_frame.text = notes_text
                    self._safe_save(prs, file_path)
                return {
                    "success": True, "file": file_path,
                    "slide_index": slide_index,
                    "notes": notes_text[:200],
                    "mode": "write", "backup": backup or None,
                }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def update_slide_text(
        self, file_path: str, slide_index: int,
        title: str = None, body: str = None,
    ) -> Dict[str, Any]:
        """Update the title and/or body text of an existing slide."""
        if not PPTX_AVAILABLE:
            return {"success": False, "error": "python-pptx not installed"}
        try:
            with self._file_lock(file_path):
                backup = self._snapshot_backup(file_path)
                prs = Presentation(file_path)
                total = len(prs.slides)
                if slide_index < 0 or slide_index >= total:
                    return {"success": False, "error": f"Slide index {slide_index} out of range"}
                slide = prs.slides[slide_index]
                updated = []
                if title is not None and slide.shapes.title:
                    slide.shapes.title.text = title
                    updated.append("title")
                if body is not None:
                    for shape in slide.shapes:
                        if shape.has_text_frame and shape != slide.shapes.title:
                            shape.text_frame.text = body
                            updated.append("body")
                            break
                self._safe_save(prs, file_path)
            return {
                "success": True, "file": file_path,
                "slide_index": slide_index,
                "updated_fields": updated,
                "backup": backup or None,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def slide_count(self, file_path: str) -> Dict[str, Any]:
        """Get the number of slides in a presentation."""
        if not PPTX_AVAILABLE:
            return {"success": False, "error": "python-pptx not installed"}
        try:
            prs = Presentation(file_path)
            count = len(prs.slides)
            titles = []
            for slide in prs.slides:
                t = ""
                if slide.shapes.title:
                    t = slide.shapes.title.text
                titles.append(t)
            return {
                "success": True, "file": file_path,
                "slide_count": count,
                "slide_titles": titles,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ==================== RDF OPERATIONS ====================

    _RDF_FORMAT_MAP = {
        ".ttl": "turtle", ".n3": "n3", ".nt": "nt",
        ".nq": "nquads", ".jsonld": "json-ld", ".json": "json-ld",
        ".rdf": "xml", ".xml": "xml", ".trig": "trig",
    }

    def _rdf_format_from_path(self, file_path: str) -> str:
        """Infer RDF serialisation format from file extension. Defaults to turtle."""
        return self._RDF_FORMAT_MAP.get(Path(file_path).suffix.lower(), "turtle")

    def _rdf_safe_save(self, data: str, file_path: str):
        """Atomic write of a serialised RDF string via temp-file + os.replace."""
        target = Path(file_path)
        fd, tmp_path = tempfile.mkstemp(prefix=f".{target.name}.", dir=str(target.parent))
        try:
            with os.fdopen(fd, "w") as f:
                f.write(data)
            os.replace(tmp_path, str(target))
        except Exception:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass
            raise

    def _get_rdf_graph(self, file_path: str = None):
        """Load or create an RDF graph. Returns (graph, error_string)."""
        try:
            from rdflib import Graph
        except ImportError:
            return None, "rdflib not installed. pip install rdflib"
        g = Graph()
        if file_path:
            if not os.path.exists(file_path):
                return None, f"File not found: {file_path}"
            fmt = self._rdf_format_from_path(file_path)
            g.parse(file_path, format=fmt)
        return g, None

    def rdf_create(
        self, file_path: str, base_uri: str = None, format: str = "turtle",
        prefixes: Dict[str, str] = None,
    ) -> Dict[str, Any]:
        """Create a new empty RDF graph file."""
        try:
            from rdflib import Graph, Namespace
            from rdflib.namespace import RDF, RDFS, OWL, XSD, FOAF, DCTERMS, SKOS
            with self._file_lock(file_path):
                self._snapshot_backup(file_path)
                g = Graph()
                if base_uri:
                    g.bind("base", Namespace(base_uri))
                if prefixes:
                    for prefix, uri in prefixes.items():
                        g.bind(prefix, Namespace(uri))
                for ns_prefix, ns in [
                    ("rdf", RDF), ("rdfs", RDFS), ("owl", OWL),
                    ("xsd", XSD), ("foaf", FOAF), ("dcterms", DCTERMS), ("skos", SKOS),
                ]:
                    g.bind(ns_prefix, ns)
                data = g.serialize(format=format)
                self._rdf_safe_save(data, file_path)
            return {
                "success": True, "file": file_path,
                "format": format, "triples": len(g),
                "prefixes": {p: str(n) for p, n in g.namespaces()},
            }
        except ImportError:
            return {"success": False, "error": "rdflib not installed. pip install rdflib"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def rdf_read(
        self, file_path: str, limit: int = 100
    ) -> Dict[str, Any]:
        """Read and parse an RDF file, returning triples and stats."""
        try:
            g, err = self._get_rdf_graph(file_path)
            if err:
                return {"success": False, "error": err}
            triples = []
            truncated = False
            for i, (s, p, o) in enumerate(g):
                if i >= limit:
                    truncated = True
                    break
                triples.append({
                    "subject": str(s), "predicate": str(p), "object": str(o),
                })
            subjects = set(str(s) for s in g.subjects())
            predicates = set(str(p) for p in g.predicates())
            return {
                "success": True, "file": file_path,
                "total_triples": len(g),
                "unique_subjects": len(subjects),
                "unique_predicates": len(predicates),
                "prefixes": {p: str(n) for p, n in g.namespaces()},
                "triples": triples,
                "truncated": truncated,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def rdf_add(
        self, file_path: str, subject: str, predicate: str, object_val: str,
        object_type: str = "uri", lang: str = None, datatype: str = None,
        format: str = None,
    ) -> Dict[str, Any]:
        """Add a triple to an RDF graph. object_type: uri|literal|bnode"""
        try:
            from rdflib import URIRef, Literal, BNode
            if lang and datatype:
                return {"success": False, "error": "Cannot specify both --lang and --datatype for a literal"}
            with self._file_lock(file_path):
                self._snapshot_backup(file_path)
                g, err = self._get_rdf_graph(file_path)
                if err:
                    return {"success": False, "error": err}
                fmt = format or self._rdf_format_from_path(file_path)
                s = URIRef(subject)
                p = URIRef(predicate)
                if object_type == "literal":
                    dt = URIRef(datatype) if datatype else None
                    o = Literal(object_val, lang=lang, datatype=dt)
                elif object_type == "bnode":
                    o = BNode(object_val)
                else:
                    o = URIRef(object_val)
                before = len(g)
                g.add((s, p, o))
                data = g.serialize(format=fmt)
                self._rdf_safe_save(data, file_path)
            return {
                "success": True, "file": file_path,
                "triple": {"subject": str(s), "predicate": str(p), "object": str(o)},
                "triples_before": before, "triples_after": len(g),
                "format": fmt,
            }
        except ImportError:
            return {"success": False, "error": "rdflib not installed"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def rdf_remove(
        self, file_path: str, subject: str = None, predicate: str = None,
        object_val: str = None, object_type: str = "uri", format: str = None,
    ) -> Dict[str, Any]:
        """Remove triples matching a pattern. None values act as wildcards.
        object_type: uri|literal|bnode — controls how object_val is interpreted."""
        try:
            from rdflib import URIRef, Literal, BNode
            with self._file_lock(file_path):
                self._snapshot_backup(file_path)
                g, err = self._get_rdf_graph(file_path)
                if err:
                    return {"success": False, "error": err}
                fmt = format or self._rdf_format_from_path(file_path)
                s = URIRef(subject) if subject else None
                p = URIRef(predicate) if predicate else None
                if object_val is None:
                    o = None
                elif object_type == "literal":
                    o = Literal(object_val)
                elif object_type == "bnode":
                    o = BNode(object_val)
                else:
                    o = URIRef(object_val)
                before = len(g)
                g.remove((s, p, o))
                data = g.serialize(format=fmt)
                self._rdf_safe_save(data, file_path)
            return {
                "success": True, "file": file_path,
                "triples_before": before, "triples_after": len(g),
                "removed": before - len(g), "format": fmt,
            }
        except ImportError:
            return {"success": False, "error": "rdflib not installed"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def rdf_query(
        self, file_path: str, sparql: str, limit: int = 100
    ) -> Dict[str, Any]:
        """Execute a SPARQL query against an RDF graph.
        Handles SELECT, ASK, CONSTRUCT, and DESCRIBE query types correctly."""
        try:
            g, err = self._get_rdf_graph(file_path)
            if err:
                return {"success": False, "error": err}
            results = g.query(sparql)
            qtype = results.type  # "SELECT" | "ASK" | "CONSTRUCT" | "DESCRIBE"

            if qtype == "ASK":
                return {
                    "success": True, "file": file_path,
                    "query_type": "ASK",
                    "result": bool(results),
                }

            if qtype in ("CONSTRUCT", "DESCRIBE"):
                triples = []
                truncated = False
                for i, (s, p, o) in enumerate(results.graph):
                    if i >= limit:
                        truncated = True
                        break
                    triples.append({"subject": str(s), "predicate": str(p), "object": str(o)})
                return {
                    "success": True, "file": file_path,
                    "query_type": qtype,
                    "triples": triples,
                    "result_count": len(triples),
                    "truncated": truncated,
                }

            # SELECT
            variables = [str(v) for v in results.vars] if results.vars else []
            rows = []
            truncated = False
            for i, row in enumerate(results):
                if i >= limit:
                    truncated = True
                    break
                rows.append({
                    str(v): str(row[v]) if row[v] is not None else None
                    for v in results.vars
                })
            return {
                "success": True, "file": file_path,
                "query_type": "SELECT",
                "variables": variables,
                "result_count": len(rows),
                "rows": rows,
                "truncated": truncated,
            }
        except ImportError:
            return {"success": False, "error": "rdflib not installed"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def rdf_export(
        self, file_path: str, output_path: str, output_format: str = "turtle"
    ) -> Dict[str, Any]:
        """Convert/export an RDF graph to a different serialization format.
        Formats: turtle, xml, n3, nt, nquads, json-ld, trig"""
        try:
            g, err = self._get_rdf_graph(file_path)
            if err:
                return {"success": False, "error": err}
            data = g.serialize(format=output_format)
            self._rdf_safe_save(data, output_path)
            return {
                "success": True, "source": file_path,
                "output": output_path, "format": output_format,
                "triples": len(g),
                "size": os.path.getsize(output_path),
            }
        except ImportError:
            return {"success": False, "error": "rdflib not installed"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def rdf_merge(
        self, file_path: str, other_path: str, output_path: str = None,
        format: str = None,
    ) -> Dict[str, Any]:
        """Merge two RDF graphs. Result goes to output_path or overwrites file_path."""
        try:
            abs_a = str(Path(file_path).resolve())
            abs_b = str(Path(other_path).resolve())
            if abs_a == abs_b:
                return {
                    "success": False,
                    "error": "Cannot merge a file with itself — use rdf-export to copy",
                }
            target = output_path or file_path
            with self._file_lock(target):
                self._snapshot_backup(target)
                g1, err = self._get_rdf_graph(file_path)
                if err:
                    return {"success": False, "error": err}
                g2, err = self._get_rdf_graph(other_path)
                if err:
                    return {"success": False, "error": err}
                fmt = format or self._rdf_format_from_path(target)
                before = len(g1)
                g1 += g2
                data = g1.serialize(format=fmt)
                self._rdf_safe_save(data, target)
            return {
                "success": True, "file": target,
                "graph_a_triples": before, "graph_b_triples": len(g2),
                "merged_triples": len(g1), "format": fmt,
            }
        except ImportError:
            return {"success": False, "error": "rdflib not installed"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def rdf_stats(self, file_path: str) -> Dict[str, Any]:
        """Get detailed statistics about an RDF graph."""
        try:
            from rdflib import Literal as RDFLiteral
            from rdflib.namespace import RDF
            g, err = self._get_rdf_graph(file_path)
            if err:
                return {"success": False, "error": err}
            subjects = set(str(s) for s in g.subjects())
            predicates = set(str(p) for p in g.predicates())
            objects_uri = set()
            objects_literal = set()
            for o in g.objects():
                if isinstance(o, RDFLiteral):
                    objects_literal.add(str(o))
                else:
                    objects_uri.add(str(o))
            # Class instances via rdf:type
            classes = set()
            for s, p, o in g.triples((None, RDF.type, None)):
                classes.add(str(o))
            # Predicate frequency
            pred_freq: Dict[str, int] = {}
            for s, p, o in g:
                ps = str(p)
                pred_freq[ps] = pred_freq.get(ps, 0) + 1
            top_predicates = sorted(pred_freq.items(), key=lambda x: x[1], reverse=True)[:20]
            return {
                "success": True, "file": file_path,
                "total_triples": len(g),
                "unique_subjects": len(subjects),
                "unique_predicates": len(predicates),
                "unique_objects_uri": len(objects_uri),
                "unique_objects_literal": len(objects_literal),
                "rdf_types": sorted(classes),
                "type_count": len(classes),
                "prefixes": {p: str(n) for p, n in g.namespaces()},
                "top_predicates": [{"predicate": p, "count": c} for p, c in top_predicates],
            }
        except ImportError:
            return {"success": False, "error": "rdflib not installed"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def rdf_namespace(
        self, file_path: str, prefix: str = None, uri: str = None,
        format: str = None,
    ) -> Dict[str, Any]:
        """Add a namespace prefix or list all prefixes. If prefix+uri given, adds. Otherwise lists."""
        try:
            from rdflib import Namespace
            if prefix and uri:
                with self._file_lock(file_path):
                    self._snapshot_backup(file_path)
                    g, err = self._get_rdf_graph(file_path)
                    if err:
                        return {"success": False, "error": err}
                    fmt = format or self._rdf_format_from_path(file_path)
                    g.bind(prefix, Namespace(uri))
                    data = g.serialize(format=fmt)
                    self._rdf_safe_save(data, file_path)
                return {
                    "success": True, "file": file_path,
                    "action": "added", "prefix": prefix, "uri": uri,
                    "all_prefixes": {p: str(n) for p, n in g.namespaces()},
                }
            else:
                g, err = self._get_rdf_graph(file_path)
                if err:
                    return {"success": False, "error": err}
                return {
                    "success": True, "file": file_path,
                    "prefixes": {p: str(n) for p, n in g.namespaces()},
                    "prefix_count": len(list(g.namespaces())),
                }
        except ImportError:
            return {"success": False, "error": "rdflib not installed"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def rdf_validate(self, file_path: str, shapes_path: str) -> Dict[str, Any]:
        """Validate an RDF graph against SHACL shapes."""
        try:
            from pyshacl import validate as shacl_validate
            from rdflib import URIRef
            g, err = self._get_rdf_graph(file_path)
            if err:
                return {"success": False, "error": err}
            sg, err = self._get_rdf_graph(shapes_path)
            if err:
                return {"success": False, "error": f"Shapes file error: {err}"}
            conforms, results_graph, results_text = shacl_validate(
                g, shacl_graph=sg, inference="rdfs"
            )
            # Parse structured violations from results_graph
            violations = []
            try:
                SH = "http://www.w3.org/ns/shacl#"
                for result_node in results_graph.objects(None, URIRef(f"{SH}result")):
                    v: Dict[str, Any] = {}
                    fn = next(results_graph.objects(result_node, URIRef(f"{SH}focusNode")), None)
                    if fn:
                        v["focusNode"] = str(fn)
                    rp = next(results_graph.objects(result_node, URIRef(f"{SH}resultPath")), None)
                    if rp:
                        v["resultPath"] = str(rp)
                    rm = next(results_graph.objects(result_node, URIRef(f"{SH}resultMessage")), None)
                    if rm:
                        v["message"] = str(rm)
                    rs = next(results_graph.objects(result_node, URIRef(f"{SH}resultSeverity")), None)
                    if rs:
                        v["severity"] = str(rs).split("#")[-1]
                    sc = next(results_graph.objects(result_node, URIRef(f"{SH}sourceConstraintComponent")), None)
                    if sc:
                        v["constraint"] = str(sc).split("#")[-1]
                    violations.append(v)
            except Exception:
                pass  # violation parsing is best-effort
            return {
                "success": True, "file": file_path,
                "shapes_file": shapes_path,
                "conforms": conforms,
                "violation_count": len(violations),
                "violations": violations,
                "results_text": results_text[:5000],
            }
        except ImportError:
            return {"success": False, "error": "pyshacl not installed. pip install pyshacl"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ==================== HELPERS ====================

    def check_health(self) -> bool:
        try:
            response = requests.get(f"{self.server_url}/healthcheck", timeout=5)
            return response.text.strip() == "true"
        except Exception:
            return False


_client = None


def get_client():
    global _client
    if _client is None:
        _client = DocumentServerClient(
            server_url=os.environ.get("ONLYOFFICE_URL", "http://localhost:8080"),
            secret=os.environ.get("ONLYOFFICE_SECRET", "sloane-os-secret-key-2026"),
        )
    return _client
