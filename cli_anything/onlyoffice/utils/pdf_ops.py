#!/usr/bin/env python3
"""PDF operations for the OnlyOffice CLI."""

from __future__ import annotations

import io
import os
import re
import tempfile
from pathlib import Path
from typing import Any, Dict, List, Optional


class PDFOperations:
    """Encapsulate PDF inspection, extraction, rendering, and sanitization."""

    def __init__(self, host: Any):
        self.host = host

    @staticmethod
    def parse_page_range(pages_str: str, total: int) -> List[int]:
        """Parse a page range string like ``0-3`` or ``1,3,5``."""
        if not pages_str:
            return list(range(total))
        result = []
        for part in pages_str.split(","):
            part = part.strip()
            if "-" in part:
                start, end = part.split("-", 1)
                start = max(0, int(start))
                end = min(total - 1, int(end))
                result.extend(range(start, end + 1))
            else:
                index = int(part)
                if 0 <= index < total:
                    result.append(index)
        return sorted(set(result))

    @staticmethod
    def normalize_bbox(bbox) -> Optional[Dict[str, Any]]:
        if not bbox or len(bbox) < 4:
            return None
        left, top, right, bottom = [float(value) for value in bbox[:4]]
        return {
            "left": left,
            "top": top,
            "right": right,
            "bottom": bottom,
            "width": max(0.0, right - left),
            "height": max(0.0, bottom - top),
        }

    def inspect_hidden_data(self, file_path: str) -> Dict[str, Any]:
        """Inspect hidden PDF metadata, annotations, embedded files, and page-size consistency."""
        try:
            import fitz
        except ImportError:
            return {"success": False, "error": "PyMuPDF (fitz) not installed. Run: pip install PyMuPDF"}

        try:
            abs_path = str(Path(file_path).resolve())
            pdf = fitz.open(abs_path)
            metadata = dict(pdf.metadata or {})
            has_xml_metadata = False
            try:
                has_xml_metadata = bool(pdf.xref_xml_metadata())
            except Exception:
                has_xml_metadata = False

            page_sizes = []
            labels = []
            annotations_total = 0
            annotation_types: Dict[str, int] = {}

            for page_index, page in enumerate(pdf):
                label = self.host._label_pdf_page_size(page.rect.width, page.rect.height)
                labels.append(label)
                page_sizes.append(
                    {
                        "page_index": page_index,
                        "width": float(page.rect.width),
                        "height": float(page.rect.height),
                        "label": label,
                    }
                )
                annot = page.first_annot
                while annot is not None:
                    annotations_total += 1
                    try:
                        annot_type = annot.type[1]
                    except Exception:
                        annot_type = "unknown"
                    annotation_types[annot_type] = annotation_types.get(annot_type, 0) + 1
                    annot = annot.next

            embedded_files = 0
            for accessor in ("embfile_count", "embedded_file_count"):
                if hasattr(pdf, accessor):
                    try:
                        embedded_files = int(getattr(pdf, accessor)())
                        break
                    except Exception:
                        pass

            form_fields = None
            if hasattr(pdf, "is_form_pdf"):
                try:
                    form_fields = bool(pdf.is_form_pdf)
                except Exception:
                    form_fields = None

            pdf.close()

            nonempty_metadata = {
                key: value
                for key, value in metadata.items()
                if key != "format" and str(value or "").strip()
            }
            unique_labels = sorted(set(labels))
            return {
                "success": True,
                "file": abs_path,
                "metadata": metadata,
                "nonempty_metadata": nonempty_metadata,
                "has_xml_metadata": has_xml_metadata,
                "pages": len(page_sizes),
                "page_sizes": page_sizes,
                "page_size_labels": unique_labels,
                "page_size_consistent": len(unique_labels) <= 1,
                "annotations_count": annotations_total,
                "annotation_types": annotation_types,
                "embedded_files_count": embedded_files,
                "has_embedded_files": embedded_files > 0,
                "has_forms": form_fields,
            }
        except Exception as exc:
            return {"success": False, "error": str(exc)}

    def extract_images(
        self,
        file_path: str,
        output_dir: str,
        fmt: str = "png",
        pages: str = None,
    ) -> Dict[str, Any]:
        """Extract embedded image objects from a PDF using PyMuPDF."""
        try:
            import fitz
            from PIL import Image as PILImage
        except ImportError:
            return {"success": False, "error": "PyMuPDF (fitz) not installed. Run: pip install PyMuPDF"}

        try:
            os.makedirs(output_dir, exist_ok=True)
            doc = fitz.open(file_path)
            total_pages = len(doc)
            page_indices = self.parse_page_range(pages, total_pages)
            extracted = []
            seen_xrefs = set()
            index = 0
            for page_index in page_indices:
                page = doc[page_index]
                for img_info in page.get_images(full=True):
                    xref = img_info[0]
                    if xref in seen_xrefs:
                        continue
                    seen_xrefs.add(xref)
                    try:
                        base_image = doc.extract_image(xref)
                        if not base_image:
                            continue
                        image_bytes = base_image["image"]
                        img_ext = base_image.get("ext", "png")
                        out_name = f"pdf_img_{page_index:03d}_{index:03d}.{fmt}"
                        out_path = os.path.join(output_dir, out_name)
                        img = PILImage.open(io.BytesIO(image_bytes))
                        if fmt.lower() == "jpg":
                            img = img.convert("RGB")
                            img.save(out_path, "JPEG", quality=90)
                        else:
                            img.save(out_path, fmt.upper())
                        extracted.append(
                            {
                                "index": index,
                                "page": page_index,
                                "xref": xref,
                                "file": out_path,
                                "format": fmt,
                                "width": img.width,
                                "height": img.height,
                                "size_bytes": os.path.getsize(out_path),
                                "original_format": img_ext,
                            }
                        )
                    except Exception as img_err:
                        extracted.append(
                            {
                                "index": index,
                                "page": page_index,
                                "xref": xref,
                                "error": str(img_err),
                            }
                        )
                    index += 1
            doc.close()
            return {
                "success": True,
                "file": file_path,
                "output_dir": output_dir,
                "total_pages": total_pages,
                "pages_scanned": len(page_indices),
                "images_extracted": len([entry for entry in extracted if "file" in entry]),
                "images": extracted,
            }
        except Exception as exc:
            return {"success": False, "error": str(exc)}

    def read_blocks(
        self,
        file_path: str,
        pages: str = None,
        include_spans: bool = True,
        include_images: bool = True,
        include_empty: bool = False,
    ) -> Dict[str, Any]:
        """Read native PDF text blocks, lines, and spans with bounding boxes."""
        try:
            import fitz
        except ImportError:
            return {"success": False, "error": "PyMuPDF (fitz) not installed. Run: pip install PyMuPDF"}

        try:
            doc = fitz.open(file_path)
            total_pages = len(doc)
            page_indices = self.parse_page_range(pages, total_pages)
            pages_payload = []
            text_block_count = 0
            image_block_count = 0
            span_count = 0

            for page_index in page_indices:
                page = doc[page_index]
                page_dict = page.get_text("dict", sort=True)
                page_blocks = []

                for block_index, block in enumerate(page_dict.get("blocks", [])):
                    block_type = int(block.get("type", 0))
                    block_bbox = self.normalize_bbox(block.get("bbox"))
                    block_id = f"page_{page_index}_block_{block_index}"

                    if block_type == 0:
                        lines_payload = []
                        block_text_parts = []
                        for line_index, line in enumerate(block.get("lines", [])):
                            line_bbox = self.normalize_bbox(line.get("bbox"))
                            line_id = f"{block_id}_line_{line_index}"
                            line_spans = []
                            line_text_parts = []

                            for span_index, span in enumerate(line.get("spans", [])):
                                span_text = str(span.get("text", "") or "")
                                if not include_empty and not span_text.strip():
                                    continue
                                span_id = f"{line_id}_span_{span_index}"
                                span_payload = {
                                    "span_id": span_id,
                                    "text": span_text,
                                    "bbox": self.normalize_bbox(span.get("bbox")),
                                    "font": span.get("font"),
                                    "size": float(span.get("size")) if span.get("size") is not None else None,
                                    "flags": int(span.get("flags", 0)) if span.get("flags") is not None else None,
                                    "color": int(span.get("color", 0)) if span.get("color") is not None else None,
                                    "origin": list(span.get("origin", [])) if isinstance(span.get("origin"), (list, tuple)) else None,
                                }
                                span_count += 1
                                line_spans.append(span_payload)
                                if span_text.strip():
                                    line_text_parts.append(span_text)

                            line_text = re.sub(r"\s+", " ", " ".join(line_text_parts)).strip()
                            if line_spans or include_empty or line_text:
                                line_payload = {
                                    "line_id": line_id,
                                    "bbox": line_bbox,
                                    "text": line_text,
                                }
                                if include_spans:
                                    line_payload["spans"] = line_spans
                                lines_payload.append(line_payload)
                                if line_text:
                                    block_text_parts.append(line_text)

                        block_text = "\n".join([entry for entry in block_text_parts if entry]).strip()
                        if block_text or lines_payload or include_empty:
                            text_block_count += 1
                            page_blocks.append(
                                {
                                    "block_id": block_id,
                                    "block_index": block_index,
                                    "type": "text",
                                    "bbox": block_bbox,
                                    "text": block_text,
                                    "line_count": len(lines_payload),
                                    "span_count": sum(len(line.get("spans", [])) for line in lines_payload),
                                    "lines": lines_payload,
                                }
                            )
                    elif block_type == 1 and include_images:
                        image_block_count += 1
                        page_blocks.append(
                            {
                                "block_id": block_id,
                                "block_index": block_index,
                                "type": "image",
                                "bbox": block_bbox,
                                "width": block.get("width"),
                                "height": block.get("height"),
                                "ext": block.get("ext"),
                                "transform": list(block.get("transform", [])) if isinstance(block.get("transform"), (list, tuple)) else None,
                            }
                        )

                pages_payload.append(
                    {
                        "page_index": page_index,
                        "page_number": page_index + 1,
                        "width": float(page.rect.width),
                        "height": float(page.rect.height),
                        "block_count": len(page_blocks),
                        "blocks": page_blocks,
                    }
                )

            doc.close()
            return {
                "success": True,
                "file": file_path,
                "total_pages": total_pages,
                "pages_scanned": len(page_indices),
                "text_block_count": text_block_count,
                "image_block_count": image_block_count,
                "span_count": span_count,
                "pages": pages_payload,
            }
        except Exception as exc:
            return {"success": False, "error": str(exc)}

    def search_blocks(
        self,
        file_path: str,
        query: str,
        pages: str = None,
        case_sensitive: bool = False,
        include_spans: bool = True,
    ) -> Dict[str, Any]:
        """Search native PDF block/span text and return exact anchors."""
        if not str(query or "").strip():
            return {"success": False, "error": "query is required"}

        payload = self.read_blocks(
            file_path,
            pages=pages,
            include_spans=include_spans,
            include_images=False,
            include_empty=False,
        )
        if not payload.get("success"):
            return payload

        query_text = str(query)
        haystack_query = query_text if case_sensitive else query_text.lower()
        matches = []

        for page in payload.get("pages", []):
            for block in page.get("blocks", []):
                if block.get("type") != "text":
                    continue
                block_text = str(block.get("text", "") or "")
                block_cmp = block_text if case_sensitive else block_text.lower()
                if haystack_query in block_cmp:
                    matches.append(
                        {
                            "page_index": page.get("page_index"),
                            "page_number": page.get("page_number"),
                            "block_id": block.get("block_id"),
                            "line_id": None,
                            "span_id": None,
                            "scope": "block",
                            "bbox": block.get("bbox"),
                            "text": block_text,
                            "match_text": query_text,
                        }
                    )

                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        span_text = str(span.get("text", "") or "")
                        span_cmp = span_text if case_sensitive else span_text.lower()
                        if haystack_query not in span_cmp:
                            continue
                        matches.append(
                            {
                                "page_index": page.get("page_index"),
                                "page_number": page.get("page_number"),
                                "block_id": block.get("block_id"),
                                "line_id": line.get("line_id"),
                                "span_id": span.get("span_id"),
                                "scope": "span",
                                "bbox": span.get("bbox"),
                                "text": span_text,
                                "match_text": query_text,
                                "font": span.get("font"),
                                "size": span.get("size"),
                                "flags": span.get("flags"),
                            }
                        )

        return {
            "success": True,
            "file": file_path,
            "query": query_text,
            "case_sensitive": bool(case_sensitive),
            "match_count": len(matches),
            "matches": matches,
            "pages_scanned": payload.get("pages_scanned", 0),
            "total_pages": payload.get("total_pages", 0),
        }

    def page_to_image(
        self,
        file_path: str,
        output_dir: str,
        pages: str = None,
        dpi: int = 150,
        fmt: str = "png",
    ) -> Dict[str, Any]:
        """Render full PDF pages as images using PyMuPDF."""
        try:
            import fitz
        except ImportError:
            return {"success": False, "error": "PyMuPDF (fitz) not installed. Run: pip install PyMuPDF"}

        try:
            os.makedirs(output_dir, exist_ok=True)
            doc = fitz.open(file_path)
            total_pages = len(doc)
            page_indices = self.parse_page_range(pages, total_pages)
            rendered = []
            for page_index in page_indices:
                page = doc[page_index]
                pix = page.get_pixmap(dpi=dpi)
                if fmt.lower() == "jpg":
                    out_name = f"page_{page_index:03d}.jpg"
                    out_path = os.path.join(output_dir, out_name)
                    pix.pil_save(out_path, format="JPEG", quality=90)
                else:
                    out_name = f"page_{page_index:03d}.png"
                    out_path = os.path.join(output_dir, out_name)
                    pix.save(out_path)
                rendered.append(
                    {
                        "page": page_index,
                        "file": out_path,
                        "format": fmt,
                        "width": pix.width,
                        "height": pix.height,
                        "dpi": dpi,
                        "size_bytes": os.path.getsize(out_path),
                    }
                )
            doc.close()
            return {
                "success": True,
                "file": file_path,
                "output_dir": output_dir,
                "total_pages": total_pages,
                "pages_rendered": len(rendered),
                "dpi": dpi,
                "images": rendered,
            }
        except Exception as exc:
            return {"success": False, "error": str(exc)}

    def sanitize(
        self,
        file_path: str,
        output_path: str = None,
        *,
        clear_metadata: bool = False,
        remove_xml_metadata: bool = False,
        author: Optional[str] = None,
        title: Optional[str] = None,
        subject: Optional[str] = None,
        keywords: Optional[str] = None,
        creator: Optional[str] = None,
        producer: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Sanitize PDF metadata for submission workflows."""
        try:
            import fitz
        except ImportError:
            return {"success": False, "error": "PyMuPDF (fitz) not installed. Run: pip install PyMuPDF"}

        if not any(
            [
                clear_metadata,
                remove_xml_metadata,
                author is not None,
                title is not None,
                subject is not None,
                keywords is not None,
                creator is not None,
                producer is not None,
            ]
        ):
            return {"success": False, "error": "No sanitization options provided"}

        try:
            abs_path = str(Path(file_path).resolve())
            out_path = str(Path(output_path).resolve()) if output_path else abs_path
            overwrite = out_path == abs_path

            with self.host._file_lock(abs_path):
                backup_target = abs_path if overwrite else out_path
                backup = self.host._snapshot_backup(backup_target) if Path(backup_target).exists() else None
                pdf = fitz.open(abs_path)
                try:
                    before = dict(pdf.metadata or {})
                    metadata = {} if clear_metadata else dict(before)
                    assignments = {
                        "author": author,
                        "title": title,
                        "subject": subject,
                        "keywords": keywords,
                        "creator": creator,
                        "producer": producer,
                    }
                    if clear_metadata:
                        for key in [
                            "author",
                            "title",
                            "subject",
                            "keywords",
                            "creator",
                            "producer",
                            "creationDate",
                            "modDate",
                        ]:
                            metadata[key] = ""
                    for key, value in assignments.items():
                        if value is not None:
                            metadata[key] = value

                    pdf.set_metadata(metadata)
                    if remove_xml_metadata and hasattr(pdf, "del_xml_metadata"):
                        pdf.del_xml_metadata()

                    target = Path(out_path)
                    target.parent.mkdir(parents=True, exist_ok=True)
                    fd, temp_output = tempfile.mkstemp(
                        prefix=f".{target.name}.",
                        suffix=".tmp",
                        dir=str(target.parent),
                    )
                    os.close(fd)
                    try:
                        pdf.save(temp_output, garbage=4, deflate=True, clean=True)
                        os.replace(temp_output, out_path)
                    finally:
                        if os.path.exists(temp_output):
                            os.unlink(temp_output)
                finally:
                    pdf.close()

                out_doc = fitz.open(out_path)
                after = dict(out_doc.metadata or {})
                has_xml_metadata = False
                try:
                    has_xml_metadata = bool(out_doc.xref_xml_metadata())
                except Exception:
                    has_xml_metadata = False
                page_count = len(out_doc)
                page_size = None
                if page_count:
                    page = out_doc[0]
                    page_size = {
                        "width": float(page.rect.width),
                        "height": float(page.rect.height),
                    }
                out_doc.close()

            return {
                "success": True,
                "file": out_path,
                "input_file": abs_path,
                "output_file": out_path,
                "clear_metadata": clear_metadata,
                "remove_xml_metadata": remove_xml_metadata,
                "before": before,
                "after": after,
                "has_xml_metadata": has_xml_metadata,
                "pages": page_count,
                "page_size": page_size,
                "backup": backup or None,
            }
        except Exception as exc:
            return {"success": False, "error": str(exc)}
