#!/usr/bin/env python3
"""PDF operations for the OnlyOffice CLI."""

from __future__ import annotations

import io
import math
import os
import re
import tempfile
import warnings
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple


class PDFOperations:
    """Encapsulate PDF inspection, extraction, rendering, and sanitization."""

    MIN_RENDER_DPI = 36
    MAX_RENDER_DPI = 600
    MAX_RENDER_PAGES = 50
    MAX_RENDER_PIXELS_PER_PAGE = 64_000_000
    MAX_RENDER_TOTAL_PIXELS = 512_000_000
    MAX_EXTRACT_PAGES = 250
    MAX_EXTRACT_IMAGES = 1_000
    MAX_EXTRACT_IMAGE_COMPRESSED_BYTES = 50 * 1024 * 1024
    MAX_EXTRACT_IMAGE_PIXELS = 64_000_000
    MAX_EXTRACT_TOTAL_COMPRESSED_BYTES = 250 * 1024 * 1024
    MAX_EXTRACT_TOTAL_PIXELS = 512_000_000
    MAX_READ_PAGES = 250
    MAX_READ_FILE_SIZE_BYTES = 512 * 1024 * 1024
    MAX_READ_BLOCKS = 10_000
    MAX_READ_SPANS = 100_000
    MAX_READ_TEXT_CHARS = 2_000_000
    MAX_MUTATE_PAGES = 1_000
    MAX_MERGE_INPUTS = 100
    MAX_STAMP_TEXT_CHARS = 5_000
    MAX_STAMP_IMAGE_BYTES = 50 * 1024 * 1024
    MAX_REDACTION_MATCHES = 1_000
    IMAGE_OUTPUT_FORMATS = {
        "png": ("png", "PNG"),
        "jpg": ("jpg", "JPEG"),
        "jpeg": ("jpg", "JPEG"),
    }

    def __init__(self, host: Any):
        self.host = host

    @staticmethod
    def parse_page_range(
        pages_str: str,
        total: int,
        max_pages: int = None,
    ) -> List[int]:
        """Parse a page range string like ``0-3`` or ``1,3,5``."""
        total = max(0, int(total))
        max_pages = None if max_pages is None else max(0, int(max_pages))

        def ensure_capacity(count: int) -> None:
            if max_pages is not None and count > max_pages:
                raise ValueError(
                    f"Page selection contains {count} pages, exceeding the safety limit of {max_pages}"
                )

        def range_label() -> str:
            return "no pages" if total == 0 else f"0-{total - 1}"

        if not pages_str:
            ensure_capacity(total)
            return list(range(total))
        result = set()
        for part in pages_str.split(","):
            part = part.strip()
            if not part:
                raise ValueError(
                    f"Invalid pages component {part!r}; use zero-based pages like 0,2-4"
                )
            if "-" in part:
                if part.count("-") != 1:
                    raise ValueError(
                        f"Invalid pages component {part!r}; use zero-based pages like 0,2-4"
                    )
                start, end = part.split("-", 1)
                if not start.strip().isdigit() or not end.strip().isdigit():
                    raise ValueError(
                        f"Invalid pages component {part!r}; use zero-based pages like 0,2-4"
                    )
                start = int(start)
                end = int(end)
                if start > end:
                    raise ValueError(
                        f"Invalid pages component {part!r}; range start must be <= end"
                    )
                if start < 0 or end >= total:
                    raise ValueError(
                        f"Pages component {part!r} is out of range for PDF pages {range_label()}"
                    )
                ensure_capacity(len(result) + (end - start + 1))
                for index in range(start, end + 1):
                    result.add(index)
                ensure_capacity(len(result))
            else:
                if not part.isdigit():
                    raise ValueError(
                        f"Invalid pages component {part!r}; use zero-based pages like 0,2-4"
                    )
                index = int(part)
                if index < 0 or index >= total:
                    raise ValueError(
                        f"Page {index} is out of range for PDF pages {range_label()}"
                    )
                result.add(index)
                ensure_capacity(len(result))
        return sorted(set(result))

    @staticmethod
    def parse_ordered_page_range(
        pages_str: str,
        total: int,
        max_pages: int = None,
    ) -> List[int]:
        """Parse zero-based page specs while preserving order and duplicates."""
        total = max(0, int(total))
        max_pages = None if max_pages is None else max(0, int(max_pages))

        def ensure_capacity(count: int) -> None:
            if max_pages is not None and count > max_pages:
                raise ValueError(
                    f"Page selection contains {count} pages, exceeding the safety limit of {max_pages}"
                )

        def range_label() -> str:
            return "no pages" if total == 0 else f"0-{total - 1}"

        if not pages_str:
            ensure_capacity(total)
            return list(range(total))
        result: List[int] = []
        for part in str(pages_str).split(","):
            part = part.strip()
            if not part:
                raise ValueError(
                    f"Invalid pages component {part!r}; use zero-based pages like 2,0,1 or 0-3"
                )
            if "-" in part:
                if part.count("-") != 1:
                    raise ValueError(
                        f"Invalid pages component {part!r}; use zero-based pages like 2,0,1 or 0-3"
                    )
                start_text, end_text = part.split("-", 1)
                if not start_text.strip().isdigit() or not end_text.strip().isdigit():
                    raise ValueError(
                        f"Invalid pages component {part!r}; use zero-based pages like 2,0,1 or 0-3"
                    )
                start = int(start_text)
                end = int(end_text)
                if start > end:
                    raise ValueError(
                        f"Invalid pages component {part!r}; range start must be <= end"
                    )
                if start < 0 or end >= total:
                    raise ValueError(
                        f"Pages component {part!r} is out of range for PDF pages {range_label()}"
                    )
                ensure_capacity(len(result) + (end - start + 1))
                result.extend(range(start, end + 1))
            else:
                if not part.isdigit():
                    raise ValueError(
                        f"Invalid pages component {part!r}; use zero-based pages like 2,0,1 or 0-3"
                    )
                index = int(part)
                if index < 0 or index >= total:
                    raise ValueError(
                        f"Page {index} is out of range for PDF pages {range_label()}"
                    )
                ensure_capacity(len(result) + 1)
                result.append(index)
        return result

    @staticmethod
    def _coerce_bool(value: bool) -> bool:
        return bool(value)

    @staticmethod
    def _safe_pdf_prefix(value: str) -> str:
        prefix = re.sub(r"[^A-Za-z0-9_.-]+", "_", str(value or "page")).strip("._-")
        if not prefix:
            raise ValueError("Unsafe or empty output prefix")
        return prefix[:80]

    @staticmethod
    def _pdf_header_ok(path: str) -> bool:
        try:
            with open(path, "rb") as handle:
                return handle.read(5) == b"%PDF-"
        except OSError:
            return False

    @staticmethod
    def _fsync_file(path: str) -> None:
        try:
            fd = os.open(path, os.O_RDONLY)
        except OSError:
            return
        try:
            os.fsync(fd)
        finally:
            os.close(fd)

    @staticmethod
    def _hex_to_rgb(value: str) -> Tuple[float, float, float]:
        text = str(value or "000000").strip().lstrip("#")
        if len(text) != 6 or not re.fullmatch(r"[0-9A-Fa-f]{6}", text):
            raise ValueError("Color must be a 6-digit hex value such as 000000 or #FF0000")
        return (
            int(text[0:2], 16) / 255.0,
            int(text[2:4], 16) / 255.0,
            int(text[4:6], 16) / 255.0,
        )

    @staticmethod
    def _page_index(value: int, page_count: int) -> int:
        index = int(value)
        if index < 0 or index >= page_count:
            raise ValueError(
                f"Page {index} is out of range for PDF pages "
                f"{'no pages' if page_count == 0 else f'0-{page_count - 1}'}"
            )
        return index

    @staticmethod
    def _rect_payload(rect: Any) -> Dict[str, float]:
        return {
            "left": round(float(rect.x0), 3),
            "top": round(float(rect.y0), 3),
            "right": round(float(rect.x1), 3),
            "bottom": round(float(rect.y1), 3),
            "width": round(float(rect.width), 3),
            "height": round(float(rect.height), 3),
        }

    @classmethod
    def _rect_from_points(
        cls,
        fitz: Any,
        page: Any,
        left: float,
        top: float,
        right: float,
        bottom: float,
    ) -> Any:
        values = [float(left), float(top), float(right), float(bottom)]
        if not all(math.isfinite(value) for value in values):
            raise ValueError("Rectangle coordinates must be finite numbers")
        rect = fitz.Rect(*values)
        if rect.is_empty or rect.width <= 0 or rect.height <= 0:
            raise ValueError("Rectangle must have positive width and height")
        page_rect = page.rect
        if (
            rect.x0 < page_rect.x0
            or rect.y0 < page_rect.y0
            or rect.x1 > page_rect.x1
            or rect.y1 > page_rect.y1
        ):
            raise ValueError("Rectangle must be inside the target page bounds")
        return rect

    @classmethod
    def _rect_from_position(
        cls,
        fitz: Any,
        page: Any,
        x: float,
        y: float,
        width: float,
        height: float,
    ) -> Any:
        x = float(x)
        y = float(y)
        width = float(width)
        height = float(height)
        return cls._rect_from_points(fitz, page, x, y, x + width, y + height)

    @staticmethod
    def _save_options(
        *,
        garbage: int = 4,
        clean: bool = True,
        deflate: bool = False,
        linearize: bool = False,
    ) -> Dict[str, Any]:
        garbage = int(garbage)
        if garbage < 0 or garbage > 4:
            raise ValueError("--garbage must be between 0 and 4")
        return {
            "garbage": garbage,
            "clean": bool(clean),
            "deflate": bool(deflate),
            "deflate_images": bool(deflate),
            "deflate_fonts": bool(deflate),
            "linear": bool(linearize),
            "use_objstms": 0 if linearize else 1,
        }

    def _save_pdf_atomic_unlocked(
        self,
        fitz: Any,
        pdf: Any,
        out_path: str,
        *,
        expected_pages: Optional[int] = None,
        save_options: Optional[Dict[str, Any]] = None,
    ) -> Dict[str, Any]:
        target = Path(out_path).expanduser().resolve()
        target.parent.mkdir(parents=True, exist_ok=True)
        fd, temp_output = tempfile.mkstemp(
            prefix=f".{target.name}.",
            suffix=".tmp.pdf",
            dir=str(target.parent),
        )
        os.close(fd)
        backup = None
        try:
            pdf.save(temp_output, **(save_options or self._save_options()))
            if not os.path.exists(temp_output) or os.path.getsize(temp_output) <= 0:
                return {"success": False, "error": "PDF save produced no output"}
            if not self._pdf_header_ok(temp_output):
                return {"success": False, "error": "PDF save output is not a PDF"}
            self._fsync_file(temp_output)
            probe = fitz.open(temp_output)
            try:
                saved_pages = len(probe)
            finally:
                probe.close()
            if expected_pages is not None and saved_pages != expected_pages:
                return {
                    "success": False,
                    "error": (
                        f"PDF save page count mismatch: expected {expected_pages}, got {saved_pages}"
                    ),
                }
            if target.exists():
                backup = self.host._snapshot_backup(str(target))
            os.replace(temp_output, str(target))
            temp_output = None
            self.host._fsync_directory(target.parent)
            return {
                "success": True,
                "output_file": str(target),
                "file_size": os.path.getsize(target),
                "pages": saved_pages,
                "backup": backup or None,
            }
        finally:
            if temp_output and os.path.exists(temp_output):
                try:
                    os.unlink(temp_output)
                except OSError:
                    pass

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

    @classmethod
    def normalize_image_format(cls, fmt: str) -> Any:
        fmt_key = str(fmt or "").strip().lower().lstrip(".")
        normalized = cls.IMAGE_OUTPUT_FORMATS.get(fmt_key)
        if normalized:
            return normalized, None
        return None, {
            "success": False,
            "error": "Unsupported image format. Use png or jpg.",
            "error_code": "unsupported_image_format",
            "allowed_formats": ["png", "jpg"],
        }

    @staticmethod
    def prepare_output_dir(output_dir: str) -> Path:
        output_root = Path(output_dir).expanduser()
        output_root.mkdir(parents=True, exist_ok=True)
        return output_root.resolve()

    @staticmethod
    def safe_child_path(output_root: Path, filename: str) -> Path:
        candidate = (output_root / filename).resolve()
        try:
            common = os.path.commonpath([str(output_root), str(candidate)])
        except ValueError:
            common = ""
        if common != str(output_root):
            raise ValueError("Refusing to write outside output_dir")
        return candidate

    @classmethod
    def image_extract_resource_limits(cls) -> Dict[str, int]:
        return {
            "max_images": cls.MAX_EXTRACT_IMAGES,
            "max_compressed_image_bytes": cls.MAX_EXTRACT_IMAGE_COMPRESSED_BYTES,
            "max_decoded_image_pixels": cls.MAX_EXTRACT_IMAGE_PIXELS,
            "max_total_compressed_image_bytes": cls.MAX_EXTRACT_TOTAL_COMPRESSED_BYTES,
            "max_total_decoded_image_pixels": cls.MAX_EXTRACT_TOTAL_PIXELS,
        }

    @classmethod
    def load_bounded_image(
        cls,
        PILImage: Any,
        image_bytes: bytes,
        *,
        total_compressed_bytes: int = 0,
        total_decoded_pixels: int = 0,
    ) -> Any:
        compressed_size = len(image_bytes)
        compressed_after = total_compressed_bytes + compressed_size
        if compressed_after > cls.MAX_EXTRACT_TOTAL_COMPRESSED_BYTES:
            return None, None, {
                "skipped": True,
                "stop_extraction": True,
                "error": (
                    f"Embedded images would total {compressed_after} compressed bytes, exceeding "
                    f"the safe limit {cls.MAX_EXTRACT_TOTAL_COMPRESSED_BYTES}"
                ),
                "error_code": "aggregate_image_compressed_bytes_limit_exceeded",
                "compressed_size_bytes": compressed_size,
                "total_compressed_size_bytes": total_compressed_bytes,
                "would_total_compressed_size_bytes": compressed_after,
                "max_total_compressed_size_bytes": cls.MAX_EXTRACT_TOTAL_COMPRESSED_BYTES,
            }
        if compressed_size > cls.MAX_EXTRACT_IMAGE_COMPRESSED_BYTES:
            return None, None, {
                "skipped": True,
                "error": (
                    f"Embedded image is {compressed_size} compressed bytes, exceeding "
                    f"the safe limit {cls.MAX_EXTRACT_IMAGE_COMPRESSED_BYTES}"
                ),
                "error_code": "image_compressed_bytes_limit_exceeded",
                "compressed_size_bytes": compressed_size,
                "max_compressed_size_bytes": cls.MAX_EXTRACT_IMAGE_COMPRESSED_BYTES,
                "total_compressed_size_bytes": compressed_after,
            }

        img = None
        try:
            stream = io.BytesIO(image_bytes)
            bomb_warning = getattr(PILImage, "DecompressionBombWarning", None)
            with warnings.catch_warnings():
                if bomb_warning is not None:
                    warnings.simplefilter("error", bomb_warning)
                img = PILImage.open(stream)
                width, height = [int(value) for value in img.size]
                pixels = width * height
                pixels_after = total_decoded_pixels + pixels
                if pixels_after > cls.MAX_EXTRACT_TOTAL_PIXELS:
                    img.close()
                    return None, None, {
                        "skipped": True,
                        "stop_extraction": True,
                        "error": (
                            f"Embedded images would decode to {pixels_after} pixels, exceeding "
                            f"the safe limit {cls.MAX_EXTRACT_TOTAL_PIXELS}"
                        ),
                        "error_code": "aggregate_image_pixel_limit_exceeded",
                        "width": width,
                        "height": height,
                        "pixels": pixels,
                        "total_decoded_image_pixels": total_decoded_pixels,
                        "would_total_decoded_image_pixels": pixels_after,
                        "max_total_decoded_image_pixels": cls.MAX_EXTRACT_TOTAL_PIXELS,
                        "compressed_size_bytes": compressed_size,
                        "total_compressed_size_bytes": compressed_after,
                    }
                if pixels > cls.MAX_EXTRACT_IMAGE_PIXELS:
                    img.close()
                    return None, None, {
                        "skipped": True,
                        "error": (
                            f"Embedded image decodes to {pixels} pixels, exceeding "
                            f"the safe limit {cls.MAX_EXTRACT_IMAGE_PIXELS}"
                        ),
                        "error_code": "image_pixel_limit_exceeded",
                        "width": width,
                        "height": height,
                        "pixels": pixels,
                        "max_decoded_image_pixels": cls.MAX_EXTRACT_IMAGE_PIXELS,
                        "compressed_size_bytes": compressed_size,
                        "total_compressed_size_bytes": compressed_after,
                        "total_decoded_image_pixels": pixels_after,
                    }
                img.load()
            return img, {
                "width": width,
                "height": height,
                "pixels": pixels,
                "compressed_size_bytes": compressed_size,
                "total_compressed_size_bytes": compressed_after,
                "total_decoded_image_pixels": pixels_after,
            }, None
        except Exception as exc:
            if img is not None:
                img.close()
            return None, None, {
                "skipped": True,
                "error": f"Could not safely decode embedded image: {exc}",
                "error_code": "unsafe_image_decode",
                "compressed_size_bytes": compressed_size,
                "total_compressed_size_bytes": compressed_after,
            }

    @staticmethod
    def save_bounded_image(
        img: Any,
        out_path: Path,
        fmt_ext: str,
        pillow_format: str,
    ) -> None:
        tmp_path = out_path.with_name(f".{out_path.name}.tmp")
        try:
            if fmt_ext == "jpg":
                converted = img.convert("RGB")
                try:
                    converted.save(str(tmp_path), "JPEG", quality=90)
                finally:
                    converted.close()
            else:
                img.save(str(tmp_path), pillow_format)
            os.replace(str(tmp_path), str(out_path))
        except Exception:
            try:
                tmp_path.unlink()
            except FileNotFoundError:
                pass
            raise

    @staticmethod
    def save_pixmap_atomic(pix: Any, out_path: Path, fmt_ext: str) -> None:
        fd, temp_output = tempfile.mkstemp(
            prefix=f".{out_path.stem}.",
            suffix=f".{fmt_ext}",
            dir=str(out_path.parent),
        )
        os.close(fd)
        try:
            if fmt_ext == "jpg":
                pix.pil_save(temp_output, format="JPEG", quality=90)
            else:
                pix.save(temp_output)
            os.replace(temp_output, str(out_path))
        finally:
            if os.path.exists(temp_output):
                os.unlink(temp_output)

    @classmethod
    def page_range_error_result(
        cls,
        exc: Exception,
        total_pages: int,
        max_pages: int,
        operation: str,
    ) -> Dict[str, Any]:
        message = str(exc)
        limit_error = "safety limit" in message
        return {
            "success": False,
            "error": message,
            "error_code": "unsafe_pdf_resource_request" if limit_error else "invalid_page_range",
            "preflight": {
                "operation": operation,
                "status": "fail",
                "total_pages": total_pages,
                "max_pages": max_pages,
                "reason": message,
            },
        }

    @classmethod
    def page_selection_preflight(
        cls,
        page_indices: List[int],
        total_pages: int,
        max_pages: int,
        operation: str,
    ) -> Dict[str, Any]:
        report = {
            "operation": operation,
            "status": "pass",
            "total_pages": total_pages,
            "pages_requested": len(page_indices),
            "max_pages": max_pages,
            "selected_pages_sample": page_indices[:20],
        }
        if len(page_indices) > max_pages:
            report["status"] = "fail"
            report["reason"] = (
                f"{operation} selected {len(page_indices)} pages, exceeding the "
                f"safety limit of {max_pages}. Pass --pages with a smaller range."
            )
        return report

    @classmethod
    def render_preflight(cls, doc: Any, page_indices: List[int], dpi: int) -> Any:
        report = {
            "operation": "pdf_page_to_image",
            "status": "pass",
            "dpi": dpi,
            "min_dpi": cls.MIN_RENDER_DPI,
            "max_dpi": cls.MAX_RENDER_DPI,
            "pages_requested": len(page_indices),
            "max_pages": cls.MAX_RENDER_PAGES,
            "max_pixels_per_page": cls.MAX_RENDER_PIXELS_PER_PAGE,
            "max_total_pixels": cls.MAX_RENDER_TOTAL_PIXELS,
            "page_estimates": [],
        }
        reasons = []
        if isinstance(dpi, bool):
            reasons.append("dpi must be an integer")
        else:
            try:
                dpi = int(dpi)
            except (TypeError, ValueError):
                reasons.append("dpi must be an integer")
        report["dpi"] = dpi

        if not reasons and (dpi < cls.MIN_RENDER_DPI or dpi > cls.MAX_RENDER_DPI):
            reasons.append(
                f"dpi {dpi} is outside the safe range {cls.MIN_RENDER_DPI}-{cls.MAX_RENDER_DPI}"
            )
        if len(page_indices) > cls.MAX_RENDER_PAGES:
            reasons.append(
                f"{len(page_indices)} pages selected; max safe render pages is {cls.MAX_RENDER_PAGES}"
            )

        estimated_total = 0
        if not reasons:
            for page_index in page_indices:
                page = doc[page_index]
                width_px = max(1, int(math.ceil(float(page.rect.width) * dpi / 72.0)))
                height_px = max(1, int(math.ceil(float(page.rect.height) * dpi / 72.0)))
                pixels = width_px * height_px
                estimated_total += pixels
                if len(report["page_estimates"]) < 20:
                    report["page_estimates"].append(
                        {
                            "page": page_index,
                            "width_px": width_px,
                            "height_px": height_px,
                            "pixels": pixels,
                        }
                    )
                if pixels > cls.MAX_RENDER_PIXELS_PER_PAGE:
                    reasons.append(
                        f"page {page_index} would render {pixels} pixels, exceeding "
                        f"the per-page limit {cls.MAX_RENDER_PIXELS_PER_PAGE}"
                    )
                    break
            if estimated_total > cls.MAX_RENDER_TOTAL_PIXELS:
                reasons.append(
                    f"selected pages would render {estimated_total} pixels total, "
                    f"exceeding the limit {cls.MAX_RENDER_TOTAL_PIXELS}"
                )
        report["estimated_total_pixels"] = estimated_total

        if reasons:
            report["status"] = "fail"
            report["reasons"] = reasons
            return None, report, {
                "success": False,
                "error": "Unsafe PDF render request",
                "error_code": "unsafe_pdf_render_request",
                "preflight": report,
            }
        return dpi, report, None

    @staticmethod
    def hidden_data_summary(pdf: Any) -> Dict[str, Any]:
        has_xml_metadata = False
        try:
            has_xml_metadata = bool(pdf.xref_xml_metadata())
        except Exception:
            has_xml_metadata = False

        annotations_total = 0
        annotation_types: Dict[str, int] = {}
        for page in pdf:
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

        return {
            "has_xml_metadata": has_xml_metadata,
            "annotations_count": annotations_total,
            "annotation_types": annotation_types,
            "embedded_files_count": embedded_files,
            "has_embedded_files": embedded_files > 0,
            "has_forms": form_fields,
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
            fmt_info, fmt_error = self.normalize_image_format(fmt)
            if fmt_error:
                return fmt_error
            fmt_ext, pillow_format = fmt_info
            output_root = self.prepare_output_dir(output_dir)
            doc = fitz.open(file_path)
            total_pages = len(doc)
            try:
                page_indices = self.parse_page_range(
                    pages,
                    total_pages,
                    max_pages=self.MAX_EXTRACT_PAGES,
                )
            except ValueError as exc:
                doc.close()
                return self.page_range_error_result(
                    exc,
                    total_pages,
                    self.MAX_EXTRACT_PAGES,
                    "pdf_extract_images",
                )
            preflight = self.page_selection_preflight(
                page_indices,
                total_pages,
                self.MAX_EXTRACT_PAGES,
                "pdf_extract_images",
            )
            resource_limits = self.image_extract_resource_limits()
            preflight.update(resource_limits)
            if preflight["status"] != "pass":
                doc.close()
                return {
                    "success": False,
                    "error": preflight["reason"],
                    "error_code": "unsafe_pdf_resource_request",
                    "preflight": preflight,
                }
            extracted = []
            seen_xrefs = set()
            index = 0
            truncated = False
            warnings_list = []
            total_compressed_bytes = 0
            total_decoded_pixels = 0
            for page_index in page_indices:
                page = doc[page_index]
                for img_info in page.get_images(full=True):
                    xref = img_info[0]
                    if xref in seen_xrefs:
                        continue
                    if index >= self.MAX_EXTRACT_IMAGES:
                        truncated = True
                        warning = f"Stopped after {self.MAX_EXTRACT_IMAGES} images"
                        if warning not in warnings_list:
                            warnings_list.append(warning)
                        break
                    seen_xrefs.add(xref)
                    try:
                        base_image = doc.extract_image(xref)
                        if not base_image:
                            continue
                        image_bytes = base_image["image"]
                        img_ext = base_image.get("ext", "png")
                        out_name = f"pdf_img_{page_index:03d}_{index:03d}.{fmt_ext}"
                        out_path = self.safe_child_path(output_root, out_name)
                        img, image_meta, skip_entry = self.load_bounded_image(
                            PILImage,
                            image_bytes,
                            total_compressed_bytes=total_compressed_bytes,
                            total_decoded_pixels=total_decoded_pixels,
                        )
                        stop_after_current = False
                        if skip_entry:
                            if (
                                skip_entry.get("error_code")
                                != "aggregate_image_compressed_bytes_limit_exceeded"
                            ):
                                total_compressed_bytes = skip_entry.get(
                                    "total_compressed_size_bytes",
                                    total_compressed_bytes,
                                )
                            if (
                                skip_entry.get("error_code")
                                != "aggregate_image_pixel_limit_exceeded"
                            ):
                                total_decoded_pixels = skip_entry.get(
                                    "total_decoded_image_pixels",
                                    total_decoded_pixels,
                                )
                            stop_after_current = bool(skip_entry.get("stop_extraction"))
                            if stop_after_current:
                                truncated = True
                                warning = str(skip_entry.get("error") or "Stopped at image extraction budget")
                                if warning not in warnings_list:
                                    warnings_list.append(warning)
                            skip_entry.update(
                                {
                                    "index": index,
                                    "page": page_index,
                                    "xref": xref,
                                    "original_format": img_ext,
                                }
                            )
                            extracted.append(skip_entry)
                        else:
                            total_compressed_bytes = image_meta["total_compressed_size_bytes"]
                            total_decoded_pixels = image_meta["total_decoded_image_pixels"]
                            try:
                                self.save_bounded_image(img, out_path, fmt_ext, pillow_format)
                            finally:
                                img.close()
                            extracted.append(
                                {
                                    "index": index,
                                    "page": page_index,
                                    "xref": xref,
                                    "file": str(out_path),
                                    "format": fmt_ext,
                                    "width": image_meta["width"],
                                    "height": image_meta["height"],
                                    "pixels": image_meta["pixels"],
                                    "compressed_size_bytes": image_meta["compressed_size_bytes"],
                                    "size_bytes": os.path.getsize(str(out_path)),
                                    "original_format": img_ext,
                                }
                            )
                        if stop_after_current:
                            index += 1
                            break
                    except Exception as img_err:
                        extracted.append(
                            {
                                "index": index,
                                "page": page_index,
                                "xref": xref,
                                "skipped": True,
                                "error": str(img_err),
                                "error_code": "image_extract_failed",
                            }
                        )
                    index += 1
                if truncated:
                    break
            doc.close()
            return {
                "success": True,
                "file": file_path,
                "output_dir": str(output_root),
                "total_pages": total_pages,
                "pages_scanned": len(page_indices),
                "images_extracted": len([entry for entry in extracted if "file" in entry]),
                "images_skipped": len([entry for entry in extracted if entry.get("skipped")]),
                "images": extracted,
                "truncated": truncated,
                "warnings": warnings_list,
                "resource_limits": resource_limits,
                "resource_usage": {
                    "compressed_image_bytes": total_compressed_bytes,
                    "decoded_image_pixels": total_decoded_pixels,
                },
                "preflight": preflight,
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
            abs_path = Path(file_path).resolve()
            file_size = abs_path.stat().st_size
            if file_size > self.MAX_READ_FILE_SIZE_BYTES:
                return {
                    "success": False,
                    "error": (
                        f"PDF is {file_size} bytes, exceeding the safe read limit "
                        f"{self.MAX_READ_FILE_SIZE_BYTES}"
                    ),
                    "error_code": "unsafe_pdf_resource_request",
                    "preflight": {
                        "operation": "pdf_read_blocks",
                        "status": "fail",
                        "file_size_bytes": file_size,
                        "max_file_size_bytes": self.MAX_READ_FILE_SIZE_BYTES,
                    },
                }
            doc = fitz.open(str(abs_path))
            total_pages = len(doc)
            try:
                page_indices = self.parse_page_range(
                    pages,
                    total_pages,
                    max_pages=self.MAX_READ_PAGES,
                )
            except ValueError as exc:
                doc.close()
                return self.page_range_error_result(
                    exc,
                    total_pages,
                    self.MAX_READ_PAGES,
                    "pdf_read_blocks",
                )
            preflight = self.page_selection_preflight(
                page_indices,
                total_pages,
                self.MAX_READ_PAGES,
                "pdf_read_blocks",
            )
            preflight["file_size_bytes"] = file_size
            preflight["max_file_size_bytes"] = self.MAX_READ_FILE_SIZE_BYTES
            if preflight["status"] != "pass":
                doc.close()
                return {
                    "success": False,
                    "error": preflight["reason"],
                    "error_code": "unsafe_pdf_resource_request",
                    "preflight": preflight,
                }
            pages_payload = []
            text_block_count = 0
            image_block_count = 0
            span_count = 0
            text_char_count = 0
            truncated = False
            warnings = []
            resource_limits = {
                "max_pages": self.MAX_READ_PAGES,
                "max_file_size_bytes": self.MAX_READ_FILE_SIZE_BYTES,
                "max_blocks": self.MAX_READ_BLOCKS,
                "max_spans": self.MAX_READ_SPANS,
                "max_text_chars": self.MAX_READ_TEXT_CHARS,
            }

            def mark_truncated(reason: str) -> None:
                nonlocal truncated
                truncated = True
                if reason not in warnings:
                    warnings.append(reason)

            for page_index in page_indices:
                if truncated:
                    break
                page = doc[page_index]
                page_dict = page.get_text("dict", sort=True)
                page_blocks = []

                for block_index, block in enumerate(page_dict.get("blocks", [])):
                    if truncated:
                        break
                    if text_block_count + image_block_count >= self.MAX_READ_BLOCKS:
                        mark_truncated(
                            f"Stopped after {self.MAX_READ_BLOCKS} output blocks"
                        )
                        break
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
                                if span_count >= self.MAX_READ_SPANS:
                                    mark_truncated(
                                        f"Stopped after {self.MAX_READ_SPANS} text spans"
                                    )
                                    break
                                if text_char_count >= self.MAX_READ_TEXT_CHARS:
                                    mark_truncated(
                                        f"Stopped after {self.MAX_READ_TEXT_CHARS} text characters"
                                    )
                                    break
                                span_text = str(span.get("text", "") or "")
                                if not include_empty and not span_text.strip():
                                    continue
                                remaining_chars = self.MAX_READ_TEXT_CHARS - text_char_count
                                if len(span_text) > remaining_chars:
                                    span_text = span_text[:remaining_chars]
                                    mark_truncated(
                                        f"Stopped after {self.MAX_READ_TEXT_CHARS} text characters"
                                    )
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
                                text_char_count += len(span_text)
                                line_spans.append(span_payload)
                                if span_text.strip():
                                    line_text_parts.append(span_text)
                                if truncated:
                                    break

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
                            if truncated:
                                break

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
                "text_char_count": text_char_count,
                "truncated": truncated,
                "warnings": warnings,
                "resource_limits": resource_limits,
                "preflight": preflight,
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
            "truncated": payload.get("truncated", False),
            "warnings": payload.get("warnings", []),
            "resource_limits": payload.get("resource_limits"),
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
            fmt_info, fmt_error = self.normalize_image_format(fmt)
            if fmt_error:
                return fmt_error
            fmt_ext, _ = fmt_info
            output_root = self.prepare_output_dir(output_dir)
            doc = fitz.open(file_path)
            total_pages = len(doc)
            try:
                page_indices = self.parse_page_range(
                    pages,
                    total_pages,
                    max_pages=self.MAX_RENDER_PAGES,
                )
            except ValueError as exc:
                doc.close()
                return self.page_range_error_result(
                    exc,
                    total_pages,
                    self.MAX_RENDER_PAGES,
                    "pdf_page_to_image",
                )
            dpi, preflight, preflight_error = self.render_preflight(doc, page_indices, dpi)
            if preflight_error:
                doc.close()
                return preflight_error
            rendered = []
            for page_index in page_indices:
                page = doc[page_index]
                pix = page.get_pixmap(dpi=dpi)
                if fmt_ext == "jpg":
                    out_name = f"page_{page_index:03d}.jpg"
                    out_path = self.safe_child_path(output_root, out_name)
                else:
                    out_name = f"page_{page_index:03d}.png"
                    out_path = self.safe_child_path(output_root, out_name)
                self.save_pixmap_atomic(pix, out_path, fmt_ext)
                rendered.append(
                    {
                        "page": page_index,
                        "file": str(out_path),
                        "format": fmt_ext,
                        "width": pix.width,
                        "height": pix.height,
                        "dpi": dpi,
                        "size_bytes": os.path.getsize(str(out_path)),
                    }
                )
            doc.close()
            return {
                "success": True,
                "file": file_path,
                "output_dir": str(output_root),
                "total_pages": total_pages,
                "pages_rendered": len(rendered),
                "dpi": dpi,
                "images": rendered,
                "preflight": preflight,
            }
        except Exception as exc:
            return {"success": False, "error": str(exc)}

    @staticmethod
    def _atomic_save_pil_image(image: Any, output_path: str, image_format: str) -> None:
        target = Path(output_path).expanduser().resolve()
        target.parent.mkdir(parents=True, exist_ok=True)
        fd, temp_path = tempfile.mkstemp(
            prefix=f".{target.name}.",
            suffix=f".{target.suffix or '.image'}.tmp",
            dir=str(target.parent),
        )
        os.close(fd)
        try:
            image.save(temp_path, format=image_format)
            PDFOperations._fsync_file(temp_path)
            os.replace(temp_path, str(target))
            try:
                fd_dir = os.open(str(target.parent), os.O_RDONLY)
            except OSError:
                fd_dir = None
            if fd_dir is not None:
                try:
                    os.fsync(fd_dir)
                finally:
                    os.close(fd_dir)
        finally:
            if os.path.exists(temp_path):
                try:
                    os.unlink(temp_path)
                except OSError:
                    pass

    def map_page(
        self,
        file_path: str,
        page_index: int,
        output_image: str,
        *,
        dpi: int = 150,
        fmt: str = "png",
        labels: bool = True,
        include_images: bool = True,
    ) -> Dict[str, Any]:
        """Render one PDF page with visible block bounding boxes and block ids."""
        try:
            import fitz
            from PIL import Image as PILImage
            from PIL import ImageDraw, ImageFont
        except ImportError:
            return {"success": False, "error": "PyMuPDF and Pillow are required for PDF page mapping"}

        try:
            fmt_info, fmt_error = self.normalize_image_format(fmt)
            if fmt_error:
                return fmt_error
            fmt_ext, pil_format = fmt_info
            abs_path = str(Path(file_path).expanduser().resolve())
            out_path = str(Path(output_image).expanduser().resolve())
            if not os.path.exists(abs_path):
                return {"success": False, "error": f"File not found: {file_path}"}
            doc = fitz.open(abs_path)
            try:
                page_number = self._page_index(page_index, len(doc))
                dpi, preflight, preflight_error = self.render_preflight(
                    doc,
                    [page_number],
                    dpi,
                )
                if preflight_error:
                    return preflight_error
                page = doc[page_number]
                pix = page.get_pixmap(dpi=dpi, alpha=False)
                if pix.n >= 4:
                    mode = "RGBA"
                else:
                    mode = "RGB"
                image = PILImage.frombytes(mode, (pix.width, pix.height), pix.samples)
            finally:
                doc.close()

            block_payload = self.read_blocks(
                abs_path,
                pages=str(page_number),
                include_spans=False,
                include_images=include_images,
            )
            if not block_payload.get("success"):
                return block_payload
            page_payload = block_payload.get("pages", [{}])[0]
            blocks = page_payload.get("blocks", [])
            draw = ImageDraw.Draw(image)
            scale = float(dpi) / 72.0
            font = ImageFont.load_default()
            for block in blocks:
                bbox = block.get("bbox") or {}
                if not bbox:
                    continue
                left = float(bbox.get("left", 0.0)) * scale
                top = float(bbox.get("top", 0.0)) * scale
                right = float(bbox.get("right", 0.0)) * scale
                bottom = float(bbox.get("bottom", 0.0)) * scale
                color = (220, 38, 38) if block.get("type") == "text" else (37, 99, 235)
                draw.rectangle((left, top, right, bottom), outline=color, width=3)
                if labels:
                    label = str(block.get("block_id", "block"))
                    try:
                        label_box = draw.textbbox((left + 2, top + 2), label, font=font)
                        draw.rectangle(label_box, fill=(255, 255, 255))
                    except Exception:
                        pass
                    draw.text((left + 2, top + 2), label, fill=color, font=font)
            if fmt_ext == "jpg" and image.mode != "RGB":
                image = image.convert("RGB")
            self._atomic_save_pil_image(image, out_path, pil_format)
            return {
                "success": True,
                "file": abs_path,
                "output_image": out_path,
                "page": page_number,
                "dpi": dpi,
                "format": fmt_ext,
                "labels": bool(labels),
                "blocks_mapped": len(blocks),
                "blocks": blocks,
                "preflight": preflight,
                "size_bytes": os.path.getsize(out_path),
            }
        except Exception as exc:
            return {"success": False, "error": str(exc)}

    def compact(
        self,
        file_path: str,
        output_path: str = None,
        *,
        garbage: int = 4,
        deflate: bool = True,
        clean: bool = True,
        linearize: bool = False,
    ) -> Dict[str, Any]:
        """Explicitly compact a PDF. This is never applied by default."""
        try:
            import fitz
        except ImportError:
            return {"success": False, "error": "PyMuPDF (fitz) not installed. Run: pip install PyMuPDF"}

        try:
            abs_path = str(Path(file_path).expanduser().resolve())
            out_path = str(Path(output_path).expanduser().resolve()) if output_path else abs_path
            save_options = self._save_options(
                garbage=garbage,
                clean=clean,
                deflate=deflate,
                linearize=linearize,
            )
            if not os.path.exists(abs_path):
                return {"success": False, "error": f"File not found: {file_path}"}
            with self.host._file_locks(abs_path, out_path):
                before_size = os.path.getsize(abs_path)
                pdf = fitz.open(abs_path)
                try:
                    before_pages = len(pdf)
                    save = self._save_pdf_atomic_unlocked(
                        fitz,
                        pdf,
                        out_path,
                        expected_pages=before_pages,
                        save_options=save_options,
                    )
                finally:
                    pdf.close()
                if not save.get("success"):
                    return save
                after_size = int(save["file_size"])
            delta = after_size - before_size
            percent = round((delta / before_size) * 100.0, 3) if before_size else None
            return {
                "success": True,
                "file": save["output_file"],
                "input_file": abs_path,
                "output_file": save["output_file"],
                "operation": "pdf_compact",
                "compression_requested": True,
                "default_applied": False,
                "options": {
                    "garbage": int(garbage),
                    "deflate": bool(deflate),
                    "clean": bool(clean),
                    "linearize": bool(linearize),
                },
                "pages": save["pages"],
                "before_size": before_size,
                "after_size": after_size,
                "size_delta_bytes": delta,
                "size_delta_percent": percent,
                "smaller": after_size < before_size,
                "backup": save.get("backup"),
                "note": (
                    "Compaction completed, but output is not smaller; this can happen with already-compressed PDFs."
                    if after_size >= before_size
                    else None
                ),
            }
        except Exception as exc:
            return {"success": False, "error": str(exc)}

    def merge(
        self,
        input_files: List[str],
        output_path: str,
    ) -> Dict[str, Any]:
        """Merge multiple PDFs into one output PDF."""
        try:
            import fitz
        except ImportError:
            return {"success": False, "error": "PyMuPDF (fitz) not installed. Run: pip install PyMuPDF"}

        try:
            if len(input_files) < 2:
                return {"success": False, "error": "pdf-merge requires at least two input PDFs"}
            if len(input_files) > self.MAX_MERGE_INPUTS:
                return {
                    "success": False,
                    "error": f"pdf-merge supports at most {self.MAX_MERGE_INPUTS} input PDFs",
                    "error_code": "unsafe_pdf_resource_request",
                }
            inputs = [str(Path(path).expanduser().resolve()) for path in input_files]
            out_path = str(Path(output_path).expanduser().resolve())
            if out_path in inputs:
                return {
                    "success": False,
                    "error": "Merge output must not be the same path as an input PDF",
                    "error_code": "unsafe_pdf_output_path",
                }
            for path in inputs:
                if not os.path.exists(path):
                    return {"success": False, "error": f"File not found: {path}"}

            with self.host._file_locks(*inputs, out_path):
                output = fitz.open()
                page_map = []
                total_pages = 0
                try:
                    for input_index, path in enumerate(inputs):
                        src = fitz.open(path)
                        try:
                            page_count = len(src)
                            if total_pages + page_count > self.MAX_MUTATE_PAGES:
                                return {
                                    "success": False,
                                    "error": (
                                        f"Merged PDF would contain {total_pages + page_count} pages, exceeding "
                                        f"the safety limit {self.MAX_MUTATE_PAGES}"
                                    ),
                                    "error_code": "unsafe_pdf_resource_request",
                                }
                            start_at = len(output)
                            output.insert_pdf(src)
                            for source_page in range(page_count):
                                page_map.append(
                                    {
                                        "output_page": start_at + source_page,
                                        "input_index": input_index,
                                        "input_file": path,
                                        "input_page": source_page,
                                    }
                                )
                            total_pages += page_count
                        finally:
                            src.close()
                    save = self._save_pdf_atomic_unlocked(
                        fitz,
                        output,
                        out_path,
                        expected_pages=total_pages,
                    )
                finally:
                    output.close()
                if not save.get("success"):
                    return save
            return {
                "success": True,
                "file": save["output_file"],
                "output_file": save["output_file"],
                "input_files": inputs,
                "input_count": len(inputs),
                "pages": save["pages"],
                "page_map": page_map,
                "file_size": save["file_size"],
                "backup": save.get("backup"),
            }
        except Exception as exc:
            return {"success": False, "error": str(exc)}

    def split(
        self,
        file_path: str,
        output_dir: str,
        *,
        pages: str = None,
        prefix: str = "page",
    ) -> Dict[str, Any]:
        """Split selected PDF pages into one PDF per page."""
        try:
            import fitz
        except ImportError:
            return {"success": False, "error": "PyMuPDF (fitz) not installed. Run: pip install PyMuPDF"}

        try:
            abs_path = str(Path(file_path).expanduser().resolve())
            if not os.path.exists(abs_path):
                return {"success": False, "error": f"File not found: {file_path}"}
            output_root = self.prepare_output_dir(output_dir)
            safe_prefix = self._safe_pdf_prefix(prefix)
            src_probe = fitz.open(abs_path)
            try:
                total_pages = len(src_probe)
                page_indices = self.parse_page_range(
                    pages,
                    total_pages,
                    max_pages=self.MAX_MUTATE_PAGES,
                )
            finally:
                src_probe.close()
            output_paths = [
                str(self.safe_child_path(output_root, f"{safe_prefix}_{page_index:03d}.pdf"))
                for page_index in page_indices
            ]
            with self.host._file_locks(abs_path, *output_paths):
                src = fitz.open(abs_path)
                try:
                    outputs = []
                    for page_index, out_path in zip(page_indices, output_paths):
                        out_doc = fitz.open()
                        try:
                            out_doc.insert_pdf(src, from_page=page_index, to_page=page_index)
                            save = self._save_pdf_atomic_unlocked(
                                fitz,
                                out_doc,
                                out_path,
                                expected_pages=1,
                            )
                        finally:
                            out_doc.close()
                        if not save.get("success"):
                            return save
                        outputs.append(
                            {
                                "source_page": page_index,
                                "output_file": save["output_file"],
                                "file_size": save["file_size"],
                                "backup": save.get("backup"),
                            }
                        )
                finally:
                    src.close()
            return {
                "success": True,
                "file": abs_path,
                "output_dir": str(output_root),
                "pages_selected": len(page_indices),
                "total_pages": total_pages,
                "outputs": outputs,
            }
        except ValueError as exc:
            return self.page_range_error_result(
                exc,
                0,
                self.MAX_MUTATE_PAGES,
                "pdf_split",
            )
        except Exception as exc:
            return {"success": False, "error": str(exc)}

    def reorder(
        self,
        file_path: str,
        page_order: str,
        output_path: str = None,
    ) -> Dict[str, Any]:
        """Create a PDF with pages in an explicit user-provided order."""
        try:
            import fitz
        except ImportError:
            return {"success": False, "error": "PyMuPDF (fitz) not installed. Run: pip install PyMuPDF"}

        try:
            abs_path = str(Path(file_path).expanduser().resolve())
            out_path = str(Path(output_path).expanduser().resolve()) if output_path else abs_path
            if not os.path.exists(abs_path):
                return {"success": False, "error": f"File not found: {file_path}"}
            src_probe = fitz.open(abs_path)
            try:
                total_pages = len(src_probe)
                order = self.parse_ordered_page_range(
                    page_order,
                    total_pages,
                    max_pages=self.MAX_MUTATE_PAGES,
                )
            finally:
                src_probe.close()
            with self.host._file_locks(abs_path, out_path):
                src = fitz.open(abs_path)
                output = fitz.open()
                try:
                    for source_page in order:
                        output.insert_pdf(src, from_page=source_page, to_page=source_page)
                    save = self._save_pdf_atomic_unlocked(
                        fitz,
                        output,
                        out_path,
                        expected_pages=len(order),
                    )
                finally:
                    output.close()
                    src.close()
                if not save.get("success"):
                    return save
            return {
                "success": True,
                "file": save["output_file"],
                "input_file": abs_path,
                "output_file": save["output_file"],
                "source_total_pages": total_pages,
                "pages": save["pages"],
                "page_order": order,
                "backup": save.get("backup"),
            }
        except ValueError as exc:
            return self.page_range_error_result(
                exc,
                0,
                self.MAX_MUTATE_PAGES,
                "pdf_reorder",
            )
        except Exception as exc:
            return {"success": False, "error": str(exc)}

    def add_text(
        self,
        file_path: str,
        page_index: int,
        text: str,
        output_path: str = None,
        *,
        x: float = 72.0,
        y: float = 72.0,
        width: float = 300.0,
        height: float = 72.0,
        font_size: float = 11.0,
        font_name: str = "helv",
        color: str = "000000",
        rotation: int = 0,
    ) -> Dict[str, Any]:
        """Overlay text into a bounded rectangle. This does not reflow PDF content."""
        try:
            import fitz
        except ImportError:
            return {"success": False, "error": "PyMuPDF (fitz) not installed. Run: pip install PyMuPDF"}

        try:
            stamp_text = str(text or "")
            if not stamp_text:
                return {"success": False, "error": "Text is required"}
            if len(stamp_text) > self.MAX_STAMP_TEXT_CHARS:
                return {
                    "success": False,
                    "error": f"Text length exceeds safe limit {self.MAX_STAMP_TEXT_CHARS}",
                    "error_code": "unsafe_pdf_resource_request",
                }
            font_size = float(font_size)
            if not math.isfinite(font_size) or font_size <= 0 or font_size > 144:
                return {"success": False, "error": "font size must be > 0 and <= 144"}
            rotation = int(rotation)
            if rotation not in {0, 90, 180, 270}:
                return {"success": False, "error": "rotation must be one of 0, 90, 180, 270"}
            color_rgb = self._hex_to_rgb(color)
            abs_path = str(Path(file_path).expanduser().resolve())
            out_path = str(Path(output_path).expanduser().resolve()) if output_path else abs_path
            if not os.path.exists(abs_path):
                return {"success": False, "error": f"File not found: {file_path}"}
            with self.host._file_locks(abs_path, out_path):
                pdf = fitz.open(abs_path)
                try:
                    page = pdf[self._page_index(page_index, len(pdf))]
                    rect = self._rect_from_position(fitz, page, x, y, width, height)
                    remaining_height = page.insert_textbox(
                        rect,
                        stamp_text,
                        fontsize=font_size,
                        fontname=str(font_name or "helv"),
                        color=color_rgb,
                        rotate=rotation,
                        overlay=True,
                    )
                    if remaining_height < 0:
                        return {
                            "success": False,
                            "error": "Text does not fit inside the requested rectangle",
                            "error_code": "pdf_text_overflow",
                            "remaining_height": remaining_height,
                            "rect": self._rect_payload(rect),
                        }
                    save = self._save_pdf_atomic_unlocked(
                        fitz,
                        pdf,
                        out_path,
                        expected_pages=len(pdf),
                    )
                finally:
                    pdf.close()
                if not save.get("success"):
                    return save
            return {
                "success": True,
                "file": save["output_file"],
                "input_file": abs_path,
                "output_file": save["output_file"],
                "page": int(page_index),
                "text_length": len(stamp_text),
                "rect": self._rect_payload(rect),
                "font_size": font_size,
                "font_name": str(font_name or "helv"),
                "color": color,
                "rotation": rotation,
                "backup": save.get("backup"),
            }
        except Exception as exc:
            return {"success": False, "error": str(exc)}

    def add_image(
        self,
        file_path: str,
        page_index: int,
        image_path: str,
        output_path: str = None,
        *,
        x: float = 72.0,
        y: float = 72.0,
        width: float = 144.0,
        height: float = 144.0,
        keep_proportion: bool = True,
    ) -> Dict[str, Any]:
        """Overlay an image into a bounded rectangle."""
        try:
            import fitz
            from PIL import Image as PILImage
        except ImportError:
            return {"success": False, "error": "PyMuPDF and Pillow are required for PDF image insertion"}

        try:
            abs_path = str(Path(file_path).expanduser().resolve())
            img_path = str(Path(image_path).expanduser().resolve())
            out_path = str(Path(output_path).expanduser().resolve()) if output_path else abs_path
            if not os.path.exists(abs_path):
                return {"success": False, "error": f"File not found: {file_path}"}
            if not os.path.exists(img_path):
                return {"success": False, "error": f"Image not found: {image_path}"}
            image_size = os.path.getsize(img_path)
            if image_size > self.MAX_STAMP_IMAGE_BYTES:
                return {
                    "success": False,
                    "error": f"Image is {image_size} bytes, exceeding safe limit {self.MAX_STAMP_IMAGE_BYTES}",
                    "error_code": "unsafe_pdf_resource_request",
                }
            with open(img_path, "rb") as handle:
                image_bytes = handle.read()
            img, image_meta, skip_entry = self.load_bounded_image(PILImage, image_bytes)
            if skip_entry:
                return {
                    "success": False,
                    "error": skip_entry.get("error", "Unsafe image"),
                    "error_code": skip_entry.get("error_code", "unsafe_image_decode"),
                    "preflight": skip_entry,
                }
            if img is not None:
                img.close()

            with self.host._file_locks(abs_path, out_path):
                pdf = fitz.open(abs_path)
                try:
                    page = pdf[self._page_index(page_index, len(pdf))]
                    rect = self._rect_from_position(fitz, page, x, y, width, height)
                    page.insert_image(
                        rect,
                        filename=img_path,
                        keep_proportion=bool(keep_proportion),
                        overlay=True,
                    )
                    save = self._save_pdf_atomic_unlocked(
                        fitz,
                        pdf,
                        out_path,
                        expected_pages=len(pdf),
                    )
                finally:
                    pdf.close()
                if not save.get("success"):
                    return save
            return {
                "success": True,
                "file": save["output_file"],
                "input_file": abs_path,
                "output_file": save["output_file"],
                "page": int(page_index),
                "image_file": img_path,
                "image": image_meta,
                "rect": self._rect_payload(rect),
                "keep_proportion": bool(keep_proportion),
                "backup": save.get("backup"),
            }
        except Exception as exc:
            return {"success": False, "error": str(exc)}

    @staticmethod
    def _rect_from_cli_spec(rect_spec: str) -> Tuple[int, float, float, float, float]:
        parts = [part.strip() for part in str(rect_spec or "").split(",")]
        if len(parts) != 5:
            raise ValueError("--rect must be page,left,top,right,bottom")
        try:
            page = int(parts[0])
            left, top, right, bottom = [float(part) for part in parts[1:]]
        except ValueError:
            raise ValueError("--rect values must be numeric: page,left,top,right,bottom") from None
        return page, left, top, right, bottom

    @staticmethod
    def _parse_block_id(block_id: str) -> Tuple[int, int]:
        match = re.fullmatch(r"page_(\d+)_block_(\d+)", str(block_id or "").strip())
        if not match:
            raise ValueError("block_id must look like page_0_block_2 from pdf-read-blocks or pdf-map-page")
        return int(match.group(1)), int(match.group(2))

    @staticmethod
    def _substring_offsets(haystack: str, needle: str) -> List[int]:
        offsets: List[int] = []
        if not needle:
            return offsets
        search_from = 0
        while True:
            found = haystack.find(needle, search_from)
            if found < 0:
                return offsets
            offsets.append(found)
            search_from = found + len(needle)

    def _redaction_rects_for_text(
        self,
        doc: Any,
        query: str,
        *,
        pages: str = None,
        case_sensitive: bool = False,
    ) -> List[Dict[str, Any]]:
        page_indices = self.parse_page_range(
            pages,
            len(doc),
            max_pages=self.MAX_MUTATE_PAGES,
        )
        needle = str(query or "")
        if not needle:
            raise ValueError("--text requires a non-empty query")
        needle_cmp = needle if case_sensitive else needle.lower()
        matches = []
        for page_index in page_indices:
            page = doc[page_index]
            text_dict = page.get_text("rawdict", sort=True)
            for block in text_dict.get("blocks", []):
                if int(block.get("type", 0)) != 0:
                    continue
                for line in block.get("lines", []):
                    line_chars = []
                    for span in line.get("spans", []):
                        for char in span.get("chars", []):
                            char_text = str(char.get("c", "") or "")
                            if not char_text:
                                continue
                            bbox = self.normalize_bbox(char.get("bbox"))
                            if not bbox:
                                continue
                            line_chars.append({"text": char_text, "bbox": bbox})
                    if not line_chars:
                        continue
                    line_text = "".join(char["text"] for char in line_chars)
                    haystack = line_text if case_sensitive else line_text.lower()
                    for offset in self._substring_offsets(haystack, needle_cmp):
                        selected_chars = line_chars[offset : offset + len(needle)]
                        if len(selected_chars) != len(needle):
                            continue
                        left = min(char["bbox"]["left"] for char in selected_chars)
                        top = min(char["bbox"]["top"] for char in selected_chars)
                        right = max(char["bbox"]["right"] for char in selected_chars)
                        bottom = max(char["bbox"]["bottom"] for char in selected_chars)
                        bbox = self.normalize_bbox((left, top, right, bottom))
                        if not bbox:
                            continue
                        matches.append(
                            {
                                "page": page_index,
                                "text": line_text[offset : offset + len(needle)],
                                "line_text_preview": line_text[:160],
                                "rect": (
                                    bbox["left"],
                                    bbox["top"],
                                    bbox["right"],
                                    bbox["bottom"],
                                ),
                                "selector": "text",
                                "geometry": "character_match",
                            }
                        )
                        if len(matches) > self.MAX_REDACTION_MATCHES:
                            raise ValueError(
                                f"Redaction query produced more than {self.MAX_REDACTION_MATCHES} matches; narrow --text or --pages"
                            )
        return matches

    def redact(
        self,
        file_path: str,
        output_path: str = None,
        *,
        text: str = None,
        rects: Optional[List[str]] = None,
        pages: str = None,
        case_sensitive: bool = False,
        fill: str = "000000",
        dry_run: bool = False,
    ) -> Dict[str, Any]:
        """Apply true PDF redactions by exact text match or explicit rectangle."""
        try:
            import fitz
        except ImportError:
            return {"success": False, "error": "PyMuPDF (fitz) not installed. Run: pip install PyMuPDF"}

        try:
            if bool(text) == bool(rects):
                return {
                    "success": False,
                    "error": "Provide exactly one redaction selector: --text or one or more --rect values",
                }
            fill_rgb = self._hex_to_rgb(fill)
            abs_path = str(Path(file_path).expanduser().resolve())
            out_path = str(Path(output_path).expanduser().resolve()) if output_path else abs_path
            if not os.path.exists(abs_path):
                return {"success": False, "error": f"File not found: {file_path}"}
            with self.host._file_locks(abs_path, out_path):
                pdf = fitz.open(abs_path)
                try:
                    matches = []
                    if text:
                        matches = self._redaction_rects_for_text(
                            pdf,
                            text,
                            pages=pages,
                            case_sensitive=case_sensitive,
                        )
                    else:
                        for rect_spec in rects or []:
                            page_num, left, top, right, bottom = self._rect_from_cli_spec(rect_spec)
                            page = pdf[self._page_index(page_num, len(pdf))]
                            rect = self._rect_from_points(fitz, page, left, top, right, bottom)
                            matches.append(
                                {
                                    "page": page_num,
                                    "text": None,
                                    "rect": (rect.x0, rect.y0, rect.x1, rect.y1),
                                }
                            )
                    if dry_run:
                        return {
                            "success": True,
                            "file": abs_path,
                            "dry_run": True,
                            "match_count": len(matches),
                            "matches": [
                                {
                                    "page": match["page"],
                                    "text_preview": (match.get("text") or "")[:80],
                                    "rect": self._rect_payload(fitz.Rect(match["rect"])),
                                }
                                for match in matches[:50]
                            ],
                            "truncated": len(matches) > 50,
                        }
                    pages_touched = set()
                    for match in matches:
                        page = pdf[match["page"]]
                        rect = fitz.Rect(match["rect"])
                        page.add_redact_annot(rect, fill=fill_rgb, cross_out=False)
                        pages_touched.add(match["page"])
                    for page_index_touched in sorted(pages_touched):
                        pdf[page_index_touched].apply_redactions(
                            images=fitz.PDF_REDACT_IMAGE_PIXELS,
                            graphics=fitz.PDF_REDACT_LINE_ART_REMOVE_IF_COVERED,
                            text=fitz.PDF_REDACT_TEXT_REMOVE,
                        )
                    save = self._save_pdf_atomic_unlocked(
                        fitz,
                        pdf,
                        out_path,
                        expected_pages=len(pdf),
                    )
                finally:
                    pdf.close()
                if not save.get("success"):
                    return save
            verification = None
            if text:
                verification = self.search_blocks(
                    save["output_file"],
                    text,
                    pages=pages,
                    case_sensitive=case_sensitive,
                    include_spans=True,
                )
            return {
                "success": True,
                "file": save["output_file"],
                "input_file": abs_path,
                "output_file": save["output_file"],
                "redactions_applied": len(matches),
                "pages_touched": sorted(pages_touched),
                "fill": fill,
                "backup": save.get("backup"),
                "verification": verification,
                "warnings": (
                    ["Redaction text still appears in the output PDF; inspect manually."]
                    if verification and verification.get("match_count", 0) > 0
                    else []
                ),
            }
        except ValueError as exc:
            return {
                "success": False,
                "error": str(exc),
                "error_code": "usage_error",
            }
        except Exception as exc:
            return {"success": False, "error": str(exc)}

    def redact_block(
        self,
        file_path: str,
        block_id: str,
        output_path: str = None,
        *,
        fill: str = "000000",
        dry_run: bool = False,
    ) -> Dict[str, Any]:
        """Redact one native PDF block by block_id from pdf-read-blocks/pdf-map-page."""
        try:
            page_index, block_index = self._parse_block_id(block_id)
            self._hex_to_rgb(fill)
            payload = self.read_blocks(
                file_path,
                pages=str(page_index),
                include_spans=False,
                include_images=True,
            )
            if not payload.get("success"):
                return payload
            page_payload = payload.get("pages", [{}])[0]
            block = None
            for candidate in page_payload.get("blocks", []):
                if candidate.get("block_id") == block_id:
                    block = candidate
                    break
            if not block:
                return {
                    "success": False,
                    "error": f"Block not found: {block_id}",
                    "error_code": "usage_error",
                    "page": page_index,
                    "block_index": block_index,
                }
            bbox = block.get("bbox") or {}
            if not bbox:
                return {
                    "success": False,
                    "error": f"Block has no bounding box: {block_id}",
                    "error_code": "usage_error",
                    "block": block,
                }
            rect_spec = (
                f"{page_index},{bbox['left']},{bbox['top']},"
                f"{bbox['right']},{bbox['bottom']}"
            )
            if dry_run:
                return {
                    "success": True,
                    "file": str(Path(file_path).expanduser().resolve()),
                    "dry_run": True,
                    "block_id": block_id,
                    "page": page_index,
                    "block": block,
                    "rect": bbox,
                }
            result = self.redact(
                file_path,
                output_path=output_path,
                rects=[rect_spec],
                fill=fill,
                dry_run=False,
            )
            if result.get("success"):
                result["block_id"] = block_id
                result["block"] = block
                result["selector"] = "block"
            return result
        except ValueError as exc:
            return {
                "success": False,
                "error": str(exc),
                "error_code": "usage_error",
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
        remove_annotations: bool = False,
        remove_embedded_files: bool = False,
        flatten_forms: bool = False,
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
                remove_annotations,
                remove_embedded_files,
                flatten_forms,
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
            abs_path = str(Path(file_path).expanduser().resolve())
            out_path = str(Path(output_path).expanduser().resolve()) if output_path else abs_path

            with self.host._file_locks(abs_path, out_path):
                pdf = fitz.open(abs_path)
                try:
                    before = dict(pdf.metadata or {})
                    before_hidden_data = self.hidden_data_summary(pdf)
                    metadata = {} if clear_metadata else dict(before)
                    assignments = {
                        "author": author,
                        "title": title,
                        "subject": subject,
                        "keywords": keywords,
                        "creator": creator,
                        "producer": producer,
                    }
                    metadata_updated = any(value is not None for value in assignments.values())
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
                    xml_metadata_action = "not_requested"
                    if remove_xml_metadata:
                        if hasattr(pdf, "del_xml_metadata"):
                            pdf.del_xml_metadata()
                            xml_metadata_action = "requested"
                        else:
                            xml_metadata_action = "unsupported_by_pymupdf"

                    annotations_action = "not_requested"
                    annotations_removed = 0
                    annotation_types_removed: Dict[str, int] = {}
                    if remove_annotations:
                        annotations_action = "removed"
                        for page in pdf:
                            annot = page.first_annot
                            while annot is not None:
                                next_annot = annot.next
                                try:
                                    annot_type = annot.type[1]
                                except Exception:
                                    annot_type = "unknown"
                                try:
                                    page.delete_annot(annot)
                                    annotations_removed += 1
                                    annotation_types_removed[annot_type] = (
                                        annotation_types_removed.get(annot_type, 0) + 1
                                    )
                                except Exception:
                                    annotations_action = "requested_incomplete"
                                annot = next_annot

                    embedded_files_action = "not_requested"
                    embedded_files_removed = 0
                    if remove_embedded_files:
                        if hasattr(pdf, "embfile_names") and hasattr(pdf, "embfile_del"):
                            embedded_files_action = "removed"
                            try:
                                for name in list(pdf.embfile_names()):
                                    pdf.embfile_del(name)
                                    embedded_files_removed += 1
                            except Exception:
                                embedded_files_action = "requested_incomplete"
                        else:
                            embedded_files_action = "unsupported_by_pymupdf"

                    forms_action = "not_requested"
                    if flatten_forms:
                        if hasattr(pdf, "bake"):
                            forms_action = "flattened"
                            try:
                                pdf.bake(annots=False, widgets=True)
                            except Exception:
                                forms_action = "requested_incomplete"
                        else:
                            forms_action = "unsupported_by_pymupdf"

                    save_result = self._save_pdf_atomic_unlocked(
                        fitz,
                        pdf,
                        out_path,
                        expected_pages=len(pdf),
                        save_options=self._save_options(
                            garbage=4,
                            clean=True,
                            deflate=False,
                        ),
                    )
                    if not save_result.get("success"):
                        return save_result
                    backup = save_result.get("backup")
                finally:
                    pdf.close()

                out_doc = fitz.open(out_path)
                after = dict(out_doc.metadata or {})
                after_hidden_data = self.hidden_data_summary(out_doc)
                has_xml_metadata = after_hidden_data["has_xml_metadata"]
                page_count = len(out_doc)
                page_size = None
                if page_count:
                    page = out_doc[0]
                    page_size = {
                        "width": float(page.rect.width),
                        "height": float(page.rect.height),
                    }
                out_doc.close()

            residual_hidden_data = {
                "annotations_count": after_hidden_data.get("annotations_count", 0),
                "annotation_types": after_hidden_data.get("annotation_types", {}),
                "embedded_files_count": after_hidden_data.get("embedded_files_count", 0),
                "has_embedded_files": after_hidden_data.get("has_embedded_files", False),
                "has_forms": after_hidden_data.get("has_forms"),
                "has_xml_metadata": after_hidden_data.get("has_xml_metadata", False),
            }
            if remove_xml_metadata and xml_metadata_action == "requested":
                xml_metadata_action = (
                    "requested_incomplete"
                    if residual_hidden_data["has_xml_metadata"]
                    else "removed"
                )
            if remove_annotations and residual_hidden_data["annotations_count"]:
                annotations_action = "requested_incomplete"
            if remove_embedded_files and residual_hidden_data["has_embedded_files"]:
                embedded_files_action = "requested_incomplete"
            if flatten_forms and residual_hidden_data["has_forms"]:
                forms_action = "requested_incomplete"
            if clear_metadata:
                document_metadata_action = "cleared"
            elif metadata_updated:
                document_metadata_action = "updated"
            else:
                document_metadata_action = "unchanged"

            warnings = []
            if residual_hidden_data["has_xml_metadata"]:
                if remove_xml_metadata:
                    warnings.append(
                        "PDF sanitize was asked to remove XML/XMP metadata, but XML metadata remains; inspect manually."
                    )
                else:
                    warnings.append(
                        "PDF still contains XML/XMP metadata; use --remove-xml-metadata if removal is intended."
                    )
            if residual_hidden_data["annotations_count"]:
                if remove_annotations:
                    warnings.append(
                        "PDF sanitize was asked to remove annotations, but annotations remain; inspect manually."
                    )
                else:
                    warnings.append(
                        "PDF still contains annotations; use --remove-annotations if removal is intended."
                    )
            if residual_hidden_data["has_embedded_files"]:
                if remove_embedded_files:
                    warnings.append(
                        "PDF sanitize was asked to remove embedded files/attachments, but embedded files remain; inspect manually."
                    )
                else:
                    warnings.append(
                        "PDF still contains embedded files/attachments; use --remove-embedded-files if removal is intended."
                    )
            if residual_hidden_data["has_forms"]:
                if flatten_forms:
                    warnings.append(
                        "PDF sanitize was asked to flatten forms, but form fields remain; inspect manually."
                    )
                else:
                    warnings.append(
                        "PDF still contains form fields; use --flatten-forms if flattening is intended."
                    )
            if remove_xml_metadata and xml_metadata_action == "unsupported_by_pymupdf":
                warnings.append("PyMuPDF did not expose XML metadata deletion for this PDF.")
            if remove_embedded_files and embedded_files_action == "unsupported_by_pymupdf":
                warnings.append("PyMuPDF did not expose embedded-file deletion for this PDF.")
            if flatten_forms and forms_action == "unsupported_by_pymupdf":
                warnings.append("PyMuPDF did not expose form flattening for this PDF.")

            return {
                "success": True,
                "file": out_path,
                "input_file": abs_path,
                "output_file": out_path,
                "clear_metadata": clear_metadata,
                "remove_xml_metadata": remove_xml_metadata,
                "remove_annotations": remove_annotations,
                "remove_embedded_files": remove_embedded_files,
                "flatten_forms": flatten_forms,
                "before": before,
                "after": after,
                "has_xml_metadata": has_xml_metadata,
                "sanitization_scope": {
                    "document_metadata": document_metadata_action,
                    "xml_metadata": xml_metadata_action,
                    "annotations": annotations_action
                    if remove_annotations
                    else "reported_not_removed",
                    "embedded_files": embedded_files_action
                    if remove_embedded_files
                    else "reported_not_removed",
                    "forms": forms_action if flatten_forms else "reported_not_removed",
                },
                "actions": {
                    "annotations_removed": annotations_removed,
                    "annotation_types_removed": annotation_types_removed,
                    "embedded_files_removed": embedded_files_removed,
                    "forms_action": forms_action,
                },
                "before_hidden_data": before_hidden_data,
                "after_hidden_data": after_hidden_data,
                "residual_hidden_data": residual_hidden_data,
                "warnings": warnings,
                "pages": page_count,
                "page_size": page_size,
                "backup": backup or None,
            }
        except Exception as exc:
            return {"success": False, "error": str(exc)}
