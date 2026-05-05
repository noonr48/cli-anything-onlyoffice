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
from typing import Any, Dict, List, Optional


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

            with self.host._file_locks(abs_path, out_path):
                backup_target = abs_path if overwrite else out_path
                backup = self.host._snapshot_backup(backup_target) if Path(backup_target).exists() else None
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
            warnings = []
            if residual_hidden_data["annotations_count"]:
                warnings.append(
                    "PDF sanitize does not remove annotations; inspect/remove them separately if needed."
                )
            if residual_hidden_data["has_embedded_files"]:
                warnings.append(
                    "PDF sanitize does not remove embedded files/attachments; inspect/remove them separately if needed."
                )
            if residual_hidden_data["has_forms"]:
                warnings.append(
                    "PDF sanitize does not flatten or remove form fields; inspect/remove them separately if needed."
                )
            if remove_xml_metadata and xml_metadata_action == "unsupported_by_pymupdf":
                warnings.append("PyMuPDF did not expose XML metadata deletion for this PDF.")

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
                "sanitization_scope": {
                    "document_metadata": "cleared" if clear_metadata else "updated",
                    "xml_metadata": xml_metadata_action,
                    "annotations": "reported_not_removed",
                    "embedded_files": "reported_not_removed",
                    "forms": "reported_not_removed",
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
