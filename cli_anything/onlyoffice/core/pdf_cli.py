#!/usr/bin/env python3
"""PDF CLI command handlers."""

from __future__ import annotations

from typing import Any, Callable, List

from cli_anything.onlyoffice.core.command_registry import command_usage


def handle_pdf_command(
    command: str,
    raw_args: List[str],
    doc_server: Any,
    json_output: bool,
    print_result: Callable[[dict, bool], None],
) -> bool:
    """Handle PDF commands and return True when the command was recognised."""
    if command == "pdf-extract-images":
        if len(raw_args) < 2:
            print_result(
                {
                    "success": False,
                    "error": command_usage("pdf-extract-images"),
                },
                json_output,
            )
            return True
        fmt = "png"
        pages = None
        index = 2
        while index < len(raw_args):
            if raw_args[index] == "--format" and index + 1 < len(raw_args):
                fmt = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--pages" and index + 1 < len(raw_args):
                pages = raw_args[index + 1]
                index += 2
            else:
                index += 1
        print_result(
            doc_server.pdf_extract_images(raw_args[0], raw_args[1], fmt=fmt, pages=pages),
            json_output,
        )
        return True

    if command == "pdf-page-to-image":
        if len(raw_args) < 2:
            print_result(
                {
                    "success": False,
                    "error": command_usage("pdf-page-to-image"),
                },
                json_output,
            )
            return True
        pages = None
        dpi = 150
        fmt = "png"
        index = 2
        while index < len(raw_args):
            if raw_args[index] == "--pages" and index + 1 < len(raw_args):
                pages = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--dpi" and index + 1 < len(raw_args):
                dpi = int(raw_args[index + 1])
                index += 2
            elif raw_args[index] == "--format" and index + 1 < len(raw_args):
                fmt = raw_args[index + 1]
                index += 2
            else:
                index += 1
        print_result(
            doc_server.pdf_page_to_image(raw_args[0], raw_args[1], pages=pages, dpi=dpi, fmt=fmt),
            json_output,
        )
        return True

    if command == "pdf-read-blocks":
        if not raw_args:
            print_result(
                {
                    "success": False,
                    "error": command_usage("pdf-read-blocks"),
                },
                json_output,
            )
            return True
        pages = None
        include_spans = True
        include_images = True
        include_empty = False
        index = 1
        while index < len(raw_args):
            if raw_args[index] == "--pages" and index + 1 < len(raw_args):
                pages = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--no-spans":
                include_spans = False
                index += 1
            elif raw_args[index] == "--no-images":
                include_images = False
                index += 1
            elif raw_args[index] == "--include-empty":
                include_empty = True
                index += 1
            else:
                index += 1
        print_result(
            doc_server.pdf_read_blocks(
                raw_args[0],
                pages=pages,
                include_spans=include_spans,
                include_images=include_images,
                include_empty=include_empty,
            ),
            json_output,
        )
        return True

    if command == "pdf-search-blocks":
        if len(raw_args) < 2:
            print_result(
                {
                    "success": False,
                    "error": command_usage("pdf-search-blocks"),
                },
                json_output,
            )
            return True
        pages = None
        case_sensitive = False
        include_spans = True
        index = 2
        while index < len(raw_args):
            if raw_args[index] == "--pages" and index + 1 < len(raw_args):
                pages = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--case-sensitive":
                case_sensitive = True
                index += 1
            elif raw_args[index] == "--no-spans":
                include_spans = False
                index += 1
            else:
                index += 1
        print_result(
            doc_server.pdf_search_blocks(
                raw_args[0],
                raw_args[1],
                pages=pages,
                case_sensitive=case_sensitive,
                include_spans=include_spans,
            ),
            json_output,
        )
        return True

    if command == "pdf-inspect-hidden-data":
        if not raw_args:
            print_result(
                {"success": False, "error": command_usage("pdf-inspect-hidden-data")},
                json_output,
            )
            return True
        print_result(doc_server.inspect_pdf_hidden_data(raw_args[0]), json_output)
        return True

    if command == "pdf-sanitize":
        if not raw_args:
            print_result(
                {
                    "success": False,
                    "error": command_usage("pdf-sanitize"),
                },
                json_output,
            )
            return True
        file_path = raw_args[0]
        output_path = None
        clear_metadata = False
        remove_xml_metadata = False
        author = None
        title = None
        subject = None
        keywords = None
        creator = None
        producer = None
        index = 1
        while index < len(raw_args):
            if raw_args[index] == "--clear-metadata":
                clear_metadata = True
                index += 1
            elif raw_args[index] == "--remove-xml-metadata":
                remove_xml_metadata = True
                index += 1
            elif raw_args[index] == "--author" and index + 1 < len(raw_args):
                author = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--title" and index + 1 < len(raw_args):
                title = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--subject" and index + 1 < len(raw_args):
                subject = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--keywords" and index + 1 < len(raw_args):
                keywords = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--creator" and index + 1 < len(raw_args):
                creator = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--producer" and index + 1 < len(raw_args):
                producer = raw_args[index + 1]
                index += 2
            elif output_path is None and not raw_args[index].startswith("--"):
                output_path = raw_args[index]
                index += 1
            else:
                index += 1
        print_result(
            doc_server.pdf_sanitize(
                file_path,
                output_path=output_path,
                clear_metadata=clear_metadata,
                remove_xml_metadata=remove_xml_metadata,
                author=author,
                title=title,
                subject=subject,
                keywords=keywords,
                creator=creator,
                producer=producer,
            ),
            json_output,
        )
        return True

    return False
