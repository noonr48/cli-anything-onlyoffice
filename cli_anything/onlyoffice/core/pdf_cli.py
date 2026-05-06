#!/usr/bin/env python3
"""PDF CLI command handlers."""

from __future__ import annotations

from typing import Any, Callable, List, Optional, Tuple

from cli_anything.onlyoffice.core.command_registry import command_usage
from cli_anything.onlyoffice.core.parse_utils import parse_float, parse_int


def _option_value(command: str, raw_args: List[str], index: int, option: str) -> str:
    if index + 1 >= len(raw_args) or raw_args[index + 1].startswith("--"):
        raise ValueError(
            f"{command}: {option} requires a value. {command_usage(command)}"
        )
    return raw_args[index + 1]


def _raise_unexpected_option(command: str, token: str) -> None:
    if token.startswith("--"):
        raise ValueError(
            f"{command}: unknown option {token!r}. {command_usage(command)}"
        )
    raise ValueError(
        f"{command}: unexpected argument {token!r}. {command_usage(command)}"
    )


def _parse_optional_output(
    command: str,
    raw_args: List[str],
    index: int,
    output_path: Optional[str],
) -> Tuple[Optional[str], int]:
    token = raw_args[index]
    if token.startswith("--") or output_path is not None:
        _raise_unexpected_option(command, token)
    return token, index + 1


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
            if raw_args[index] == "--format":
                fmt = _option_value(command, raw_args, index, "--format")
                index += 2
            elif raw_args[index] == "--pages":
                pages = _option_value(command, raw_args, index, "--pages")
                index += 2
            else:
                _raise_unexpected_option(command, raw_args[index])
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
            if raw_args[index] == "--pages":
                pages = _option_value(command, raw_args, index, "--pages")
                index += 2
            elif raw_args[index] == "--dpi":
                dpi = parse_int(_option_value(command, raw_args, index, "--dpi"), "--dpi")
                index += 2
            elif raw_args[index] == "--format":
                fmt = _option_value(command, raw_args, index, "--format")
                index += 2
            else:
                _raise_unexpected_option(command, raw_args[index])
        print_result(
            doc_server.pdf_page_to_image(raw_args[0], raw_args[1], pages=pages, dpi=dpi, fmt=fmt),
            json_output,
        )
        return True

    if command == "pdf-map-page":
        if len(raw_args) < 3:
            print_result(
                {"success": False, "error": command_usage("pdf-map-page")},
                json_output,
            )
            return True
        dpi = 150
        fmt = "png"
        labels = True
        include_images = True
        index = 3
        while index < len(raw_args):
            if raw_args[index] == "--dpi":
                dpi = parse_int(_option_value(command, raw_args, index, "--dpi"), "--dpi")
                index += 2
            elif raw_args[index] == "--format":
                fmt = _option_value(command, raw_args, index, "--format")
                index += 2
            elif raw_args[index] == "--no-labels":
                labels = False
                index += 1
            elif raw_args[index] == "--no-images":
                include_images = False
                index += 1
            else:
                _raise_unexpected_option(command, raw_args[index])
        print_result(
            doc_server.pdf_map_page(
                raw_args[0],
                parse_int(raw_args[1], "page"),
                raw_args[2],
                dpi=dpi,
                fmt=fmt,
                labels=labels,
                include_images=include_images,
            ),
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
            if raw_args[index] == "--pages":
                pages = _option_value(command, raw_args, index, "--pages")
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
                _raise_unexpected_option(command, raw_args[index])
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
            if raw_args[index] == "--pages":
                pages = _option_value(command, raw_args, index, "--pages")
                index += 2
            elif raw_args[index] == "--case-sensitive":
                case_sensitive = True
                index += 1
            elif raw_args[index] == "--no-spans":
                include_spans = False
                index += 1
            else:
                _raise_unexpected_option(command, raw_args[index])
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
        remove_annotations = False
        remove_embedded_files = False
        flatten_forms = False
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
            elif raw_args[index] == "--remove-annotations":
                remove_annotations = True
                index += 1
            elif raw_args[index] == "--remove-embedded-files":
                remove_embedded_files = True
                index += 1
            elif raw_args[index] == "--flatten-forms":
                flatten_forms = True
                index += 1
            elif raw_args[index] == "--author":
                author = _option_value(command, raw_args, index, "--author")
                index += 2
            elif raw_args[index] == "--title":
                title = _option_value(command, raw_args, index, "--title")
                index += 2
            elif raw_args[index] == "--subject":
                subject = _option_value(command, raw_args, index, "--subject")
                index += 2
            elif raw_args[index] == "--keywords":
                keywords = _option_value(command, raw_args, index, "--keywords")
                index += 2
            elif raw_args[index] == "--creator":
                creator = _option_value(command, raw_args, index, "--creator")
                index += 2
            elif raw_args[index] == "--producer":
                producer = _option_value(command, raw_args, index, "--producer")
                index += 2
            else:
                output_path, index = _parse_optional_output(
                    command, raw_args, index, output_path
                )
        print_result(
            doc_server.pdf_sanitize(
                file_path,
                output_path=output_path,
                clear_metadata=clear_metadata,
                remove_xml_metadata=remove_xml_metadata,
                remove_annotations=remove_annotations,
                remove_embedded_files=remove_embedded_files,
                flatten_forms=flatten_forms,
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

    if command == "pdf-compact":
        if not raw_args:
            print_result(
                {"success": False, "error": command_usage("pdf-compact")},
                json_output,
            )
            return True
        output_path = None
        garbage = 4
        deflate = True
        clean = True
        linearize = False
        index = 1
        while index < len(raw_args):
            if raw_args[index] == "--garbage":
                garbage = parse_int(
                    _option_value(command, raw_args, index, "--garbage"),
                    "--garbage",
                )
                index += 2
            elif raw_args[index] == "--no-deflate":
                deflate = False
                index += 1
            elif raw_args[index] == "--no-clean":
                clean = False
                index += 1
            elif raw_args[index] == "--linearize":
                linearize = True
                index += 1
            else:
                output_path, index = _parse_optional_output(
                    command, raw_args, index, output_path
                )
        print_result(
            doc_server.pdf_compact(
                raw_args[0],
                output_path=output_path,
                garbage=garbage,
                deflate=deflate,
                clean=clean,
                linearize=linearize,
            ),
            json_output,
        )
        return True

    if command == "pdf-merge":
        if len(raw_args) < 4:
            print_result(
                {"success": False, "error": command_usage("pdf-merge")},
                json_output,
            )
            return True
        output_path = None
        input_files = []
        index = 0
        while index < len(raw_args):
            if raw_args[index] == "--output":
                output_path = _option_value(command, raw_args, index, "--output")
                index += 2
            elif raw_args[index].startswith("--"):
                _raise_unexpected_option(command, raw_args[index])
            else:
                input_files.append(raw_args[index])
                index += 1
        if not output_path:
            raise ValueError(f"pdf-merge: --output is required. {command_usage(command)}")
        print_result(
            doc_server.pdf_merge(input_files, output_path),
            json_output,
        )
        return True

    if command == "pdf-split":
        if len(raw_args) < 2:
            print_result(
                {"success": False, "error": command_usage("pdf-split")},
                json_output,
            )
            return True
        pages = None
        prefix = "page"
        index = 2
        while index < len(raw_args):
            if raw_args[index] == "--pages":
                pages = _option_value(command, raw_args, index, "--pages")
                index += 2
            elif raw_args[index] == "--prefix":
                prefix = _option_value(command, raw_args, index, "--prefix")
                index += 2
            else:
                _raise_unexpected_option(command, raw_args[index])
        print_result(
            doc_server.pdf_split(raw_args[0], raw_args[1], pages=pages, prefix=prefix),
            json_output,
        )
        return True

    if command == "pdf-reorder":
        if len(raw_args) < 2:
            print_result(
                {"success": False, "error": command_usage("pdf-reorder")},
                json_output,
            )
            return True
        output_path = None
        index = 2
        while index < len(raw_args):
            output_path, index = _parse_optional_output(
                command, raw_args, index, output_path
            )
        print_result(
            doc_server.pdf_reorder(raw_args[0], raw_args[1], output_path=output_path),
            json_output,
        )
        return True

    if command == "pdf-add-text":
        if len(raw_args) < 3:
            print_result(
                {"success": False, "error": command_usage("pdf-add-text")},
                json_output,
            )
            return True
        output_path = None
        x = 72.0
        y = 72.0
        width = 300.0
        height = 72.0
        font_size = 11.0
        font_name = "helv"
        color = "000000"
        rotation = 0
        index = 3
        while index < len(raw_args):
            if raw_args[index] == "--output":
                output_path = _option_value(command, raw_args, index, "--output")
                index += 2
            elif raw_args[index] == "--x":
                x = parse_float(_option_value(command, raw_args, index, "--x"), "--x")
                index += 2
            elif raw_args[index] == "--y":
                y = parse_float(_option_value(command, raw_args, index, "--y"), "--y")
                index += 2
            elif raw_args[index] == "--width":
                width = parse_float(
                    _option_value(command, raw_args, index, "--width"),
                    "--width",
                )
                index += 2
            elif raw_args[index] == "--height":
                height = parse_float(
                    _option_value(command, raw_args, index, "--height"),
                    "--height",
                )
                index += 2
            elif raw_args[index] == "--font-size":
                font_size = parse_float(
                    _option_value(command, raw_args, index, "--font-size"),
                    "--font-size",
                )
                index += 2
            elif raw_args[index] == "--font":
                font_name = _option_value(command, raw_args, index, "--font")
                index += 2
            elif raw_args[index] == "--color":
                color = _option_value(command, raw_args, index, "--color")
                index += 2
            elif raw_args[index] == "--rotation":
                rotation = parse_int(
                    _option_value(command, raw_args, index, "--rotation"),
                    "--rotation",
                )
                index += 2
            else:
                _raise_unexpected_option(command, raw_args[index])
        print_result(
            doc_server.pdf_add_text(
                raw_args[0],
                parse_int(raw_args[1], "page"),
                raw_args[2],
                output_path=output_path,
                x=x,
                y=y,
                width=width,
                height=height,
                font_size=font_size,
                font_name=font_name,
                color=color,
                rotation=rotation,
            ),
            json_output,
        )
        return True

    if command == "pdf-add-image":
        if len(raw_args) < 3:
            print_result(
                {"success": False, "error": command_usage("pdf-add-image")},
                json_output,
            )
            return True
        output_path = None
        x = 72.0
        y = 72.0
        width = 144.0
        height = 144.0
        keep_proportion = True
        index = 3
        while index < len(raw_args):
            if raw_args[index] == "--output":
                output_path = _option_value(command, raw_args, index, "--output")
                index += 2
            elif raw_args[index] == "--x":
                x = parse_float(_option_value(command, raw_args, index, "--x"), "--x")
                index += 2
            elif raw_args[index] == "--y":
                y = parse_float(_option_value(command, raw_args, index, "--y"), "--y")
                index += 2
            elif raw_args[index] == "--width":
                width = parse_float(
                    _option_value(command, raw_args, index, "--width"),
                    "--width",
                )
                index += 2
            elif raw_args[index] == "--height":
                height = parse_float(
                    _option_value(command, raw_args, index, "--height"),
                    "--height",
                )
                index += 2
            elif raw_args[index] == "--no-keep-proportion":
                keep_proportion = False
                index += 1
            else:
                _raise_unexpected_option(command, raw_args[index])
        print_result(
            doc_server.pdf_add_image(
                raw_args[0],
                parse_int(raw_args[1], "page"),
                raw_args[2],
                output_path=output_path,
                x=x,
                y=y,
                width=width,
                height=height,
                keep_proportion=keep_proportion,
            ),
            json_output,
        )
        return True

    if command == "pdf-redact":
        if not raw_args:
            print_result(
                {"success": False, "error": command_usage("pdf-redact")},
                json_output,
            )
            return True
        output_path = None
        text = None
        rects = []
        pages = None
        case_sensitive = False
        fill = "000000"
        dry_run = False
        index = 1
        while index < len(raw_args):
            if raw_args[index] == "--text":
                text = _option_value(command, raw_args, index, "--text")
                index += 2
            elif raw_args[index] == "--rect":
                rects.append(_option_value(command, raw_args, index, "--rect"))
                index += 2
            elif raw_args[index] == "--pages":
                pages = _option_value(command, raw_args, index, "--pages")
                index += 2
            elif raw_args[index] == "--case-sensitive":
                case_sensitive = True
                index += 1
            elif raw_args[index] == "--fill":
                fill = _option_value(command, raw_args, index, "--fill")
                index += 2
            elif raw_args[index] == "--dry-run":
                dry_run = True
                index += 1
            else:
                output_path, index = _parse_optional_output(
                    command, raw_args, index, output_path
                )
        print_result(
            doc_server.pdf_redact(
                raw_args[0],
                output_path=output_path,
                text=text,
                rects=rects or None,
                pages=pages,
                case_sensitive=case_sensitive,
                fill=fill,
                dry_run=dry_run,
            ),
            json_output,
        )
        return True

    if command == "pdf-redact-block":
        if len(raw_args) < 2:
            print_result(
                {"success": False, "error": command_usage("pdf-redact-block")},
                json_output,
            )
            return True
        output_path = None
        fill = "000000"
        dry_run = False
        index = 2
        while index < len(raw_args):
            if raw_args[index] == "--fill":
                fill = _option_value(command, raw_args, index, "--fill")
                index += 2
            elif raw_args[index] == "--dry-run":
                dry_run = True
                index += 1
            else:
                output_path, index = _parse_optional_output(
                    command,
                    raw_args,
                    index,
                    output_path,
                )
        print_result(
            doc_server.pdf_redact_block(
                raw_args[0],
                raw_args[1],
                output_path=output_path,
                fill=fill,
                dry_run=dry_run,
            ),
            json_output,
        )
        return True

    return False
