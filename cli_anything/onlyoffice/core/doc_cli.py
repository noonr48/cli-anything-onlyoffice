#!/usr/bin/env python3
"""DOCX CLI command handlers."""

from __future__ import annotations

from typing import Any, Callable, List

from cli_anything.onlyoffice.core.command_registry import command_usage
from cli_anything.onlyoffice.core.parse_utils import parse_float, parse_int, print_usage_error


def _docx_unavailable(
    json_output: bool,
    print_result: Callable[[dict, bool], None],
) -> bool:
    print_result(
        {"success": False, "error": "python-docx not installed"},
        json_output,
    )
    return True


def _option_key(option: str) -> str:
    return option.lstrip("-").replace("-", "_")


def _literal_start(raw_args: List[str]) -> int | None:
    return getattr(raw_args, "literal_start", None)


def _is_literal_arg(raw_args: List[str], index: int) -> bool:
    start = _literal_start(raw_args)
    return start is not None and index >= start


def _is_option_like(value: str) -> bool:
    return value.startswith("--")


def _doc_usage_error(
    command: str,
    json_output: bool,
    print_result: Callable[[dict, bool], None],
    message: str,
) -> bool:
    print_usage_error(
        print_result,
        json_output,
        message,
        usage=command_usage(command),
    )
    return True


def _option_value(
    command: str,
    raw_args: List[str],
    index: int,
    json_output: bool,
    print_result: Callable[[dict, bool], None],
) -> str | None:
    option = raw_args[index]
    value_index = index + 1
    if value_index >= len(raw_args):
        _doc_usage_error(
            command,
            json_output,
            print_result,
            f"Missing value for {option}",
        )
        return None
    value = raw_args[value_index]
    if _is_option_like(value) and not _is_literal_arg(raw_args, value_index):
        _doc_usage_error(
            command,
            json_output,
            print_result,
            f"Missing value for {option}",
        )
        return None
    return value


def _reject_extra_args(
    command: str,
    raw_args: List[str],
    expected_count: int,
    json_output: bool,
    print_result: Callable[[dict, bool], None],
) -> bool:
    if len(raw_args) <= expected_count:
        return False
    token = raw_args[expected_count]
    if _is_literal_arg(raw_args, expected_count):
        message = f"Unexpected argument after -- for {command}: {token}"
    elif _is_option_like(token):
        message = f"Unknown option for {command}: {token}"
    else:
        message = f"Unexpected argument for {command}: {token}"
    return _doc_usage_error(command, json_output, print_result, message)


def _parse_options(
    command: str,
    raw_args: List[str],
    start_index: int,
    json_output: bool,
    print_result: Callable[[dict, bool], None],
    *,
    value_options: tuple[str, ...] = (),
    flag_options: tuple[str, ...] = (),
    converters: dict[str, Callable[[str, str], Any]] | None = None,
    max_positionals: int = 0,
) -> tuple[dict[str, Any], dict[str, bool], list[str]] | None:
    values: dict[str, Any] = {}
    flags = {_option_key(option): False for option in flag_options}
    positionals: list[str] = []
    converters = converters or {}
    value_options_set = set(value_options)
    flag_options_set = set(flag_options)
    index = start_index
    while index < len(raw_args):
        token = raw_args[index]
        if _is_literal_arg(raw_args, index):
            if len(positionals) >= max_positionals:
                _doc_usage_error(
                    command,
                    json_output,
                    print_result,
                    f"Unexpected argument after -- for {command}: {token}",
                )
                return None
            positionals.append(token)
            index += 1
            continue
        if token in flag_options_set:
            flags[_option_key(token)] = True
            index += 1
            continue
        if token in value_options_set:
            value = _option_value(command, raw_args, index, json_output, print_result)
            if value is None:
                return None
            converter = converters.get(token)
            values[_option_key(token)] = converter(value, token) if converter else value
            index += 2
            continue
        if _is_option_like(token):
            _doc_usage_error(
                command,
                json_output,
                print_result,
                f"Unknown option for {command}: {token}",
            )
            return None
        if len(positionals) < max_positionals:
            positionals.append(token)
            index += 1
            continue
        _doc_usage_error(
            command,
            json_output,
            print_result,
            f"Unexpected argument for {command}: {token}",
        )
        return None
    return values, flags, positionals


def handle_doc_command(
    command: str,
    raw_args: List[str],
    doc_server: Any,
    json_output: bool,
    print_result: Callable[[dict, bool], None],
    docx_available: bool = True,
) -> bool:
    """Handle DOCX commands and return True when the command was recognised."""
    if command == "doc-create":
        if len(raw_args) < 3:
            print_result(
                {"success": False, "error": command_usage("doc-create")},
                json_output,
            )
            return True
        if not docx_available:
            return _docx_unavailable(json_output, print_result)
        print_result(
            doc_server.create_document(raw_args[0], raw_args[1], " ".join(raw_args[2:])),
            json_output,
        )
        return True

    if command == "doc-read":
        if not raw_args:
            print_result(
                {"success": False, "error": command_usage("doc-read")},
                json_output,
            )
            return True
        if _reject_extra_args(command, raw_args, 1, json_output, print_result):
            return True
        if not docx_available:
            return _docx_unavailable(json_output, print_result)
        print_result(doc_server.read_document(raw_args[0]), json_output)
        return True

    if command == "doc-append":
        if len(raw_args) < 2:
            print_result(
                {"success": False, "error": command_usage("doc-append")},
                json_output,
            )
            return True
        if not docx_available:
            return _docx_unavailable(json_output, print_result)
        print_result(doc_server.append_to_document(raw_args[0], " ".join(raw_args[1:])), json_output)
        return True

    if command == "doc-replace":
        if len(raw_args) < 3:
            print_result(
                {"success": False, "error": command_usage("doc-replace")},
                json_output,
            )
            return True
        if not docx_available:
            return _docx_unavailable(json_output, print_result)
        print_result(
            doc_server.search_replace_document(raw_args[0], raw_args[1], " ".join(raw_args[2:])),
            json_output,
        )
        return True

    if command == "doc-format":
        if len(raw_args) < 2:
            print_result(
                {
                    "success": False,
                    "error": command_usage("doc-format"),
                },
                json_output,
            )
            return True
        if not docx_available:
            return _docx_unavailable(json_output, print_result)
        parsed = _parse_options(
            command,
            raw_args,
            2,
            json_output,
            print_result,
            value_options=("--font-name", "--font-size", "--color", "--align"),
            flag_options=("--bold", "--italic", "--underline"),
            converters={"--font-size": parse_int},
        )
        if parsed is None:
            return True
        values, flags, _positionals = parsed
        print_result(
            doc_server.format_paragraph(
                file_path=raw_args[0],
                paragraph_index=parse_int(raw_args[1], "paragraph_index"),
                bold=flags["bold"],
                italic=flags["italic"],
                underline=flags["underline"],
                font_name=values.get("font_name"),
                font_size=values.get("font_size"),
                color=values.get("color"),
                alignment=values.get("align"),
            ),
            json_output,
        )
        return True

    if command == "doc-highlight":
        if len(raw_args) < 2:
            print_result(
                {
                    "success": False,
                    "error": command_usage("doc-highlight"),
                },
                json_output,
            )
            return True
        if not docx_available:
            return _docx_unavailable(json_output, print_result)
        color = "yellow"
        parsed = _parse_options(
            command,
            raw_args,
            2,
            json_output,
            print_result,
            value_options=("--color",),
        )
        if parsed is None:
            return True
        values, _flags, _positionals = parsed
        color = values.get("color", color)
        print_result(doc_server.highlight_text(raw_args[0], raw_args[1], color=color), json_output)
        return True

    if command == "doc-comment":
        if len(raw_args) < 2:
            print_result(
                {
                    "success": False,
                    "error": command_usage("doc-comment"),
                },
                json_output,
            )
            return True
        if not docx_available:
            return _docx_unavailable(json_output, print_result)
        paragraph_index = 0
        parsed = _parse_options(
            command,
            raw_args,
            2,
            json_output,
            print_result,
            value_options=("--paragraph",),
            converters={"--paragraph": parse_int},
        )
        if parsed is None:
            return True
        values, _flags, _positionals = parsed
        paragraph_index = values.get("paragraph", paragraph_index)
        print_result(
            doc_server.add_comment(raw_args[0], raw_args[1], paragraph_index=paragraph_index),
            json_output,
        )
        return True

    if command == "doc-add-reference":
        if len(raw_args) < 2:
            print_result(
                {
                    "success": False,
                    "error": command_usage("doc-add-reference"),
                },
                json_output,
            )
            return True
        if _reject_extra_args(command, raw_args, 2, json_output, print_result):
            return True
        print_result(doc_server.add_reference(raw_args[0], raw_args[1]), json_output)
        return True

    if command == "doc-build-references":
        if not raw_args:
            print_result(
                {
                    "success": False,
                    "error": command_usage("doc-build-references"),
                },
                json_output,
            )
            return True
        if _reject_extra_args(command, raw_args, 1, json_output, print_result):
            return True
        print_result(doc_server.build_references(raw_args[0]), json_output)
        return True

    if command == "doc-add-table":
        if len(raw_args) < 3:
            print_result(
                {
                    "success": False,
                    "error": command_usage("doc-add-table"),
                },
                json_output,
            )
            return True
        if _reject_extra_args(command, raw_args, 3, json_output, print_result):
            return True
        print_result(doc_server.add_table(raw_args[0], raw_args[1], raw_args[2]), json_output)
        return True

    if command == "doc-set-style":
        if len(raw_args) < 3:
            print_result(
                {
                    "success": False,
                    "error": command_usage("doc-set-style"),
                },
                json_output,
            )
            return True
        if _reject_extra_args(command, raw_args, 3, json_output, print_result):
            return True
        print_result(
            doc_server.set_paragraph_style(
                raw_args[0], parse_int(raw_args[1], "paragraph_index"), raw_args[2]
            ),
            json_output,
        )
        return True

    if command == "doc-add-image":
        if len(raw_args) < 2:
            print_result(
                {
                    "success": False,
                    "error": command_usage("doc-add-image"),
                },
                json_output,
            )
            return True
        width = 5.5
        caption = None
        paragraph_index = None
        position = "after"
        parsed = _parse_options(
            command,
            raw_args,
            2,
            json_output,
            print_result,
            value_options=("--width", "--caption", "--paragraph", "--position"),
            converters={"--width": parse_float, "--paragraph": parse_int},
        )
        if parsed is None:
            return True
        values, _flags, _positionals = parsed
        width = values.get("width", width)
        caption = values.get("caption", caption)
        paragraph_index = values.get("paragraph", paragraph_index)
        position = values.get("position", position)
        print_result(
            doc_server.add_image(
                raw_args[0],
                raw_args[1],
                width_inches=width,
                caption=caption,
                paragraph_index=paragraph_index,
                position=position,
            ),
            json_output,
        )
        return True

    if command == "doc-layout":
        if not raw_args:
            print_result(
                {
                    "success": False,
                    "error": command_usage("doc-layout"),
                },
                json_output,
            )
            return True
        if not docx_available:
            return _docx_unavailable(json_output, print_result)
        page_size = None
        orientation = None
        margins = {}
        header_text = None
        page_numbers = False
        index = 1
        margin_options = {
            "--margin-top",
            "--margin-bottom",
            "--margin-left",
            "--margin-right",
        }
        while index < len(raw_args):
            token = raw_args[index]
            if token == "--size":
                value = _option_value(command, raw_args, index, json_output, print_result)
                if value is None:
                    return True
                page_size = value
                index += 2
            elif token == "--orientation":
                value = _option_value(command, raw_args, index, json_output, print_result)
                if value is None:
                    return True
                orientation = value
                index += 2
            elif token in margin_options:
                value = _option_value(command, raw_args, index, json_output, print_result)
                if value is None:
                    return True
                side = token.replace("--margin-", "")
                margins[side] = parse_float(value, token)
                index += 2
            elif token == "--header":
                value = _option_value(command, raw_args, index, json_output, print_result)
                if value is None:
                    return True
                header_text = value
                index += 2
            elif token == "--page-numbers":
                page_numbers = True
                index += 1
            elif _is_literal_arg(raw_args, index):
                if _doc_usage_error(
                    command,
                    json_output,
                    print_result,
                    f"Unexpected argument after -- for {command}: {token}",
                ):
                    return True
            elif _is_option_like(token):
                if _doc_usage_error(
                    command,
                    json_output,
                    print_result,
                    f"Unknown option for {command}: {token}",
                ):
                    return True
            else:
                if _doc_usage_error(
                    command,
                    json_output,
                    print_result,
                    f"Unexpected argument for {command}: {token}",
                ):
                    return True
        print_result(
            doc_server.set_page_layout(
                raw_args[0],
                page_size=page_size,
                orientation=orientation,
                margins=margins,
                header_text=header_text,
                page_numbers=page_numbers,
            ),
            json_output,
        )
        return True

    if command == "doc-formatting-info":
        if not raw_args:
            print_result(
                {"success": False, "error": command_usage("doc-formatting-info")},
                json_output,
            )
            return True
        if not docx_available:
            return _docx_unavailable(json_output, print_result)
        start = 0
        limit = 10
        all_paragraphs = False
        parsed = _parse_options(
            command,
            raw_args,
            1,
            json_output,
            print_result,
            value_options=("--start", "--limit"),
            flag_options=("--all",),
            converters={"--start": parse_int, "--limit": parse_int},
        )
        if parsed is None:
            return True
        values, flags, _positionals = parsed
        start = values.get("start", start)
        limit = values.get("limit", limit)
        all_paragraphs = flags["all"]
        print_result(
            doc_server.get_formatting_info(
                raw_args[0],
                start=start,
                limit=limit,
                all_paragraphs=all_paragraphs,
            ),
            json_output,
        )
        return True

    if command == "doc-search":
        if len(raw_args) < 2:
            print_result(
                {"success": False, "error": command_usage("doc-search")},
                json_output,
            )
            return True
        parsed = _parse_options(
            command,
            raw_args,
            2,
            json_output,
            print_result,
            flag_options=("--case-sensitive",),
        )
        if parsed is None:
            return True
        _values, flags, _positionals = parsed
        print_result(
            doc_server.search_document(
                raw_args[0],
                raw_args[1],
                case_sensitive=flags["case_sensitive"],
            ),
            json_output,
        )
        return True

    if command == "doc-insert":
        if len(raw_args) < 3:
            print_result(
                {"success": False, "error": command_usage("doc-insert")},
                json_output,
            )
            return True
        style = None
        parsed = _parse_options(
            command,
            raw_args,
            3,
            json_output,
            print_result,
            value_options=("--style",),
        )
        if parsed is None:
            return True
        values, _flags, _positionals = parsed
        style = values.get("style", style)
        print_result(
            doc_server.insert_paragraph(
                raw_args[0],
                raw_args[1],
                parse_int(raw_args[2], "index"),
                style=style,
            ),
            json_output,
        )
        return True

    if command == "doc-delete":
        if len(raw_args) < 2:
            print_result(
                {"success": False, "error": command_usage("doc-delete")},
                json_output,
            )
            return True
        if _reject_extra_args(command, raw_args, 2, json_output, print_result):
            return True
        print_result(
            doc_server.delete_paragraph(
                raw_args[0], parse_int(raw_args[1], "paragraph_index")
            ),
            json_output,
        )
        return True

    if command == "doc-read-tables":
        if not raw_args:
            print_result(
                {"success": False, "error": command_usage("doc-read-tables")},
                json_output,
            )
            return True
        if _reject_extra_args(command, raw_args, 1, json_output, print_result):
            return True
        print_result(doc_server.read_tables(raw_args[0]), json_output)
        return True

    if command == "doc-add-hyperlink":
        if len(raw_args) < 3:
            print_result(
                {
                    "success": False,
                    "error": command_usage("doc-add-hyperlink"),
                },
                json_output,
            )
            return True
        paragraph_index = -1
        parsed = _parse_options(
            command,
            raw_args,
            3,
            json_output,
            print_result,
            value_options=("--paragraph",),
            converters={"--paragraph": parse_int},
        )
        if parsed is None:
            return True
        values, _flags, _positionals = parsed
        paragraph_index = values.get("paragraph", paragraph_index)
        print_result(
            doc_server.add_hyperlink(
                raw_args[0],
                raw_args[1],
                raw_args[2],
                paragraph_index=paragraph_index,
            ),
            json_output,
        )
        return True

    if command == "doc-add-page-break":
        if not raw_args:
            print_result(
                {"success": False, "error": command_usage("doc-add-page-break")},
                json_output,
            )
            return True
        if _reject_extra_args(command, raw_args, 1, json_output, print_result):
            return True
        print_result(doc_server.add_page_break(raw_args[0]), json_output)
        return True

    if command == "doc-list-styles":
        if not raw_args:
            print_result(
                {"success": False, "error": command_usage("doc-list-styles")},
                json_output,
            )
            return True
        if _reject_extra_args(command, raw_args, 1, json_output, print_result):
            return True
        print_result(doc_server.list_styles(raw_args[0]), json_output)
        return True

    if command == "doc-add-list":
        if len(raw_args) < 2:
            print_result(
                {
                    "success": False,
                    "error": command_usage("doc-add-list"),
                },
                json_output,
            )
            return True
        list_type = "bullet"
        parsed = _parse_options(
            command,
            raw_args,
            2,
            json_output,
            print_result,
            value_options=("--type",),
        )
        if parsed is None:
            return True
        values, _flags, _positionals = parsed
        list_type = values.get("type", list_type)
        items = [item.strip() for item in raw_args[1].split(";") if item.strip()]
        print_result(doc_server.add_list(raw_args[0], items, list_type=list_type), json_output)
        return True

    if command == "doc-set-metadata":
        if not raw_args:
            print_result(
                {
                    "success": False,
                    "error": command_usage("doc-set-metadata"),
                },
                json_output,
            )
            return True
        parsed = _parse_options(
            command,
            raw_args,
            1,
            json_output,
            print_result,
            value_options=(
                "--author",
                "--title",
                "--subject",
                "--keywords",
                "--comments",
                "--category",
            ),
        )
        if parsed is None:
            return True
        options, _flags, _positionals = parsed
        print_result(doc_server.set_metadata(raw_args[0], **options), json_output)
        return True

    if command == "doc-get-metadata":
        if not raw_args:
            print_result(
                {"success": False, "error": command_usage("doc-get-metadata")},
                json_output,
            )
            return True
        if _reject_extra_args(command, raw_args, 1, json_output, print_result):
            return True
        print_result(doc_server.get_metadata(raw_args[0]), json_output)
        return True

    if command == "doc-inspect-hidden-data":
        if not raw_args:
            print_result(
                {"success": False, "error": command_usage("doc-inspect-hidden-data")},
                json_output,
            )
            return True
        if _reject_extra_args(command, raw_args, 1, json_output, print_result):
            return True
        print_result(doc_server.inspect_hidden_data(raw_args[0]), json_output)
        return True

    if command == "doc-sanitize":
        if not raw_args:
            print_result(
                {
                    "success": False,
                    "error": command_usage("doc-sanitize"),
                },
                json_output,
            )
            return True
        output_path = None
        remove_comments = False
        accept_revisions = False
        clear_metadata = False
        remove_custom_xml = False
        set_remove_personal_information = False
        canonicalize_ooxml = False
        author = None
        title = None
        subject = None
        keywords = None
        parsed = _parse_options(
            command,
            raw_args,
            1,
            json_output,
            print_result,
            value_options=("--author", "--title", "--subject", "--keywords"),
            flag_options=(
                "--remove-comments",
                "--accept-revisions",
                "--clear-metadata",
                "--remove-custom-xml",
                "--set-remove-personal-information",
                "--canonicalize-ooxml",
            ),
            max_positionals=1,
        )
        if parsed is None:
            return True
        values, flags, positionals = parsed
        if positionals:
            output_path = positionals[0]
        remove_comments = flags["remove_comments"]
        accept_revisions = flags["accept_revisions"]
        clear_metadata = flags["clear_metadata"]
        remove_custom_xml = flags["remove_custom_xml"]
        set_remove_personal_information = flags["set_remove_personal_information"]
        canonicalize_ooxml = flags["canonicalize_ooxml"]
        author = values.get("author", author)
        title = values.get("title", title)
        subject = values.get("subject", subject)
        keywords = values.get("keywords", keywords)
        print_result(
            doc_server.sanitize_document(
                raw_args[0],
                output_path=output_path,
                remove_comments=remove_comments,
                accept_revisions=accept_revisions,
                clear_metadata=clear_metadata,
                remove_custom_xml=remove_custom_xml,
                set_remove_personal_information=set_remove_personal_information,
                canonicalize_ooxml=canonicalize_ooxml,
                author=author,
                title=title,
                subject=subject,
                keywords=keywords,
            ),
            json_output,
        )
        return True

    if command == "doc-preflight":
        if not raw_args:
            print_result(
                {
                    "success": False,
                    "error": command_usage("doc-preflight"),
                },
                json_output,
            )
            return True
        expected_page_size = None
        expected_font = None
        expected_font_size = None
        rendered_layout = False
        render_profile = "auto"
        parsed = _parse_options(
            command,
            raw_args,
            1,
            json_output,
            print_result,
            value_options=(
                "--expected-page-size",
                "--expected-font",
                "--expected-font-size",
                "--profile",
            ),
            flag_options=("--rendered-layout",),
            converters={"--expected-font-size": parse_float},
        )
        if parsed is None:
            return True
        values, flags, _positionals = parsed
        expected_page_size = values.get("expected_page_size", expected_page_size)
        expected_font = values.get("expected_font", expected_font)
        expected_font_size = values.get("expected_font_size", expected_font_size)
        rendered_layout = flags["rendered_layout"]
        render_profile = values.get("profile", render_profile)
        print_result(
            doc_server.document_preflight(
                raw_args[0],
                expected_page_size=expected_page_size,
                expected_font_name=expected_font,
                expected_font_size=expected_font_size,
                rendered_layout=rendered_layout,
                render_profile=render_profile,
            ),
            json_output,
        )
        return True

    if command == "doc-word-count":
        if not raw_args:
            print_result(
                {"success": False, "error": command_usage("doc-word-count")},
                json_output,
            )
            return True
        if _reject_extra_args(command, raw_args, 1, json_output, print_result):
            return True
        print_result(doc_server.word_count(raw_args[0]), json_output)
        return True

    if command == "doc-extract-images":
        if len(raw_args) < 2:
            print_result(
                {
                    "success": False,
                    "error": command_usage("doc-extract-images"),
                },
                json_output,
            )
            return True
        fmt = "png"
        prefix = "image"
        parsed = _parse_options(
            command,
            raw_args,
            2,
            json_output,
            print_result,
            value_options=("--format", "--prefix"),
        )
        if parsed is None:
            return True
        values, _flags, _positionals = parsed
        fmt = values.get("format", fmt)
        prefix = values.get("prefix", prefix)
        print_result(
            doc_server.extract_images_from_docx(raw_args[0], raw_args[1], fmt=fmt, prefix=prefix),
            json_output,
        )
        return True

    if command == "doc-to-pdf":
        if not raw_args:
            print_result(
                {"success": False, "error": command_usage("doc-to-pdf")},
                json_output,
            )
            return True
        output_path = None
        layout_warnings = False
        render_profile = "auto"
        parsed = _parse_options(
            command,
            raw_args,
            1,
            json_output,
            print_result,
            value_options=("--profile",),
            flag_options=("--layout-warnings",),
            max_positionals=1,
        )
        if parsed is None:
            return True
        values, flags, positionals = parsed
        if positionals:
            output_path = positionals[0]
        layout_warnings = flags["layout_warnings"]
        render_profile = values.get("profile", render_profile)
        print_result(
            doc_server.doc_to_pdf(
                raw_args[0],
                output_path=output_path,
                layout_warnings=layout_warnings,
                render_profile=render_profile,
            ),
            json_output,
        )
        return True

    if command == "doc-preview":
        if len(raw_args) < 2:
            print_result(
                {
                    "success": False,
                    "error": command_usage("doc-preview"),
                },
                json_output,
            )
            return True
        pages = None
        dpi = 150
        fmt = "png"
        parsed = _parse_options(
            command,
            raw_args,
            2,
            json_output,
            print_result,
            value_options=("--pages", "--dpi", "--format"),
            converters={"--dpi": parse_int},
        )
        if parsed is None:
            return True
        values, _flags, _positionals = parsed
        pages = values.get("pages", pages)
        dpi = values.get("dpi", dpi)
        fmt = values.get("format", fmt)
        print_result(
            doc_server.preview_document(raw_args[0], raw_args[1], pages=pages, dpi=dpi, fmt=fmt),
            json_output,
        )
        return True

    if command == "doc-render-map":
        if not raw_args:
            print_result(
                {"success": False, "error": command_usage("doc-render-map")},
                json_output,
            )
            return True
        if not docx_available:
            return _docx_unavailable(json_output, print_result)
        if _reject_extra_args(command, raw_args, 1, json_output, print_result):
            return True
        print_result(doc_server.doc_render_map(raw_args[0]), json_output)
        return True

    if command == "doc-render-audit":
        if not raw_args:
            print_result(
                {"success": False, "error": command_usage("doc-render-audit")},
                json_output,
            )
            return True
        if not docx_available:
            return _docx_unavailable(json_output, print_result)
        pdf_path = None
        tolerance_points = 6.0
        profile = "auto"
        parsed = _parse_options(
            command,
            raw_args,
            1,
            json_output,
            print_result,
            value_options=("--pdf", "--tolerance-points", "--profile"),
            converters={"--tolerance-points": parse_float},
        )
        if parsed is None:
            return True
        values, _flags, _positionals = parsed
        pdf_path = values.get("pdf", pdf_path)
        tolerance_points = values.get("tolerance_points", tolerance_points)
        profile = values.get("profile", profile)
        print_result(
            doc_server.rendered_layout_audit(
                raw_args[0],
                pdf_path=pdf_path,
                tolerance_points=tolerance_points,
                profile=profile,
            ),
            json_output,
        )
        return True

    return False
