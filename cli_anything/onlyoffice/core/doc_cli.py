#!/usr/bin/env python3
"""DOCX CLI command handlers."""

from __future__ import annotations

from typing import Any, Callable, List

from cli_anything.onlyoffice.core.command_registry import command_usage


def _docx_unavailable(
    json_output: bool,
    print_result: Callable[[dict, bool], None],
) -> bool:
    print_result(
        {"success": False, "error": "python-docx not installed"},
        json_output,
    )
    return True


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
        options = {
            "bold": False,
            "italic": False,
            "underline": False,
            "font_name": None,
            "font_size": None,
            "color": None,
            "alignment": None,
        }
        index = 2
        while index < len(raw_args):
            if raw_args[index] == "--bold":
                options["bold"] = True
                index += 1
            elif raw_args[index] == "--italic":
                options["italic"] = True
                index += 1
            elif raw_args[index] == "--underline":
                options["underline"] = True
                index += 1
            elif raw_args[index] == "--font-name" and index + 1 < len(raw_args):
                options["font_name"] = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--font-size" and index + 1 < len(raw_args):
                options["font_size"] = int(raw_args[index + 1])
                index += 2
            elif raw_args[index] == "--color" and index + 1 < len(raw_args):
                options["color"] = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--align" and index + 1 < len(raw_args):
                options["alignment"] = raw_args[index + 1]
                index += 2
            else:
                index += 1
        print_result(
            doc_server.format_paragraph(
                file_path=raw_args[0],
                paragraph_index=int(raw_args[1]),
                bold=options["bold"],
                italic=options["italic"],
                underline=options["underline"],
                font_name=options["font_name"],
                font_size=options["font_size"],
                color=options["color"],
                alignment=options["alignment"],
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
        if len(raw_args) >= 4 and raw_args[2] == "--color":
            color = raw_args[3]
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
        if len(raw_args) >= 4 and raw_args[2] == "--paragraph":
            paragraph_index = int(raw_args[3])
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
        print_result(
            doc_server.set_paragraph_style(raw_args[0], int(raw_args[1]), raw_args[2]),
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
        index = 2
        while index < len(raw_args):
            if raw_args[index] == "--width" and index + 1 < len(raw_args):
                width = float(raw_args[index + 1])
                index += 2
            elif raw_args[index] == "--caption" and index + 1 < len(raw_args):
                caption = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--paragraph" and index + 1 < len(raw_args):
                paragraph_index = int(raw_args[index + 1])
                index += 2
            elif raw_args[index] == "--position" and index + 1 < len(raw_args):
                position = raw_args[index + 1]
                index += 2
            else:
                index += 1
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
        while index < len(raw_args):
            if raw_args[index] == "--size" and index + 1 < len(raw_args):
                page_size = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--orientation" and index + 1 < len(raw_args):
                orientation = raw_args[index + 1]
                index += 2
            elif raw_args[index].startswith("--margin-") and index + 1 < len(raw_args):
                side = raw_args[index].replace("--margin-", "")
                margins[side] = float(raw_args[index + 1])
                index += 2
            elif raw_args[index] == "--header" and index + 1 < len(raw_args):
                header_text = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--page-numbers":
                page_numbers = True
                index += 1
            else:
                index += 1
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
        print_result(doc_server.get_formatting_info(raw_args[0]), json_output)
        return True

    if command == "doc-search":
        if len(raw_args) < 2:
            print_result(
                {"success": False, "error": command_usage("doc-search")},
                json_output,
            )
            return True
        print_result(
            doc_server.search_document(
                raw_args[0],
                raw_args[1],
                case_sensitive="--case-sensitive" in raw_args[2:],
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
        index = 3
        while index < len(raw_args):
            if raw_args[index] == "--style" and index + 1 < len(raw_args):
                style = raw_args[index + 1]
                index += 2
            else:
                index += 1
        print_result(
            doc_server.insert_paragraph(raw_args[0], raw_args[1], int(raw_args[2]), style=style),
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
        print_result(doc_server.delete_paragraph(raw_args[0], int(raw_args[1])), json_output)
        return True

    if command == "doc-read-tables":
        if not raw_args:
            print_result(
                {"success": False, "error": command_usage("doc-read-tables")},
                json_output,
            )
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
        index = 3
        while index < len(raw_args):
            if raw_args[index] == "--paragraph" and index + 1 < len(raw_args):
                paragraph_index = int(raw_args[index + 1])
                index += 2
            else:
                index += 1
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
        print_result(doc_server.add_page_break(raw_args[0]), json_output)
        return True

    if command == "doc-list-styles":
        if not raw_args:
            print_result(
                {"success": False, "error": command_usage("doc-list-styles")},
                json_output,
            )
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
        index = 2
        while index < len(raw_args):
            if raw_args[index] == "--type" and index + 1 < len(raw_args):
                list_type = raw_args[index + 1]
                index += 2
            else:
                index += 1
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
        options = {}
        index = 1
        while index < len(raw_args):
            for key in ("--author", "--title", "--subject", "--keywords", "--comments", "--category"):
                if raw_args[index] == key and index + 1 < len(raw_args):
                    options[key.lstrip("-")] = raw_args[index + 1]
                    index += 2
                    break
            else:
                index += 1
        print_result(doc_server.set_metadata(raw_args[0], **options), json_output)
        return True

    if command == "doc-get-metadata":
        if not raw_args:
            print_result(
                {"success": False, "error": command_usage("doc-get-metadata")},
                json_output,
            )
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
        author = None
        title = None
        subject = None
        keywords = None
        index = 1
        while index < len(raw_args):
            if raw_args[index] == "--remove-comments":
                remove_comments = True
                index += 1
            elif raw_args[index] == "--accept-revisions":
                accept_revisions = True
                index += 1
            elif raw_args[index] == "--clear-metadata":
                clear_metadata = True
                index += 1
            elif raw_args[index] == "--remove-custom-xml":
                remove_custom_xml = True
                index += 1
            elif raw_args[index] == "--set-remove-personal-information":
                set_remove_personal_information = True
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
            elif output_path is None and not raw_args[index].startswith("--"):
                output_path = raw_args[index]
                index += 1
            else:
                index += 1
        print_result(
            doc_server.sanitize_document(
                raw_args[0],
                output_path=output_path,
                remove_comments=remove_comments,
                accept_revisions=accept_revisions,
                clear_metadata=clear_metadata,
                remove_custom_xml=remove_custom_xml,
                set_remove_personal_information=set_remove_personal_information,
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
        index = 1
        while index < len(raw_args):
            if raw_args[index] == "--expected-page-size" and index + 1 < len(raw_args):
                expected_page_size = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--expected-font" and index + 1 < len(raw_args):
                expected_font = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--expected-font-size" and index + 1 < len(raw_args):
                expected_font_size = float(raw_args[index + 1])
                index += 2
            else:
                index += 1
        print_result(
            doc_server.document_preflight(
                raw_args[0],
                expected_page_size=expected_page_size,
                expected_font_name=expected_font,
                expected_font_size=expected_font_size,
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
        index = 2
        while index < len(raw_args):
            if raw_args[index] == "--format" and index + 1 < len(raw_args):
                fmt = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--prefix" and index + 1 < len(raw_args):
                prefix = raw_args[index + 1]
                index += 2
            else:
                index += 1
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
        output_path = raw_args[1] if len(raw_args) >= 2 else None
        print_result(doc_server.doc_to_pdf(raw_args[0], output_path=output_path), json_output)
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
        print_result(doc_server.doc_render_map(raw_args[0]), json_output)
        return True

    return False
