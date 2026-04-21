#!/usr/bin/env python3
"""PPTX CLI command handlers."""

from __future__ import annotations

from typing import Any, Callable, List

from cli_anything.onlyoffice.core.command_registry import command_usage


def handle_pptx_command(
    command: str,
    raw_args: List[str],
    doc_server: Any,
    json_output: bool,
    print_result: Callable[[dict, bool], None],
) -> bool:
    """Handle PPTX commands and return True when the command was recognised."""
    if command == "pptx-create":
        if len(raw_args) < 2:
            print_result(
                {
                    "success": False,
                    "error": command_usage("pptx-create"),
                },
                json_output,
            )
            return True
        print_result(
            doc_server.create_presentation(
                raw_args[0],
                raw_args[1],
                raw_args[2] if len(raw_args) > 2 else "",
            ),
            json_output,
        )
        return True

    if command == "pptx-add-slide":
        if len(raw_args) < 2:
            print_result(
                {
                    "success": False,
                    "error": command_usage("pptx-add-slide"),
                },
                json_output,
            )
            return True
        print_result(
            doc_server.add_slide(
                raw_args[0],
                raw_args[1],
                raw_args[2] if len(raw_args) > 2 else "",
                raw_args[3] if len(raw_args) > 3 else "content",
            ),
            json_output,
        )
        return True

    if command == "pptx-add-bullets":
        if len(raw_args) < 3:
            print_result(
                {
                    "success": False,
                    "error": command_usage("pptx-add-bullets"),
                },
                json_output,
            )
            return True
        print_result(
            doc_server.add_bullet_slide(
                raw_args[0],
                raw_args[1],
                raw_args[2].replace("\\n", "\n"),
            ),
            json_output,
        )
        return True

    if command == "pptx-read":
        if not raw_args:
            print_result({"success": False, "error": command_usage("pptx-read")}, json_output)
            return True
        print_result(doc_server.read_presentation(raw_args[0]), json_output)
        return True

    if command == "pptx-add-image":
        if len(raw_args) < 3:
            print_result(
                {
                    "success": False,
                    "error": command_usage("pptx-add-image"),
                },
                json_output,
            )
            return True
        print_result(
            doc_server.add_image_slide(raw_args[0], raw_args[1], raw_args[2]),
            json_output,
        )
        return True

    if command == "pptx-add-table":
        if len(raw_args) < 4:
            print_result(
                {
                    "success": False,
                    "error": command_usage("pptx-add-table"),
                },
                json_output,
            )
            return True
        print_result(
            doc_server.add_table_slide(
                raw_args[0],
                raw_args[1],
                raw_args[2],
                raw_args[3],
                coerce_rows="--coerce-rows" in raw_args[4:],
            ),
            json_output,
        )
        return True

    if command == "pptx-delete-slide":
        if len(raw_args) < 2:
            print_result(
                {"success": False, "error": command_usage("pptx-delete-slide")},
                json_output,
            )
            return True
        print_result(doc_server.delete_slide(raw_args[0], int(raw_args[1])), json_output)
        return True

    if command == "pptx-speaker-notes":
        if len(raw_args) < 2:
            print_result(
                {
                    "success": False,
                    "error": command_usage("pptx-speaker-notes"),
                },
                json_output,
            )
            return True
        notes = " ".join(raw_args[2:]) if len(raw_args) > 2 else None
        print_result(
            doc_server.speaker_notes(raw_args[0], int(raw_args[1]), notes_text=notes),
            json_output,
        )
        return True

    if command == "pptx-update-text":
        if len(raw_args) < 2:
            print_result(
                {
                    "success": False,
                    "error": command_usage("pptx-update-text"),
                },
                json_output,
            )
            return True
        title = None
        body = None
        index = 2
        while index < len(raw_args):
            if raw_args[index] == "--title" and index + 1 < len(raw_args):
                title = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--body" and index + 1 < len(raw_args):
                body = raw_args[index + 1]
                index += 2
            else:
                index += 1
        print_result(
            doc_server.update_slide_text(raw_args[0], int(raw_args[1]), title=title, body=body),
            json_output,
        )
        return True

    if command == "pptx-slide-count":
        if not raw_args:
            print_result(
                {"success": False, "error": command_usage("pptx-slide-count")},
                json_output,
            )
            return True
        print_result(doc_server.slide_count(raw_args[0]), json_output)
        return True

    if command == "pptx-extract-images":
        if len(raw_args) < 2:
            print_result(
                {
                    "success": False,
                    "error": command_usage("pptx-extract-images"),
                },
                json_output,
            )
            return True
        fmt = "png"
        slide_index = None
        prefix = "slide"
        index = 2
        while index < len(raw_args):
            if raw_args[index] == "--format" and index + 1 < len(raw_args):
                fmt = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--slide" and index + 1 < len(raw_args):
                slide_index = int(raw_args[index + 1])
                index += 2
            elif raw_args[index] == "--prefix" and index + 1 < len(raw_args):
                prefix = raw_args[index + 1]
                index += 2
            else:
                index += 1
        print_result(
            doc_server.extract_images_from_pptx(
                raw_args[0],
                raw_args[1],
                slide_index=slide_index,
                fmt=fmt,
                prefix=prefix,
            ),
            json_output,
        )
        return True

    if command == "pptx-list-shapes":
        if not raw_args:
            print_result(
                {
                    "success": False,
                    "error": command_usage("pptx-list-shapes"),
                },
                json_output,
            )
            return True
        slide_index = None
        index = 1
        while index < len(raw_args):
            if raw_args[index] == "--slide" and index + 1 < len(raw_args):
                slide_index = int(raw_args[index + 1])
                index += 2
            else:
                index += 1
        print_result(doc_server.list_shapes(raw_args[0], slide_index=slide_index), json_output)
        return True

    if command == "pptx-add-textbox":
        if len(raw_args) < 3:
            print_result(
                {
                    "success": False,
                    "error": command_usage("pptx-add-textbox"),
                },
                json_output,
            )
            return True
        file_path = raw_args[0]
        slide_index = int(raw_args[1])
        text = raw_args[2]
        left = 1.0
        top = 1.0
        width = 5.0
        height = 1.5
        font_size = None
        font_name = None
        bold = False
        italic = False
        color = None
        align = None
        index = 3
        while index < len(raw_args):
            if raw_args[index] == "--left" and index + 1 < len(raw_args):
                left = float(raw_args[index + 1])
                index += 2
            elif raw_args[index] == "--top" and index + 1 < len(raw_args):
                top = float(raw_args[index + 1])
                index += 2
            elif raw_args[index] == "--width" and index + 1 < len(raw_args):
                width = float(raw_args[index + 1])
                index += 2
            elif raw_args[index] == "--height" and index + 1 < len(raw_args):
                height = float(raw_args[index + 1])
                index += 2
            elif raw_args[index] == "--font-size" and index + 1 < len(raw_args):
                font_size = float(raw_args[index + 1])
                index += 2
            elif raw_args[index] == "--font-name" and index + 1 < len(raw_args):
                font_name = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--bold":
                bold = True
                index += 1
            elif raw_args[index] == "--italic":
                italic = True
                index += 1
            elif raw_args[index] == "--color" and index + 1 < len(raw_args):
                color = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--align" and index + 1 < len(raw_args):
                align = raw_args[index + 1]
                index += 2
            else:
                index += 1
        print_result(
            doc_server.add_textbox(
                file_path,
                slide_index,
                text,
                left=left,
                top=top,
                width=width,
                height=height,
                font_size=font_size,
                font_name=font_name,
                bold=bold,
                italic=italic,
                color=color,
                align=align,
            ),
            json_output,
        )
        return True

    if command == "pptx-modify-shape":
        if len(raw_args) < 3:
            print_result(
                {
                    "success": False,
                    "error": command_usage("pptx-modify-shape"),
                },
                json_output,
            )
            return True
        file_path = raw_args[0]
        slide_index = int(raw_args[1])
        shape_name = raw_args[2]
        left = None
        top = None
        width = None
        height = None
        text = None
        font_size = None
        rotation = None
        index = 3
        while index < len(raw_args):
            if raw_args[index] == "--left" and index + 1 < len(raw_args):
                left = float(raw_args[index + 1])
                index += 2
            elif raw_args[index] == "--top" and index + 1 < len(raw_args):
                top = float(raw_args[index + 1])
                index += 2
            elif raw_args[index] == "--width" and index + 1 < len(raw_args):
                width = float(raw_args[index + 1])
                index += 2
            elif raw_args[index] == "--height" and index + 1 < len(raw_args):
                height = float(raw_args[index + 1])
                index += 2
            elif raw_args[index] == "--text" and index + 1 < len(raw_args):
                text = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--font-size" and index + 1 < len(raw_args):
                font_size = float(raw_args[index + 1])
                index += 2
            elif raw_args[index] == "--rotation" and index + 1 < len(raw_args):
                rotation = float(raw_args[index + 1])
                index += 2
            else:
                index += 1
        print_result(
            doc_server.modify_shape(
                file_path,
                slide_index,
                shape_name,
                left=left,
                top=top,
                width=width,
                height=height,
                text=text,
                font_size=font_size,
                rotation=rotation,
            ),
            json_output,
        )
        return True

    if command == "pptx-preview":
        if len(raw_args) < 2:
            print_result(
                {
                    "success": False,
                    "error": command_usage("pptx-preview"),
                },
                json_output,
            )
            return True
        slide_index = None
        dpi = 150
        index = 2
        while index < len(raw_args):
            if raw_args[index] == "--slide" and index + 1 < len(raw_args):
                slide_index = int(raw_args[index + 1])
                index += 2
            elif raw_args[index] == "--dpi" and index + 1 < len(raw_args):
                dpi = int(raw_args[index + 1])
                index += 2
            else:
                index += 1
        print_result(
            doc_server.preview_slide(raw_args[0], raw_args[1], slide_index=slide_index, dpi=dpi),
            json_output,
        )
        return True

    return False
