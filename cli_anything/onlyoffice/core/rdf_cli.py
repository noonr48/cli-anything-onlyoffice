#!/usr/bin/env python3
"""RDF CLI command handlers."""

from __future__ import annotations

from typing import Any, Callable, List

from cli_anything.onlyoffice.core.command_registry import command_usage


def handle_rdf_command(
    command: str,
    raw_args: List[str],
    doc_server: Any,
    json_output: bool,
    print_result: Callable[[dict, bool], None],
) -> bool:
    """Handle RDF commands and return True when the command was recognised."""
    if command == "rdf-create":
        if not raw_args:
            print_result(
                {
                    "success": False,
                    "error": command_usage("rdf-create"),
                },
                json_output,
            )
            return True
        base_uri = None
        fmt = "turtle"
        prefixes = {}
        index = 1
        while index < len(raw_args):
            if raw_args[index] == "--base" and index + 1 < len(raw_args):
                base_uri = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--format" and index + 1 < len(raw_args):
                fmt = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--prefix" and index + 1 < len(raw_args):
                parts = raw_args[index + 1].split("=", 1)
                if len(parts) == 2:
                    prefixes[parts[0]] = parts[1]
                index += 2
            else:
                index += 1
        print_result(
            doc_server.rdf_create(raw_args[0], base_uri=base_uri, format=fmt, prefixes=prefixes),
            json_output,
        )
        return True

    if command == "rdf-read":
        if not raw_args:
            print_result({"success": False, "error": command_usage("rdf-read")}, json_output)
            return True
        limit = 100
        if len(raw_args) > 2 and raw_args[1] == "--limit":
            limit = int(raw_args[2])
        print_result(doc_server.rdf_read(raw_args[0], limit=limit), json_output)
        return True

    if command == "rdf-add":
        if len(raw_args) < 4:
            print_result(
                {
                    "success": False,
                    "error": command_usage("rdf-add"),
                },
                json_output,
            )
            return True
        obj_type = "uri"
        lang = None
        datatype = None
        fmt = "turtle"
        index = 4
        while index < len(raw_args):
            if raw_args[index] == "--type" and index + 1 < len(raw_args):
                obj_type = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--lang" and index + 1 < len(raw_args):
                lang = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--datatype" and index + 1 < len(raw_args):
                datatype = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--format" and index + 1 < len(raw_args):
                fmt = raw_args[index + 1]
                index += 2
            else:
                index += 1
        print_result(
            doc_server.rdf_add(
                raw_args[0],
                raw_args[1],
                raw_args[2],
                raw_args[3],
                object_type=obj_type,
                lang=lang,
                datatype=datatype,
                format=fmt,
            ),
            json_output,
        )
        return True

    if command == "rdf-remove":
        if not raw_args:
            print_result(
                {
                    "success": False,
                    "error": command_usage("rdf-remove"),
                },
                json_output,
            )
            return True
        subject = None
        predicate = None
        object_val = None
        object_type = "uri"
        lang = None
        datatype = None
        fmt = None
        index = 1
        while index < len(raw_args):
            if raw_args[index] == "--subject" and index + 1 < len(raw_args):
                subject = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--predicate" and index + 1 < len(raw_args):
                predicate = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--object" and index + 1 < len(raw_args):
                object_val = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--type" and index + 1 < len(raw_args):
                object_type = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--lang" and index + 1 < len(raw_args):
                lang = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--datatype" and index + 1 < len(raw_args):
                datatype = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--format" and index + 1 < len(raw_args):
                fmt = raw_args[index + 1]
                index += 2
            else:
                index += 1
        print_result(
            doc_server.rdf_remove(
                raw_args[0],
                subject=subject,
                predicate=predicate,
                object_val=object_val,
                object_type=object_type,
                lang=lang,
                datatype=datatype,
                format=fmt,
            ),
            json_output,
        )
        return True

    if command == "rdf-query":
        if len(raw_args) < 2:
            print_result(
                {"success": False, "error": command_usage("rdf-query")},
                json_output,
            )
            return True
        limit = 100
        index = 2
        while index < len(raw_args):
            if raw_args[index] == "--limit" and index + 1 < len(raw_args):
                limit = int(raw_args[index + 1])
                index += 2
            else:
                index += 1
        print_result(doc_server.rdf_query(raw_args[0], raw_args[1], limit=limit), json_output)
        return True

    if command == "rdf-export":
        if len(raw_args) < 2:
            print_result(
                {
                    "success": False,
                    "error": command_usage("rdf-export"),
                },
                json_output,
            )
            return True
        fmt = "turtle"
        if len(raw_args) > 3 and raw_args[2] == "--format":
            fmt = raw_args[3]
        print_result(
            doc_server.rdf_export(raw_args[0], raw_args[1], output_format=fmt),
            json_output,
        )
        return True

    if command == "rdf-merge":
        if len(raw_args) < 2:
            print_result(
                {
                    "success": False,
                    "error": command_usage("rdf-merge"),
                },
                json_output,
            )
            return True
        output = None
        fmt = "turtle"
        index = 2
        while index < len(raw_args):
            if raw_args[index] == "--output" and index + 1 < len(raw_args):
                output = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--format" and index + 1 < len(raw_args):
                fmt = raw_args[index + 1]
                index += 2
            else:
                index += 1
        print_result(
            doc_server.rdf_merge(raw_args[0], raw_args[1], output_path=output, format=fmt),
            json_output,
        )
        return True

    if command == "rdf-stats":
        if not raw_args:
            print_result({"success": False, "error": command_usage("rdf-stats")}, json_output)
            return True
        print_result(doc_server.rdf_stats(raw_args[0]), json_output)
        return True

    if command == "rdf-namespace":
        if not raw_args:
            print_result(
                {"success": False, "error": command_usage("rdf-namespace")},
                json_output,
            )
            return True
        prefix = None
        uri = None
        fmt = "turtle"
        if len(raw_args) >= 3 and not raw_args[1].startswith("--"):
            prefix = raw_args[1]
            uri = raw_args[2]
        index = 1
        while index < len(raw_args):
            if raw_args[index] == "--format" and index + 1 < len(raw_args):
                fmt = raw_args[index + 1]
                index += 2
            else:
                index += 1
        print_result(
            doc_server.rdf_namespace(raw_args[0], prefix=prefix, uri=uri, format=fmt),
            json_output,
        )
        return True

    if command == "rdf-validate":
        if len(raw_args) < 2:
            print_result(
                {"success": False, "error": command_usage("rdf-validate")},
                json_output,
            )
            return True
        print_result(doc_server.rdf_validate(raw_args[0], raw_args[1]), json_output)
        return True

    return False
