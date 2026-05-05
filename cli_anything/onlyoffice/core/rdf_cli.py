#!/usr/bin/env python3
"""RDF CLI command handlers."""

from __future__ import annotations

from typing import Any, Callable, List

from cli_anything.onlyoffice.core.command_registry import command_usage
from cli_anything.onlyoffice.core.parse_utils import parse_int, print_usage_error


VALID_RDF_OBJECT_TYPES = {"uri", "literal", "bnode"}


def _option_value(command: str, raw_args: List[str], index: int, option: str) -> str:
    if index + 1 >= len(raw_args) or raw_args[index + 1].startswith("--"):
        raise ValueError(
            f"{command}: {option} requires a value. Usage: {command_usage(command)}"
        )
    return raw_args[index + 1]


def _raise_unexpected_argument(command: str, token: str) -> None:
    if token.startswith("--"):
        raise ValueError(
            f"{command}: unknown option {token!r}. Usage: {command_usage(command)}"
        )
    raise ValueError(
        f"{command}: unexpected argument {token!r}. Usage: {command_usage(command)}"
    )


def _validate_object_type(command: str, object_type: str) -> None:
    if object_type not in VALID_RDF_OBJECT_TYPES:
        allowed = ", ".join(sorted(VALID_RDF_OBJECT_TYPES))
        raise ValueError(
            f"{command}: --type must be one of {allowed}; got {object_type!r}."
        )


def _rdf_remove_dry_run(
    file_path: str,
    subject: str = None,
    predicate: str = None,
    object_val: str = None,
    object_type: str = "uri",
    lang: str = None,
    datatype: str = None,
    fmt: str = None,
) -> dict:
    """Count RDF triples that would be removed without mutating the file."""
    try:
        from rdflib import BNode, Graph, Literal, URIRef

        if lang and datatype:
            return {
                "success": False,
                "error": "Cannot specify both --lang and --datatype for a literal",
            }

        graph = Graph()
        try:
            graph.parse(file_path)
        except Exception:
            if not fmt:
                raise
            graph = Graph()
            graph.parse(file_path, format=fmt)

        subj = URIRef(subject) if subject else None
        pred = URIRef(predicate) if predicate else None
        if object_val is None:
            obj = None
        elif object_type == "literal":
            dt = URIRef(datatype) if datatype else None
            obj = Literal(object_val, lang=lang, datatype=dt)
        elif object_type == "bnode":
            obj = BNode(object_val)
        else:
            obj = URIRef(object_val)

        would_remove = sum(1 for _ in graph.triples((subj, pred, obj)))
        return {
            "success": True,
            "file": file_path,
            "dry_run": True,
            "triples_before": len(graph),
            "triples_after": len(graph),
            "would_remove": would_remove,
            "removed": 0,
            "format": fmt,
        }
    except ImportError:
        return {"success": False, "error": "rdflib not installed"}
    except Exception as exc:
        return {"success": False, "error": str(exc)}


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
        fmt = None
        prefixes = {}
        index = 1
        while index < len(raw_args):
            if raw_args[index] == "--base":
                base_uri = _option_value(command, raw_args, index, "--base")
                index += 2
            elif raw_args[index] == "--format":
                fmt = _option_value(command, raw_args, index, "--format")
                index += 2
            elif raw_args[index] == "--prefix":
                parts = _option_value(command, raw_args, index, "--prefix").split("=", 1)
                if len(parts) == 2:
                    prefixes[parts[0]] = parts[1]
                else:
                    raise ValueError(
                        f"{command}: --prefix requires <prefix>=<uri>. "
                        f"Usage: {command_usage(command)}"
                    )
                index += 2
            else:
                _raise_unexpected_argument(command, raw_args[index])
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
        index = 1
        while index < len(raw_args):
            if raw_args[index] == "--limit":
                limit = parse_int(_option_value(command, raw_args, index, "--limit"), "--limit")
                index += 2
            else:
                _raise_unexpected_argument(command, raw_args[index])
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
        fmt = None
        index = 4
        while index < len(raw_args):
            if raw_args[index] == "--type":
                obj_type = _option_value(command, raw_args, index, "--type")
                index += 2
            elif raw_args[index] == "--lang":
                lang = _option_value(command, raw_args, index, "--lang")
                index += 2
            elif raw_args[index] == "--datatype":
                datatype = _option_value(command, raw_args, index, "--datatype")
                index += 2
            elif raw_args[index] == "--format":
                fmt = _option_value(command, raw_args, index, "--format")
                index += 2
            else:
                _raise_unexpected_argument(command, raw_args[index])
        _validate_object_type(command, obj_type)
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
        remove_all = False
        dry_run = False
        index = 1
        while index < len(raw_args):
            if raw_args[index] == "--subject":
                subject = _option_value(command, raw_args, index, "--subject")
                index += 2
            elif raw_args[index] == "--predicate":
                predicate = _option_value(command, raw_args, index, "--predicate")
                index += 2
            elif raw_args[index] == "--object":
                object_val = _option_value(command, raw_args, index, "--object")
                index += 2
            elif raw_args[index] == "--type":
                object_type = _option_value(command, raw_args, index, "--type")
                index += 2
            elif raw_args[index] == "--lang":
                lang = _option_value(command, raw_args, index, "--lang")
                index += 2
            elif raw_args[index] == "--datatype":
                datatype = _option_value(command, raw_args, index, "--datatype")
                index += 2
            elif raw_args[index] == "--format":
                fmt = _option_value(command, raw_args, index, "--format")
                index += 2
            elif raw_args[index] == "--all":
                remove_all = True
                index += 1
            elif raw_args[index] == "--dry-run":
                dry_run = True
                index += 1
            else:
                _raise_unexpected_argument(command, raw_args[index])
        _validate_object_type(command, object_type)
        has_selector = subject is not None or predicate is not None or object_val is not None
        if not remove_all and not has_selector:
            print_usage_error(
                print_result,
                json_output,
                "rdf-remove requires at least one selector or explicit --all.",
                usage=command_usage("rdf-remove"),
            )
            return True
        if remove_all and has_selector:
            print_usage_error(
                print_result,
                json_output,
                "rdf-remove --all cannot be combined with selectors.",
                usage=command_usage("rdf-remove"),
            )
            return True
        if dry_run:
            print_result(
                _rdf_remove_dry_run(
                    raw_args[0],
                    subject=subject,
                    predicate=predicate,
                    object_val=object_val,
                    object_type=object_type,
                    lang=lang,
                    datatype=datatype,
                    fmt=fmt,
                ),
                json_output,
            )
            return True
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
                remove_all=remove_all,
                dry_run=False,
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
            if raw_args[index] == "--limit":
                limit = parse_int(_option_value(command, raw_args, index, "--limit"), "--limit")
                index += 2
            else:
                _raise_unexpected_argument(command, raw_args[index])
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
        fmt = None
        index = 2
        while index < len(raw_args):
            if raw_args[index] == "--format":
                fmt = _option_value(command, raw_args, index, "--format")
                index += 2
            else:
                _raise_unexpected_argument(command, raw_args[index])
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
        fmt = None
        index = 2
        while index < len(raw_args):
            if raw_args[index] == "--output":
                output = _option_value(command, raw_args, index, "--output")
                index += 2
            elif raw_args[index] == "--format":
                fmt = _option_value(command, raw_args, index, "--format")
                index += 2
            else:
                _raise_unexpected_argument(command, raw_args[index])
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
        fmt = None
        if len(raw_args) >= 3 and not raw_args[1].startswith("--"):
            prefix = raw_args[1]
            uri = raw_args[2]
            index = 3
        elif len(raw_args) >= 2 and not raw_args[1].startswith("--"):
            raise ValueError(
                f"{command}: prefix and uri must be provided together. "
                f"Usage: {command_usage(command)}"
            )
        else:
            index = 1
        while index < len(raw_args):
            if raw_args[index] == "--format":
                fmt = _option_value(command, raw_args, index, "--format")
                index += 2
            else:
                _raise_unexpected_argument(command, raw_args[index])
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
        if len(raw_args) > 2:
            _raise_unexpected_argument(command, raw_args[2])
        print_result(doc_server.rdf_validate(raw_args[0], raw_args[1]), json_output)
        return True

    return False
