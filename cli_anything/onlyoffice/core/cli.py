#!/usr/bin/env python3
"""
CLI-Anything OnlyOffice v4.4.16 - FULL OFFICE SUITE + RDF CONTROL
Programmatic control over Documents (.docx), Spreadsheets (.xlsx),
Presentations (.pptx), and RDF Knowledge Graphs.

Usage:
    cli-anything-onlyoffice <command> [options]
"""

import json
import sys

from cli_anything.onlyoffice.core.command_registry import (
    VERSION,
)
from cli_anything.onlyoffice.core.doc_cli import handle_doc_command
from cli_anything.onlyoffice.core.general_cli import (
    handle_general_command,
    normalize_command_alias,
)
from cli_anything.onlyoffice.core.pdf_cli import handle_pdf_command
from cli_anything.onlyoffice.core.pptx_cli import handle_pptx_command
from cli_anything.onlyoffice.core.rdf_cli import handle_rdf_command
from cli_anything.onlyoffice.core.xlsx_cli import handle_xlsx_command
from cli_anything.onlyoffice.core.parse_utils import print_usage_error

# Import enhanced Document Server client
try:
    from cli_anything.onlyoffice.utils.docserver import (
        get_client,
        DOCX_AVAILABLE,
        OPENPYXL_AVAILABLE,
        PPTX_AVAILABLE,
    )

    CLIENT_AVAILABLE = True
except ImportError:
    CLIENT_AVAILABLE = False
    DOCX_AVAILABLE = False
    OPENPYXL_AVAILABLE = False
    PPTX_AVAILABLE = False


# Initialize client
if CLIENT_AVAILABLE:
    doc_server = get_client()
else:
    doc_server = None


def print_result(result, json_output=False):
    """Print result in appropriate format"""
    if json_output:
        print(json.dumps(result, indent=2, default=str))
    else:
        if result.get("success"):
            if "files" in result or "documents" in result:
                print(f"Found {result['count']} files:")
                for doc in result.get("files", result.get("documents", [])):
                    print(f"  - {doc['name']} ({doc['modified'][:10]})")
            elif "paragraphs" in result:
                print(f"Document: {result.get('file', 'unknown')}")
                print(f"Paragraphs: {result.get('paragraph_count', 0)}")
                print(f"\nContent:\n{result.get('full_text', '')[:500]}")
            elif "data" in result and "headers" in result["data"].get(
                list(result["data"].keys())[0], {}
            ):
                print(f"Spreadsheet: {result.get('file', 'unknown')}")
                sheet = list(result["data"].keys())[0]
                print(f"Sheet: {sheet}")
                print(f"Headers: {result['data'][sheet].get('headers', [])}")
                print(f"Rows: {result['data'][sheet].get('row_count', 0)}")
            else:
                print("Success!")
                for key, value in result.items():
                    if key not in ["success"]:
                        print(f"  {key}: {value}")
        else:
            print(f"Error: {result.get('error', 'Unknown error')}")
    return 1 if result.get("success") is False else 0


class ResultTracker:
    """Record the process status implied by printed CLI payloads."""

    def __init__(self):
        self.exit_code = 0

    def print_result(self, result, json_output=False):
        exit_code = print_result(result, json_output)
        if exit_code:
            self.exit_code = exit_code
        return exit_code


PREFIX_COMMAND_ROUTES = (
    {
        "matches": lambda command: command.startswith("doc-"),
        "handler": handle_doc_command,
        "error_label": "document",
        "extra_kwargs": lambda: {"docx_available": DOCX_AVAILABLE},
    },
    {
        "matches": lambda command: command.startswith("xlsx-") or command.startswith("chart-"),
        "handler": handle_xlsx_command,
        "error_label": "spreadsheet/chart",
        "extra_kwargs": lambda: {},
    },
    {
        "matches": lambda command: command.startswith("pptx-"),
        "handler": handle_pptx_command,
        "error_label": "PPTX",
        "extra_kwargs": lambda: {},
    },
    {
        "matches": lambda command: command.startswith("rdf-"),
        "handler": handle_rdf_command,
        "error_label": "RDF",
        "extra_kwargs": lambda: {},
    },
    {
        "matches": lambda command: command.startswith("pdf-"),
        "handler": handle_pdf_command,
        "error_label": "PDF",
        "extra_kwargs": lambda: {},
    },
)


def dispatch_prefixed_command(command, raw_args, json_output, result_printer=print_result):
    """Dispatch prefix-based modality commands through a small route table."""
    for route in PREFIX_COMMAND_ROUTES:
        if route["matches"](command):
            if doc_server is None:
                result_printer(
                    {
                        "success": False,
                        "error": (
                            f"Cannot run {route['error_label']} command because the "
                            "OnlyOffice client failed to initialise."
                        ),
                        "error_code": "client_unavailable",
                    },
                    json_output,
                )
                return True
            handled = route["handler"](
                command,
                raw_args,
                doc_server,
                json_output,
                result_printer,
                **route["extra_kwargs"](),
            )
            if not handled:
                result_printer(
                    {
                        "success": False,
                        "error": f"Unknown {route['error_label']} command: {command}",
                    },
                    json_output,
                )
            return True
    return False

# ==================== MAIN ====================


class CommandArgs(list):
    """Command argv list with the index where `--` made args positional."""

    def __init__(self, values=(), literal_start=None):
        super().__init__(values)
        self.literal_start = literal_start


def _parse_global_args(argv):
    """Extract global flags without eating literal command arguments.

    Backward compatibility: `--json` is accepted before the command or as the
    final command argument. Use `--` before literal content that happens to end
    with `--json`.
    """
    tokens = list(argv)
    json_output = False
    if tokens and tokens[0] == "--json":
        json_output = True
        tokens = tokens[1:]
    if not tokens:
        return "help", CommandArgs(), json_output

    command = tokens[0]
    raw_args = tokens[1:]
    if raw_args and "--" in raw_args:
        delimiter = raw_args.index("--")
        before = raw_args[:delimiter]
        after = raw_args[delimiter + 1 :]
        if before and before[-1] == "--json":
            json_output = True
            before = before[:-1]
        raw_args = CommandArgs(before + after, literal_start=len(before))
    elif raw_args and raw_args[-1] == "--json":
        json_output = True
        raw_args = CommandArgs(raw_args[:-1])
    else:
        raw_args = CommandArgs(raw_args)
    return command, raw_args, json_output


def main(argv=None, *, exit_on_error=False):
    tracker = ResultTracker()
    result_printer = tracker.print_result
    command, raw_args, json_output = _parse_global_args(sys.argv[1:] if argv is None else argv)
    command, alias_meta = normalize_command_alias(command)

    try:
        if command in {"version", "--version", "-V"}:
            result_printer({"success": True, "version": VERSION}, json_output)
            return tracker.exit_code

        if command.startswith("-") and command not in {"--help", "-h"}:
            print_usage_error(
                result_printer,
                json_output,
                f"Unknown global option: {command}",
            )
            if exit_on_error and tracker.exit_code:
                raise SystemExit(tracker.exit_code)
            return tracker.exit_code

        if dispatch_prefixed_command(command, raw_args, json_output, result_printer):
            if exit_on_error and tracker.exit_code:
                raise SystemExit(tracker.exit_code)
            return tracker.exit_code

        if handle_general_command(
            command,
            raw_args,
            json_output=json_output,
            alias_meta=alias_meta,
            print_result=result_printer,
            doc_server=doc_server,
            docx_available=DOCX_AVAILABLE,
            openpyxl_available=OPENPYXL_AVAILABLE,
            pptx_available=PPTX_AVAILABLE,
        ):
            if exit_on_error and tracker.exit_code:
                raise SystemExit(tracker.exit_code)
            return tracker.exit_code

        result_printer(
            {"success": False, "error": f"Unknown command: {command}"}, json_output
        )
    except ValueError as exc:
        print_usage_error(result_printer, json_output, str(exc))

    if exit_on_error and tracker.exit_code:
        raise SystemExit(tracker.exit_code)
    return tracker.exit_code


if __name__ == "__main__":
    raise SystemExit(main(exit_on_error=True))
