#!/usr/bin/env python3
"""
CLI-Anything OnlyOffice v4.4.11 - FULL OFFICE SUITE + RDF CONTROL
Programmatic control over Documents (.docx), Spreadsheets (.xlsx),
Presentations (.pptx), and RDF Knowledge Graphs.

Usage:
    cli-anything-onlyoffice <command> [options]
"""

import argparse
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


def dispatch_prefixed_command(command, raw_args, json_output):
    """Dispatch prefix-based modality commands through a small route table."""
    for route in PREFIX_COMMAND_ROUTES:
        if route["matches"](command):
            handled = route["handler"](
                command,
                raw_args,
                doc_server,
                json_output,
                print_result,
                **route["extra_kwargs"](),
            )
            if not handled:
                print_result(
                    {
                        "success": False,
                        "error": f"Unknown {route['error_label']} command: {command}",
                    },
                    json_output,
                )
            return True
    return False

# ==================== MAIN ====================


def main():
    parser = argparse.ArgumentParser(
        description=f"CLI-Anything OnlyOffice v{VERSION} - Documents, Spreadsheets, Presentations, PDFs, RDF", add_help=False
    )
    parser.add_argument("command", nargs="?", default="help", help="Command")
    parser.add_argument("args", nargs="*", default=[], help="Arguments")
    parser.add_argument("--json", action="store_true", help="JSON output")

    # Parse known global flags only; keep raw command args order from sys.argv.
    args, _unknown = parser.parse_known_args()
    args.args = [
        a for a in (sys.argv[2:] if len(sys.argv) > 2 else []) if a != "--json"
    ]
    json_output = args.json
    args.command, alias_meta = normalize_command_alias(args.command)

    if dispatch_prefixed_command(args.command, args.args, json_output):
        return

    if handle_general_command(
        args.command,
        args.args,
        json_output=json_output,
        alias_meta=alias_meta,
        print_result=print_result,
        doc_server=doc_server,
        docx_available=DOCX_AVAILABLE,
        openpyxl_available=OPENPYXL_AVAILABLE,
        pptx_available=PPTX_AVAILABLE,
    ):
        return

    print_result(
        {"success": False, "error": f"Unknown command: {args.command}"}, json_output
    )


if __name__ == "__main__":
    main()
