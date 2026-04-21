#!/usr/bin/env python3
"""General non-modality CLI helpers and command handlers."""

from __future__ import annotations

import glob
import json
import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

from cli_anything.onlyoffice.core.command_registry import (
    TOTAL_COMMANDS,
    VERSION,
    build_help_payload,
    command_usage,
)


OPEN_COMPATIBILITY_NAMESPACES = {
    "document",
    "doc",
    "spreadsheet",
    "sheet",
    "workbook",
    "presentation",
    "slides",
    "slide",
    "pdf",
    "text",
    "file",
    "office",
}


def detect_onlyoffice_file_type(file_path: str) -> str:
    """Best-effort type detection for files OnlyOffice commonly opens."""
    ext = Path(file_path).suffix.lower()
    if ext in {".xlsx", ".xls", ".xlsm", ".xltx", ".xltm", ".ods", ".csv", ".tsv"}:
        return "spreadsheet"
    if ext in {".pptx", ".ppt", ".pptm", ".ppsx", ".odp"}:
        return "presentation"
    if ext == ".pdf":
        return "pdf"
    if ext in {
        ".docx",
        ".doc",
        ".docm",
        ".dotx",
        ".odt",
        ".rtf",
        ".txt",
        ".md",
        ".html",
        ".htm",
        ".epub",
    }:
        return "document"
    return "file"


def normalize_command_alias(command: str) -> Tuple[str, Optional[Dict[str, str]]]:
    """Map agent-style dotted compatibility commands onto the canonical CLI surface."""
    if "." not in command:
        return command, None
    namespace, action = command.rsplit(".", 1)
    if action not in {"open", "watch", "info"}:
        return command, None
    if namespace not in OPEN_COMPATIBILITY_NAMESPACES:
        return command, None
    return action, {
        "requested_command": command,
        "resolved_command": action,
        "command_namespace": namespace,
    }


def apply_alias_meta(
    result: Dict[str, object], alias_meta: Optional[Dict[str, str]]
) -> Dict[str, object]:
    """Attach compatibility alias metadata to a result payload."""
    if alias_meta:
        merged = dict(result)
        merged.update(alias_meta)
        return merged
    return result


def cmd_list(
    *,
    json_output: bool,
    print_result: Callable[[dict, bool], None],
    docx_available: bool,
    openpyxl_available: bool,
) -> None:
    """List recent documents, spreadsheets, and presentations."""
    files = []
    patterns = [
        "~/Documents/*.docx",
        "~/Documents/*.xlsx",
        "~/Documents/*.pptx",
        "~/Documents/*.txt",
        "~/Downloads/*.docx",
        "~/Downloads/*.xlsx",
        "~/Downloads/*.pptx",
        "~/Downloads/*.csv",
    ]

    for pattern in patterns:
        for filepath in glob.glob(os.path.expanduser(pattern)):
            try:
                stat = os.stat(filepath)
                ext = Path(filepath).suffix.lower()
                if ext in [".xlsx", ".csv"]:
                    ftype = "spreadsheet"
                elif ext == ".pptx":
                    ftype = "presentation"
                else:
                    ftype = "document"
                files.append(
                    {
                        "path": filepath,
                        "name": os.path.basename(filepath),
                        "type": ftype,
                        "modified": datetime.fromtimestamp(stat.st_mtime).isoformat(),
                        "size": stat.st_size,
                    }
                )
            except OSError:
                continue

    files.sort(key=lambda item: item["modified"], reverse=True)
    print_result(
        {
            "success": True,
            "count": len(files),
            "files": files[:20],
            "python_docx": docx_available,
            "openpyxl": openpyxl_available,
        },
        json_output,
    )


def cmd_open(
    file_path: str,
    mode: str = "gui",
    *,
    json_output: bool,
    alias_meta: Optional[Dict[str, str]],
    print_result: Callable[[dict, bool], None],
) -> None:
    """Open a document in OnlyOffice GUI or web viewer."""
    try:
        abs_path = os.path.abspath(file_path)
        if not os.path.exists(abs_path):
            print_result(
                apply_alias_meta(
                    {"success": False, "error": f"File not found: {file_path}"},
                    alias_meta,
                ),
                json_output,
            )
            return

        detected_type = detect_onlyoffice_file_type(abs_path)
        if mode == "gui":
            subprocess.Popen(
                ["onlyoffice-desktopeditors", abs_path], start_new_session=True
            )
            result = {
                "success": True,
                "file": abs_path,
                "mode": "gui",
                "detected_type": detected_type,
                "message": "Opened in OnlyOffice Desktop Editors",
            }
        elif mode == "web":
            result = {
                "success": True,
                "file": abs_path,
                "mode": "web",
                "detected_type": detected_type,
                "message": "Document Server URL: http://localhost:8080 (configure shared folder for web access)",
                "note": "For web viewing, copy file to Document Server documents folder",
            }
        else:
            result = {
                "success": False,
                "error": f"Unknown mode: {mode}. Use 'gui' or 'web'",
            }

        print_result(apply_alias_meta(result, alias_meta), json_output)
    except Exception as exc:
        print_result(
            apply_alias_meta({"success": False, "error": str(exc)}, alias_meta),
            json_output,
        )


def cmd_watch(
    file_path: str,
    mode: str = "gui",
    *,
    json_output: bool,
    alias_meta: Optional[Dict[str, str]],
    print_result: Callable[[dict, bool], None],
) -> None:
    """Watch a file and auto-open it in GUI for real-time viewing."""
    try:
        import time

        abs_path = os.path.abspath(file_path)
        if not os.path.exists(abs_path):
            print_result(
                apply_alias_meta(
                    {"success": False, "error": f"File not found: {file_path}"},
                    alias_meta,
                ),
                json_output,
            )
            return

        start_payload = apply_alias_meta(
            {
                "success": True,
                "watching": abs_path,
                "mode": mode,
                "detected_type": detect_onlyoffice_file_type(abs_path),
                "message": f"Started watching {abs_path}. Press Ctrl+C to stop.",
            },
            alias_meta,
        )

        if json_output:
            print(json.dumps(start_payload, indent=2))
        else:
            print(f"Watching: {abs_path}")
            print(f"Mode: {mode}")
            print("Press Ctrl+C to stop")

        if mode == "gui":
            subprocess.Popen(
                ["onlyoffice-desktopeditors", abs_path], start_new_session=True
            )

        last_mtime = os.stat(abs_path).st_mtime
        try:
            while True:
                time.sleep(1)
                if os.path.exists(abs_path):
                    current_mtime = os.stat(abs_path).st_mtime
                    if current_mtime != last_mtime:
                        last_mtime = current_mtime
                        if not json_output:
                            print(
                                f"  [Updated] File modified at {datetime.now().strftime('%H:%M:%S')}"
                            )
        except KeyboardInterrupt:
            if not json_output:
                print("\nWatch stopped.")
            print_result(
                apply_alias_meta(
                    {
                        "success": True,
                        "watching": abs_path,
                        "stopped": True,
                        "detected_type": detect_onlyoffice_file_type(abs_path),
                    },
                    alias_meta,
                ),
                json_output,
            )
    except Exception as exc:
        print_result(
            apply_alias_meta({"success": False, "error": str(exc)}, alias_meta),
            json_output,
        )


def cmd_info(
    file_path: str,
    *,
    json_output: bool,
    alias_meta: Optional[Dict[str, str]],
    print_result: Callable[[dict, bool], None],
    doc_server,
) -> None:
    """Get file information."""
    result = (
        doc_server.get_document_info(file_path)
        if doc_server
        else {"success": False, "error": "Client not available"}
    )
    print_result(apply_alias_meta(result, alias_meta), json_output)


def cmd_editor_session(
    file_path: str,
    *,
    open_if_needed: bool,
    wait_seconds: float,
    activate: bool,
    json_output: bool,
    print_result: Callable[[dict, bool], None],
    doc_server,
) -> None:
    """Locate or open a native OnlyOffice desktop editor session."""
    result = (
        doc_server.editor_session(
            file_path,
            open_if_needed=open_if_needed,
            wait_seconds=wait_seconds,
            activate=activate,
        )
        if doc_server
        else {"success": False, "error": "Client not available"}
    )
    print_result(result, json_output)


def cmd_editor_capture(
    file_path: str,
    output_path: str,
    *,
    backend: str,
    open_if_needed: bool,
    page: Optional[int],
    cell_range: Optional[str],
    slide: Optional[int],
    zoom_reset: bool,
    zoom_in_steps: int,
    zoom_out_steps: int,
    crop: Optional[str],
    settle_ms: int,
    wait_seconds: float,
    dpi: int,
    fmt: Optional[str],
    json_output: bool,
    print_result: Callable[[dict, bool], None],
    doc_server,
) -> None:
    """Capture a live editor viewport or rendered fallback image."""
    result = (
        doc_server.capture_editor_view(
            file_path,
            output_path,
            backend=backend,
            open_if_needed=open_if_needed,
            page=page,
            cell_range=cell_range,
            slide=slide,
            zoom_reset=zoom_reset,
            zoom_in_steps=zoom_in_steps,
            zoom_out_steps=zoom_out_steps,
            crop=crop,
            settle_ms=settle_ms,
            wait_seconds=wait_seconds,
            dpi=dpi,
            fmt=fmt,
        )
        if doc_server
        else {"success": False, "error": "Client not available"}
    )
    print_result(result, json_output)


def cmd_backup_list(
    file_path: str,
    *,
    limit: int,
    json_output: bool,
    print_result: Callable[[dict, bool], None],
    doc_server,
) -> None:
    """List backups for a target file."""
    if not doc_server:
        print_result({"success": False, "error": "Client not available"}, json_output)
        return
    result = doc_server.list_backups(file_path, limit=int(limit))
    print_result(result, json_output)


def cmd_backup_prune(
    *,
    file_path: Optional[str],
    keep: int,
    days: Optional[int],
    json_output: bool,
    print_result: Callable[[dict, bool], None],
    doc_server,
) -> None:
    """Prune backups by retention rules."""
    if not doc_server:
        print_result({"success": False, "error": "Client not available"}, json_output)
        return
    result = doc_server.prune_backups(
        file_path=file_path,
        keep=int(keep),
        older_than_days=int(days) if days is not None else None,
    )
    print_result(result, json_output)


def cmd_backup_restore(
    file_path: str,
    *,
    backup: Optional[str],
    latest: bool,
    dry_run: bool,
    json_output: bool,
    print_result: Callable[[dict, bool], None],
    doc_server,
) -> None:
    """Restore a file from backup."""
    if not doc_server:
        print_result({"success": False, "error": "Client not available"}, json_output)
        return
    result = doc_server.restore_backup(
        file_path=file_path,
        backup=backup,
        latest=latest,
        dry_run=dry_run,
    )
    print_result(result, json_output)


def cmd_status(
    *,
    json_output: bool,
    print_result: Callable[[dict, bool], None],
    doc_server,
    docx_available: bool,
    openpyxl_available: bool,
    pptx_available: bool,
) -> None:
    """Check installation status."""
    try:
        import rdflib

        rdflib_available = True
        rdflib_version = rdflib.__version__
    except ImportError:
        rdflib_available = False
        rdflib_version = None

    try:
        import pyshacl

        shacl_available = True
    except ImportError:
        shacl_available = False

    capabilities = {
        "docx_create": docx_available,
        "docx_read": docx_available,
        "docx_edit": docx_available,
        "docx_tables": docx_available,
        "docx_formatting": docx_available,
        "docx_references": docx_available,
        "xlsx_create": openpyxl_available,
        "xlsx_read": openpyxl_available,
        "xlsx_edit": openpyxl_available,
        "xlsx_formulas": openpyxl_available,
        "xlsx_charts": openpyxl_available,
        "xlsx_stats": openpyxl_available,
        "xlsx_csv": openpyxl_available,
        "pptx_create": pptx_available,
        "pptx_read": pptx_available,
        "pptx_edit": pptx_available,
        "pptx_notes": pptx_available,
        "rdf_create": rdflib_available,
        "rdf_query": rdflib_available,
        "rdf_validate": shacl_available,
    }
    print_result(
        {
            "success": True,
            "version": VERSION,
            "python": sys.executable,
            "document_server": {
                "healthy": doc_server.check_health() if doc_server else False
            },
            "python_docx": docx_available,
            "openpyxl": openpyxl_available,
            "python_pptx": pptx_available,
            "rdflib": rdflib_available,
            "rdflib_version": rdflib_version,
            "pyshacl": shacl_available,
            "capabilities": capabilities,
            "total_commands": TOTAL_COMMANDS,
        },
        json_output,
    )


def cmd_help(
    *,
    json_output: bool,
    print_result: Callable[[dict, bool], None],
    docx_available: bool,
    openpyxl_available: bool,
    pptx_available: bool,
) -> None:
    """Show help."""
    try:
        import rdflib

        rdflib_available = True
    except ImportError:
        rdflib_available = False

    result = build_help_payload(
        {
            "python_docx": docx_available,
            "openpyxl": openpyxl_available,
            "python_pptx": pptx_available,
            "rdflib": rdflib_available,
        }
    )
    print_result(result, json_output)


def handle_general_command(
    command: str,
    raw_args: List[str],
    *,
    json_output: bool,
    alias_meta: Optional[Dict[str, str]],
    print_result: Callable[[dict, bool], None],
    doc_server,
    docx_available: bool,
    openpyxl_available: bool,
    pptx_available: bool,
) -> bool:
    """Handle general non-prefixed CLI commands and return True when recognised."""
    if command == "list":
        cmd_list(
            json_output=json_output,
            print_result=print_result,
            docx_available=docx_available,
            openpyxl_available=openpyxl_available,
        )
        return True

    if command == "open":
        if not raw_args:
            print_result({"success": False, "error": command_usage("open")}, json_output)
        else:
            mode = raw_args[1] if len(raw_args) > 1 else "gui"
            cmd_open(
                raw_args[0],
                mode,
                json_output=json_output,
                alias_meta=alias_meta,
                print_result=print_result,
            )
        return True

    if command == "watch":
        if not raw_args:
            print_result({"success": False, "error": command_usage("watch")}, json_output)
        else:
            mode = raw_args[1] if len(raw_args) > 1 else "gui"
            cmd_watch(
                raw_args[0],
                mode,
                json_output=json_output,
                alias_meta=alias_meta,
                print_result=print_result,
            )
        return True

    if command == "info":
        if not raw_args:
            print_result({"success": False, "error": command_usage("info")}, json_output)
        else:
            cmd_info(
                raw_args[0],
                json_output=json_output,
                alias_meta=alias_meta,
                print_result=print_result,
                doc_server=doc_server,
            )
        return True

    if command == "editor-session":
        if not raw_args:
            print_result(
                {"success": False, "error": command_usage("editor-session")},
                json_output,
            )
        else:
            open_if_needed = False
            wait_seconds = 10.0
            activate = False
            i = 1
            while i < len(raw_args):
                if raw_args[i] == "--open":
                    open_if_needed = True
                    i += 1
                elif raw_args[i] == "--wait" and i + 1 < len(raw_args):
                    wait_seconds = float(raw_args[i + 1])
                    i += 2
                elif raw_args[i] == "--activate":
                    activate = True
                    i += 1
                else:
                    i += 1
            cmd_editor_session(
                raw_args[0],
                open_if_needed=open_if_needed,
                wait_seconds=wait_seconds,
                activate=activate,
                json_output=json_output,
                print_result=print_result,
                doc_server=doc_server,
            )
        return True

    if command == "editor-capture":
        if len(raw_args) < 2:
            print_result(
                {"success": False, "error": command_usage("editor-capture")},
                json_output,
            )
        else:
            backend = "auto"
            open_if_needed = False
            page = None
            cell_range = None
            slide = None
            zoom_reset = False
            zoom_in_steps = 0
            zoom_out_steps = 0
            crop = None
            settle_ms = 800
            wait_seconds = 10.0
            dpi = 150
            fmt = None
            i = 2
            while i < len(raw_args):
                if raw_args[i] == "--backend" and i + 1 < len(raw_args):
                    backend = raw_args[i + 1]
                    i += 2
                elif raw_args[i] == "--open":
                    open_if_needed = True
                    i += 1
                elif raw_args[i] == "--page" and i + 1 < len(raw_args):
                    page = int(raw_args[i + 1])
                    i += 2
                elif raw_args[i] == "--range" and i + 1 < len(raw_args):
                    cell_range = raw_args[i + 1]
                    i += 2
                elif raw_args[i] == "--slide" and i + 1 < len(raw_args):
                    slide = int(raw_args[i + 1])
                    i += 2
                elif raw_args[i] == "--zoom-reset":
                    zoom_reset = True
                    i += 1
                elif raw_args[i] == "--zoom-in" and i + 1 < len(raw_args):
                    zoom_in_steps = int(raw_args[i + 1])
                    i += 2
                elif raw_args[i] == "--zoom-out" and i + 1 < len(raw_args):
                    zoom_out_steps = int(raw_args[i + 1])
                    i += 2
                elif raw_args[i] == "--crop" and i + 1 < len(raw_args):
                    crop = raw_args[i + 1]
                    i += 2
                elif raw_args[i] == "--settle-ms" and i + 1 < len(raw_args):
                    settle_ms = int(raw_args[i + 1])
                    i += 2
                elif raw_args[i] == "--wait" and i + 1 < len(raw_args):
                    wait_seconds = float(raw_args[i + 1])
                    i += 2
                elif raw_args[i] == "--dpi" and i + 1 < len(raw_args):
                    dpi = int(raw_args[i + 1])
                    i += 2
                elif raw_args[i] == "--format" and i + 1 < len(raw_args):
                    fmt = raw_args[i + 1]
                    i += 2
                else:
                    i += 1
            cmd_editor_capture(
                raw_args[0],
                raw_args[1],
                backend=backend,
                open_if_needed=open_if_needed,
                page=page,
                cell_range=cell_range,
                slide=slide,
                zoom_reset=zoom_reset,
                zoom_in_steps=zoom_in_steps,
                zoom_out_steps=zoom_out_steps,
                crop=crop,
                settle_ms=settle_ms,
                wait_seconds=wait_seconds,
                dpi=dpi,
                fmt=fmt,
                json_output=json_output,
                print_result=print_result,
                doc_server=doc_server,
            )
        return True

    if command == "backup-list":
        if not raw_args:
            print_result(
                {"success": False, "error": command_usage("backup-list")},
                json_output,
            )
        else:
            file_path = None
            limit = 20
            i = 0
            while i < len(raw_args):
                if raw_args[i] == "--limit" and i + 1 < len(raw_args):
                    limit = int(raw_args[i + 1])
                    i += 2
                elif not raw_args[i].startswith("--") and file_path is None:
                    file_path = raw_args[i]
                    i += 1
                else:
                    i += 1
            if not file_path:
                print_result(
                    {"success": False, "error": command_usage("backup-list")},
                    json_output,
                )
            else:
                cmd_backup_list(
                    file_path,
                    limit=limit,
                    json_output=json_output,
                    print_result=print_result,
                    doc_server=doc_server,
                )
        return True

    if command == "backup-prune":
        file_path = None
        keep = 20
        days = None
        i = 0
        while i < len(raw_args):
            if raw_args[i] == "--file" and i + 1 < len(raw_args):
                file_path = raw_args[i + 1]
                i += 2
            elif raw_args[i] == "--keep" and i + 1 < len(raw_args):
                keep = int(raw_args[i + 1])
                i += 2
            elif raw_args[i] == "--days" and i + 1 < len(raw_args):
                days = int(raw_args[i + 1])
                i += 2
            elif not raw_args[i].startswith("--") and file_path is None:
                file_path = raw_args[i]
                i += 1
            else:
                i += 1
        cmd_backup_prune(
            file_path=file_path,
            keep=keep,
            days=days,
            json_output=json_output,
            print_result=print_result,
            doc_server=doc_server,
        )
        return True

    if command == "backup-restore":
        if not raw_args:
            print_result(
                {"success": False, "error": command_usage("backup-restore")},
                json_output,
            )
        else:
            file_path = raw_args[0]
            backup = None
            latest = False
            dry_run = False
            i = 1
            while i < len(raw_args):
                if raw_args[i] == "--backup" and i + 1 < len(raw_args):
                    backup = raw_args[i + 1]
                    i += 2
                elif raw_args[i] == "--latest":
                    latest = True
                    i += 1
                elif raw_args[i] == "--dry-run":
                    dry_run = True
                    i += 1
                else:
                    i += 1
            cmd_backup_restore(
                file_path,
                backup=backup,
                latest=latest,
                dry_run=dry_run,
                json_output=json_output,
                print_result=print_result,
                doc_server=doc_server,
            )
        return True

    if command == "status":
        cmd_status(
            json_output=json_output,
            print_result=print_result,
            doc_server=doc_server,
            docx_available=docx_available,
            openpyxl_available=openpyxl_available,
            pptx_available=pptx_available,
        )
        return True

    if command in {"help", "--help", "-h"}:
        cmd_help(
            json_output=json_output,
            print_result=print_result,
            docx_available=docx_available,
            openpyxl_available=openpyxl_available,
            pptx_available=pptx_available,
        )
        return True

    return False
