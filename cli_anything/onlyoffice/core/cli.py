#!/usr/bin/env python3
"""
CLI-Anything OnlyOffice v4.0 - FULL OFFICE SUITE + RDF CONTROL
Programmatic control over Documents (.docx), Spreadsheets (.xlsx),
Presentations (.pptx), and RDF Knowledge Graphs.

Usage:
    cli-anything-onlyoffice <command> [options]
"""

import argparse
import json
import subprocess
import sys
import os
import glob
from datetime import datetime
from pathlib import Path

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


# ==================== COMMANDS ====================


def cmd_list(json_output=False):
    """List recent documents, spreadsheets, and presentations"""
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
            except:
                pass

    files.sort(key=lambda x: x["modified"], reverse=True)
    print_result(
        {
            "success": True,
            "count": len(files),
            "files": files[:20],
            "python_docx": DOCX_AVAILABLE,
            "openpyxl": OPENPYXL_AVAILABLE,
        },
        json_output,
    )


def cmd_open(file_path, mode="gui", json_output=False):
    """Open a document in OnlyOffice GUI or web viewer"""
    try:
        abs_path = os.path.abspath(file_path)
        if not os.path.exists(abs_path):
            return print_result(
                {"success": False, "error": f"File not found: {file_path}"}, json_output
            )

        if mode == "gui":
            # Open in OnlyOffice Desktop Editors GUI
            subprocess.Popen(
                ["onlyoffice-desktopeditors", abs_path], start_new_session=True
            )
            result = {
                "success": True,
                "file": abs_path,
                "mode": "gui",
                "message": "Opened in OnlyOffice Desktop Editors",
            }
        elif mode == "web":
            # For web viewing, we'd need Document Server configured
            # This would typically involve setting up a shared folder
            result = {
                "success": True,
                "file": abs_path,
                "mode": "web",
                "message": "Document Server URL: http://localhost:8080 (configure shared folder for web access)",
                "note": "For web viewing, copy file to Document Server documents folder",
            }
        else:
            result = {
                "success": False,
                "error": f"Unknown mode: {mode}. Use 'gui' or 'web'",
            }

        print_result(result, json_output)
    except Exception as e:
        print_result({"success": False, "error": str(e)}, json_output)


def cmd_watch(file_path, mode="gui", json_output=False):
    """Watch a file and auto-open it in GUI for real-time viewing"""
    try:
        import time

        abs_path = os.path.abspath(file_path)
        if not os.path.exists(abs_path):
            return print_result(
                {"success": False, "error": f"File not found: {file_path}"}, json_output
            )

        if json_output:
            print(
                json.dumps(
                    {
                        "success": True,
                        "watching": abs_path,
                        "mode": mode,
                        "message": f"Started watching {abs_path}. Press Ctrl+C to stop.",
                    },
                    indent=2,
                )
            )
        else:
            print(f"Watching: {abs_path}")
            print(f"Mode: {mode}")
            print("Press Ctrl+C to stop")

        # Initial open
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
                        # OnlyOffice Desktop Editors auto-reloads, no need to restart
        except KeyboardInterrupt:
            if not json_output:
                print("\nWatch stopped.")
            print_result(
                {"success": True, "watching": abs_path, "stopped": True}, json_output
            )
    except Exception as e:
        print_result({"success": False, "error": str(e)}, json_output)


def cmd_doc_create(output_path, title, content, json_output=False):
    """Create a new .docx document"""
    if not DOCX_AVAILABLE:
        return print_result(
            {"success": False, "error": "python-docx not installed"}, json_output
        )

    result = doc_server.create_document(output_path, title, content)
    print_result(result, json_output)


def cmd_doc_read(file_path, json_output=False):
    """Read a .docx document and extract all content"""
    if not DOCX_AVAILABLE:
        return print_result(
            {"success": False, "error": "python-docx not installed"}, json_output
        )

    result = doc_server.read_document(file_path)
    print_result(result, json_output)


def cmd_doc_append(file_path, content, json_output=False):
    """Append content to a .docx document"""
    if not DOCX_AVAILABLE:
        return print_result(
            {"success": False, "error": "python-docx not installed"}, json_output
        )

    result = doc_server.append_to_document(file_path, content)
    print_result(result, json_output)


def cmd_doc_replace(file_path, search, replace, json_output=False):
    """Find and replace text in a .docx document"""
    if not DOCX_AVAILABLE:
        return print_result(
            {"success": False, "error": "python-docx not installed"}, json_output
        )

    result = doc_server.search_replace_document(file_path, search, replace)
    print_result(result, json_output)


def cmd_xlsx_create(output_path, sheet_name="Sheet1", json_output=False):
    """Create a new .xlsx spreadsheet"""
    if not OPENPYXL_AVAILABLE:
        return print_result(
            {"success": False, "error": "openpyxl not installed"}, json_output
        )

    result = doc_server.create_spreadsheet(output_path, sheet_name)
    print_result(result, json_output)


def cmd_xlsx_write(
    output_path,
    headers_csv,
    data_csv,
    sheet_name="Sheet1",
    overwrite_workbook=False,
    coerce_rows=False,
    text_columns=None,
    json_output=False,
):
    """Write data to spreadsheet (updates target sheet by default)."""
    if not OPENPYXL_AVAILABLE:
        return print_result(
            {"success": False, "error": "openpyxl not installed"}, json_output
        )

    headers = [h.strip() for h in headers_csv.split(",")]
    rows = []
    for row_str in data_csv.split(";"):
        if row_str.strip():
            rows.append([c.strip() for c in row_str.split(",")])

    result = doc_server.write_spreadsheet(
        output_path,
        headers,
        rows,
        sheet_name=sheet_name,
        overwrite_workbook=overwrite_workbook,
        coerce_rows=coerce_rows,
        text_columns=text_columns,
    )
    print_result(result, json_output)


def cmd_xlsx_read(file_path, sheet_name=None, json_output=False):
    """Read a spreadsheet"""
    if not OPENPYXL_AVAILABLE:
        return print_result(
            {"success": False, "error": "openpyxl not installed"}, json_output
        )

    result = doc_server.read_spreadsheet(file_path, sheet_name)
    print_result(result, json_output)


def cmd_xlsx_append(file_path, row_data_csv, sheet_name=None, json_output=False):
    """Append a row to a spreadsheet"""
    if not OPENPYXL_AVAILABLE:
        return print_result(
            {"success": False, "error": "openpyxl not installed"}, json_output
        )

    row_data = [c.strip() for c in row_data_csv.split(",")]
    result = doc_server.append_to_spreadsheet(
        file_path, row_data, sheet_name=sheet_name
    )
    print_result(result, json_output)


def cmd_xlsx_search(file_path, search_text, sheet_name=None, json_output=False):
    """Search for text in a spreadsheet"""
    if not OPENPYXL_AVAILABLE:
        return print_result(
            {"success": False, "error": "openpyxl not installed"}, json_output
        )

    result = doc_server.search_spreadsheet(
        file_path, search_text, sheet_name=sheet_name
    )
    print_result(result, json_output)


def cmd_xlsx_calc(
    file_path,
    column,
    operation,
    sheet_name="Sheet1",
    include_formulas=False,
    strict_formulas=False,
    json_output=False,
):
    """Calculate column statistics"""
    if not OPENPYXL_AVAILABLE:
        return print_result(
            {"success": False, "error": "openpyxl not installed"}, json_output
        )

    result = doc_server.calculate_column(
        file_path,
        column,
        operation,
        sheet_name=sheet_name,
        include_formulas=include_formulas,
        strict_formula_safety=strict_formulas,
    )
    print_result(result, json_output)


def cmd_xlsx_formula(file_path, cell, formula, sheet_name="Sheet1", json_output=False):
    """Add formula to spreadsheet"""
    if not OPENPYXL_AVAILABLE:
        return print_result(
            {"success": False, "error": "openpyxl not installed"}, json_output
        )

    result = doc_server.add_formula(file_path, cell, formula, sheet_name=sheet_name)
    print_result(result, json_output)


def cmd_xlsx_freq(
    file_path,
    column,
    sheet_name="Sheet1",
    valid_values=None,
    json_output=False,
):
    """Frequency table for one column."""
    if not OPENPYXL_AVAILABLE:
        return print_result(
            {"success": False, "error": "openpyxl not installed"}, json_output
        )
    result = doc_server.frequencies(
        file_path,
        column,
        sheet_name=sheet_name,
        allowed_values=valid_values,
    )
    print_result(result, json_output)


def cmd_xlsx_corr(
    file_path,
    x_column,
    y_column,
    sheet_name="Sheet1",
    method="pearson",
    json_output=False,
):
    """Correlation test between two numeric columns."""
    if not OPENPYXL_AVAILABLE:
        return print_result(
            {"success": False, "error": "openpyxl not installed"}, json_output
        )
    result = doc_server.correlation_test(
        file_path=file_path,
        x_column=x_column,
        y_column=y_column,
        sheet_name=sheet_name,
        method=method,
    )
    print_result(result, json_output)


def cmd_xlsx_ttest(
    file_path,
    value_column,
    group_column,
    group_a,
    group_b,
    sheet_name="Sheet1",
    equal_var=False,
    json_output=False,
):
    """Independent samples t-test."""
    if not OPENPYXL_AVAILABLE:
        return print_result(
            {"success": False, "error": "openpyxl not installed"}, json_output
        )
    result = doc_server.ttest_independent(
        file_path=file_path,
        value_column=value_column,
        group_column=group_column,
        group_a=group_a,
        group_b=group_b,
        sheet_name=sheet_name,
        equal_var=equal_var,
    )
    print_result(result, json_output)


def cmd_xlsx_chi2(
    file_path,
    row_column,
    col_column,
    sheet_name="Sheet1",
    row_valid_values=None,
    col_valid_values=None,
    json_output=False,
):
    """Chi-square test of independence for two categorical columns."""
    if not OPENPYXL_AVAILABLE:
        return print_result(
            {"success": False, "error": "openpyxl not installed"}, json_output
        )
    result = doc_server.chi_square_test(
        file_path=file_path,
        row_column=row_column,
        col_column=col_column,
        sheet_name=sheet_name,
        row_allowed_values=row_valid_values,
        col_allowed_values=col_valid_values,
    )
    print_result(result, json_output)


def cmd_xlsx_research_pack(
    file_path,
    sheet_name="Sheet0",
    profile="hlth3112",
    require_formula_safe=False,
    json_output=False,
):
    """Run standardized research analysis bundle."""
    if not OPENPYXL_AVAILABLE:
        return print_result(
            {"success": False, "error": "openpyxl not installed"}, json_output
        )
    result = doc_server.research_analysis_pack(
        file_path=file_path,
        sheet_name=sheet_name,
        profile=profile,
        require_formula_safe=require_formula_safe,
    )
    print_result(result, json_output)


def cmd_xlsx_formula_audit(file_path, sheet_name=None, json_output=False):
    """Audit workbook formula risk for production execution safety."""
    if not OPENPYXL_AVAILABLE:
        return print_result(
            {"success": False, "error": "openpyxl not installed"}, json_output
        )
    result = doc_server.audit_spreadsheet_formulas(
        file_path=file_path,
        sheet_name=sheet_name,
    )
    print_result(result, json_output)


def cmd_xlsx_text_extract(
    file_path,
    column,
    sheet_name="Sheet0",
    limit=20,
    min_length=20,
    json_output=False,
):
    """Extract open-text responses for qualitative coding."""
    if not OPENPYXL_AVAILABLE:
        return print_result(
            {"success": False, "error": "openpyxl not installed"}, json_output
        )
    result = doc_server.open_text_extract(
        file_path=file_path,
        column_letter=column,
        sheet_name=sheet_name,
        limit=int(limit),
        min_length=int(min_length),
    )
    print_result(result, json_output)


def cmd_xlsx_text_keywords(
    file_path,
    column,
    sheet_name="Sheet0",
    top=15,
    min_word_length=4,
    json_output=False,
):
    """Generate keyword summary for open-text responses."""
    if not OPENPYXL_AVAILABLE:
        return print_result(
            {"success": False, "error": "openpyxl not installed"}, json_output
        )
    result = doc_server.open_text_keywords(
        file_path=file_path,
        column_letter=column,
        sheet_name=sheet_name,
        top_n=int(top),
        min_word_length=int(min_word_length),
    )
    print_result(result, json_output)


def cmd_doc_format(
    file_path,
    paragraph_index,
    bold=False,
    italic=False,
    underline=False,
    font_name=None,
    font_size=None,
    color=None,
    alignment=None,
    json_output=False,
):
    """Apply formatting to a paragraph in a document."""
    if not DOCX_AVAILABLE:
        return print_result(
            {"success": False, "error": "python-docx not installed"}, json_output
        )
    result = doc_server.format_paragraph(
        file_path=file_path,
        paragraph_index=int(paragraph_index),
        bold=bold,
        italic=italic,
        underline=underline,
        font_name=font_name,
        font_size=int(font_size) if font_size else None,
        color=color,
        alignment=alignment,
    )
    print_result(result, json_output)


def cmd_doc_highlight(file_path, search_text, color="yellow", json_output=False):
    """Highlight text in a document."""
    if not DOCX_AVAILABLE:
        return print_result(
            {"success": False, "error": "python-docx not installed"}, json_output
        )
    result = doc_server.highlight_text(file_path, search_text, color=color)
    print_result(result, json_output)


def cmd_doc_comment(file_path, comment_text, paragraph_index=0, json_output=False):
    """Attach a metadata comment to a document."""
    if not DOCX_AVAILABLE:
        return print_result(
            {"success": False, "error": "python-docx not installed"}, json_output
        )
    result = doc_server.add_comment(
        file_path, comment_text, paragraph_index=int(paragraph_index)
    )
    print_result(result, json_output)


def cmd_doc_layout(file_path, orientation="portrait", margins=None, json_output=False):
    """Set page layout and optional margins."""
    if not DOCX_AVAILABLE:
        return print_result(
            {"success": False, "error": "python-docx not installed"}, json_output
        )
    result = doc_server.set_page_layout(
        file_path, orientation=orientation, margins=margins or {}
    )
    print_result(result, json_output)


def cmd_doc_formatting_info(file_path, json_output=False):
    """Get formatting diagnostics for document."""
    if not DOCX_AVAILABLE:
        return print_result(
            {"success": False, "error": "python-docx not installed"}, json_output
        )
    result = doc_server.get_formatting_info(file_path)
    print_result(result, json_output)


# ==================== CHART COMMANDS ====================


def cmd_chart_create(
    file_path,
    chart_type,
    data_range,
    categories_range,
    title,
    sheet_name,
    output_sheet,
    x_label,
    y_label,
    show_labels,
    show_legend,
    legend_pos,
    colors,
    json_output=False,
):
    """Create a chart in a spreadsheet"""
    if not OPENPYXL_AVAILABLE:
        return print_result(
            {"success": False, "error": "openpyxl not installed"}, json_output
        )

    # Parse colors if provided
    color_list = None
    if colors:
        color_list = colors.split(",")

    result = doc_server.create_chart(
        file_path=file_path,
        chart_type=chart_type,
        data_range=data_range,
        categories_range=categories_range,
        title=title,
        sheet_name=sheet_name,
        output_sheet=output_sheet,
        x_label=x_label,
        y_label=y_label,
        show_data_labels=show_labels,
        show_legend=show_legend,
        legend_pos=legend_pos,
        colors=color_list,
    )
    print_result(result, json_output)


def cmd_chart_comparison(
    file_path,
    chart_type,
    sheet_name,
    title,
    start_row,
    start_col,
    num_cats,
    num_series,
    cat_col,
    value_cols,
    output_cell,
    show_labels,
    show_legend,
    json_output=False,
):
    """Create a comparison chart from structured data"""
    if not OPENPYXL_AVAILABLE:
        return print_result(
            {"success": False, "error": "openpyxl not installed"}, json_output
        )

    # Parse value columns
    cols_list = None
    if value_cols:
        cols_list = [int(c) for c in value_cols.split(",")]

    result = doc_server.create_comparison_chart(
        file_path=file_path,
        chart_type=chart_type,
        sheet_name=sheet_name,
        title=title,
        start_row=int(start_row) if start_row else 1,
        start_col=int(start_col) if start_col else 1,
        num_categories=int(num_cats) if num_cats else None,
        num_series=int(num_series) if num_series else None,
        category_col=int(cat_col) if cat_col else 1,
        value_cols=cols_list,
        output_cell=output_cell,
        show_data_labels=show_labels,
        show_legend=show_legend,
    )
    print_result(result, json_output)


def cmd_chart_grade_dist(
    file_path, sheet_name, grade_col, title, output_cell, json_output=False
):
    """Create a pie chart showing grade distribution"""
    if not OPENPYXL_AVAILABLE:
        return print_result(
            {"success": False, "error": "openpyxl not installed"}, json_output
        )

    result = doc_server.create_grade_distribution_chart(
        file_path=file_path,
        sheet_name=sheet_name,
        grade_column=grade_col,
        title=title,
        output_cell=output_cell,
    )
    print_result(result, json_output)


def cmd_chart_progress(
    file_path,
    student_col,
    grade_col,
    sheet_name,
    title,
    output_cell,
    show_labels,
    json_output=False,
):
    """Create a bar chart showing student progress"""
    if not OPENPYXL_AVAILABLE:
        return print_result(
            {"success": False, "error": "openpyxl not installed"}, json_output
        )

    result = doc_server.create_progress_chart(
        file_path=file_path,
        student_column=student_col,
        grade_column=grade_col,
        sheet_name=sheet_name,
        title=title,
        output_cell=output_cell,
        show_data_labels=show_labels,
    )
    print_result(result, json_output)


def cmd_pptx_create(output_path, title, subtitle="", json_output=False):
    """Create a new .pptx presentation"""
    if not PPTX_AVAILABLE:
        return print_result(
            {"success": False, "error": "python-pptx not installed"}, json_output
        )

    result = doc_server.create_presentation(output_path, title, subtitle)
    print_result(result, json_output)


def cmd_pptx_add_slide(
    file_path, title, content="", layout="content", json_output=False
):
    """Add a slide to a presentation"""
    if not PPTX_AVAILABLE:
        return print_result(
            {"success": False, "error": "python-pptx not installed"}, json_output
        )

    result = doc_server.add_slide(file_path, title, content, layout)
    print_result(result, json_output)


def cmd_pptx_add_bullets(file_path, title, bullets, json_output=False):
    """Add a bullet-point slide"""
    if not PPTX_AVAILABLE:
        return print_result(
            {"success": False, "error": "python-pptx not installed"}, json_output
        )

    # Handle both literal \n and actual newlines
    bullets = bullets.replace("\\n", "\n")
    result = doc_server.add_bullet_slide(file_path, title, bullets)
    print_result(result, json_output)


def cmd_pptx_read(file_path, json_output=False):
    """Read a presentation"""
    if not PPTX_AVAILABLE:
        return print_result(
            {"success": False, "error": "python-pptx not installed"}, json_output
        )

    result = doc_server.read_presentation(file_path)
    print_result(result, json_output)


def cmd_pptx_add_image(file_path, title, image_path, json_output=False):
    """Add an image slide"""
    if not PPTX_AVAILABLE:
        return print_result(
            {"success": False, "error": "python-pptx not installed"}, json_output
        )

    result = doc_server.add_image_slide(file_path, title, image_path)
    print_result(result, json_output)


def cmd_pptx_add_table(
    file_path, title, headers, data, coerce_rows=False, json_output=False
):
    """Add a table slide"""
    if not PPTX_AVAILABLE:
        return print_result(
            {"success": False, "error": "python-pptx not installed"}, json_output
        )

    result = doc_server.add_table_slide(
        file_path, title, headers, data, coerce_rows=coerce_rows
    )
    print_result(result, json_output)


def cmd_backup_list(file_path, limit=20, json_output=False):
    """List backups for a target file."""
    result = doc_server.list_backups(file_path, limit=int(limit))
    print_result(result, json_output)


def cmd_backup_prune(file_path=None, keep=20, days=None, json_output=False):
    """Prune backups by retention rules."""
    result = doc_server.prune_backups(
        file_path=file_path,
        keep=int(keep),
        older_than_days=int(days) if days is not None else None,
    )
    print_result(result, json_output)


def cmd_backup_restore(
    file_path, backup=None, latest=False, dry_run=False, json_output=False
):
    """Restore a file from backup."""
    result = doc_server.restore_backup(
        file_path=file_path,
        backup=backup,
        latest=latest,
        dry_run=dry_run,
    )
    print_result(result, json_output)


def cmd_info(file_path, json_output=False):
    """Get file information"""
    result = (
        doc_server.get_document_info(file_path)
        if doc_server
        else {"success": False, "error": "Client not available"}
    )
    print_result(result, json_output)


def cmd_status(json_output=False):
    """Check installation status"""
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
    import sys
    result = {
        "success": True,
        "version": "4.1.0",
        "python": sys.executable,
        "document_server": {
            "healthy": doc_server.check_health() if doc_server else False
        },
        "python_docx": DOCX_AVAILABLE,
        "openpyxl": OPENPYXL_AVAILABLE,
        "python_pptx": PPTX_AVAILABLE,
        "rdflib": rdflib_available,
        "rdflib_version": rdflib_version,
        "pyshacl": shacl_available,
        "capabilities": {
            "docx_create": DOCX_AVAILABLE,
            "docx_read": DOCX_AVAILABLE,
            "docx_edit": DOCX_AVAILABLE,
            "docx_tables": DOCX_AVAILABLE,
            "docx_formatting": DOCX_AVAILABLE,
            "docx_references": DOCX_AVAILABLE,
            "xlsx_create": OPENPYXL_AVAILABLE,
            "xlsx_read": OPENPYXL_AVAILABLE,
            "xlsx_edit": OPENPYXL_AVAILABLE,
            "xlsx_formulas": OPENPYXL_AVAILABLE,
            "xlsx_charts": OPENPYXL_AVAILABLE,
            "xlsx_stats": OPENPYXL_AVAILABLE,
            "xlsx_csv": OPENPYXL_AVAILABLE,
            "pptx_create": PPTX_AVAILABLE,
            "pptx_read": PPTX_AVAILABLE,
            "pptx_edit": PPTX_AVAILABLE,
            "pptx_notes": PPTX_AVAILABLE,
            "rdf_create": rdflib_available,
            "rdf_query": rdflib_available,
            "rdf_validate": shacl_available,
        },
        "total_commands": 103,
    }
    print_result(result, json_output)


def cmd_help(json_output=False):
    """Show help"""
    # Check rdflib availability
    try:
        import rdflib
        rdflib_available = True
    except ImportError:
        rdflib_available = False

    result = {
        "success": True,
        "version": "4.1.0",
        "categories": {
            "DOCUMENTS (.docx)": {
                "doc-create <file> <title> <content>": "Create new .docx document",
                "doc-read <file>": "Read and extract all content from .docx",
                "doc-append <file> <content>": "Append text to .docx",
                "doc-replace <file> <search> <replace>": "Find and replace in .docx (cross-run safe)",
                "doc-search <file> <text> [--case-sensitive]": "Search for text in document (paragraphs + tables)",
                "doc-insert <file> <text> <index> [--style <name>]": "Insert paragraph at specific position",
                "doc-delete <file> <index>": "Delete paragraph by index",
                "doc-format <file> <paragraph_index> [--bold] [--italic] [--underline] [--font-name <name>] [--font-size <n>] [--color <hex>] [--align <left|center|right|justify>]": "Apply paragraph formatting",
                "doc-highlight <file> <search_text> [--color <name>]": "Highlight matching text",
                "doc-comment <file> <comment> [--paragraph <index>]": "Attach OOXML comment annotation",
                "doc-layout <file> [--orientation portrait|landscape] [--margin-* <in>] [--header <text>] [--page-numbers]": "Set page layout/margins/header/footer",
                "doc-formatting-info <file>": "Inspect paragraph/section formatting",
                "doc-set-style <file> <index> <style>": "Set paragraph style (Heading 1, Normal, etc.)",
                "doc-list-styles <file>": "List all available paragraph/character styles",
                "doc-add-table <file> <headers_csv> <data_csv>": "Add table (rows separated by ;)",
                "doc-read-tables <file>": "Read all tables from document",
                "doc-add-image <file> <image_path> [--width <in>] [--caption <text>]": "Add image with optional caption",
                "doc-add-hyperlink <file> <text> <url> [--paragraph <index>]": "Add hyperlink (-1 = new paragraph)",
                "doc-add-page-break <file>": "Insert page break",
                "doc-add-list <file> <items_csv> [--type bullet|number]": "Add bulleted or numbered list (items separated by ;)",
                "doc-add-reference <file> <ref_json>": "Add reference to sidecar .refs.json",
                "doc-build-references <file>": "Build APA 7th References section from sidecar",
                "doc-set-metadata <file> [--author <a>] [--title <t>] [--subject <s>] [--keywords <k>]": "Set document properties",
                "doc-get-metadata <file>": "Read document properties",
                "doc-word-count <file>": "Word/character/paragraph counts",
                "doc-extract-images <file> <output_dir> [--format png|jpg] [--prefix <name>]": "Extract all embedded images from .docx",
                "doc-to-pdf <file> [output_path]": "Convert .docx to PDF via OnlyOffice",
            },
            "SPREADSHEETS (.xlsx)": {
                "xlsx-create <file> [sheet]": "Create new .xlsx spreadsheet",
                "xlsx-write <file> <headers> <data> [--sheet <name>] [--overwrite] [--coerce-rows] [--text-columns <csv>]": "Write data to sheet (safe update default)",
                "xlsx-read <file> [sheet]": "Read spreadsheet data (all sheets if no sheet specified)",
                "xlsx-append <file> <row-data> [--sheet <name>]": "Append row to spreadsheet",
                "xlsx-search <file> <text> [--sheet <name>]": "Search for text in spreadsheet",
                "xlsx-cell-read <file> <cell> [--sheet <name>]": "Read a single cell (e.g., B5)",
                "xlsx-cell-write <file> <cell> <value> [--sheet <name>] [--text]": "Write to a single cell",
                "xlsx-range-read <file> <range> [--sheet <name>]": "Read a cell range (e.g., A1:D10)",
                "xlsx-delete-rows <file> <start_row> [count] [--sheet <name>]": "Delete rows (1-indexed)",
                "xlsx-delete-cols <file> <start_col> [count] [--sheet <name>]": "Delete columns (1-indexed)",
                "xlsx-sort <file> <column> [--sheet <name>] [--desc] [--numeric]": "Sort by column (preserves header)",
                "xlsx-filter <file> <column> <op> <value> [--sheet <name>]": "Filter rows (op: eq|ne|gt|lt|ge|le|contains|startswith|endswith)",
                "xlsx-calc <file> <column> <op> [--sheet <name>] [--include-formulas] [--strict-formulas]": "Column statistics (sum|avg|min|max|all)",
                "xlsx-formula <file> <cell> <formula> [--sheet <name>]": "Add formula to cell",
                "xlsx-formula-audit <file> [--sheet <name>]": "Audit formula complexity/risk",
                "xlsx-freq <file> <column> [--sheet <name>] [--valid <csv>]": "Frequency + percentage table",
                "xlsx-corr <file> <x> <y> [--sheet <name>] [--method pearson|spearman]": "Correlation test with APA output",
                "xlsx-ttest <file> <val> <grp> <a> <b> [--sheet <name>] [--equal-var]": "Independent t-test (Welch default, Cohen's d)",
                "xlsx-mannwhitney <file> <val> <grp> <a> <b> [--sheet <name>]": "Mann-Whitney U test (non-parametric)",
                "xlsx-chi2 <file> <row> <col> [--sheet <name>] [--row-valid <csv>] [--col-valid <csv>]": "Chi-square test (Cramer's V)",
                "xlsx-research-pack <file> [--sheet <name>] [--profile hlth3112]": "Bundled analysis pack",
                "xlsx-text-extract <file> <column> [--sheet <name>] [--limit <n>]": "Extract open-text responses",
                "xlsx-text-keywords <file> <column> [--sheet <name>] [--top <n>]": "Keyword frequency for themes",
                "xlsx-sheet-list <file>": "List all sheets with row/column counts",
                "xlsx-sheet-add <file> <name> [--position <n>]": "Add a new sheet",
                "xlsx-sheet-delete <file> <name>": "Delete a sheet",
                "xlsx-sheet-rename <file> <old> <new>": "Rename a sheet",
                "xlsx-merge-cells <file> <range> [--sheet <name>]": "Merge cells (e.g., A1:C1)",
                "xlsx-unmerge-cells <file> <range> [--sheet <name>]": "Unmerge cells",
                "xlsx-format-cells <file> <range> [--sheet <name>] [--bold] [--italic] [--font-name <n>] [--font-size <n>] [--color <hex>] [--bg-color <hex>] [--number-format <fmt>] [--wrap]": "Format cell range",
                "xlsx-csv-import <xlsx_file> <csv_file> [--sheet <name>] [--delimiter <char>]": "Import CSV into xlsx",
                "xlsx-csv-export <xlsx_file> <csv_file> [--sheet <name>] [--delimiter <char>]": "Export sheet to CSV",
                "xlsx-add-validation <file> <range> <type> [--operator <op>] [--formula1 <v>] [--formula2 <v>] [--error <msg>]": "Add data validation (list|whole|decimal|textLength|date|custom)",
                "xlsx-add-dropdown <file> <range> <options_csv> [--prompt <msg>]": "Add dropdown list validation (shortcut)",
                "xlsx-list-validations <file> [--sheet <name>]": "List all validation rules on a sheet",
                "xlsx-remove-validation <file> [--range <range>] [--all]": "Remove validation rules",
                "xlsx-validate-data <file> [--sheet <name>] [--max-rows <n>]": "Audit data against validation rules (pass/fail per cell)",
            },
            "CHARTS (.xlsx)": {
                "chart-create <file> <type> <data_range> <cat_range> <title> [options]": "Create chart (bar|line|pie|scatter|bar_horizontal)",
                "chart-comparison <file> <type> <title> [options]": "Comparison chart from structured data",
                "chart-grade-dist <file> <grade_col> <title>": "Grade distribution pie chart",
                "chart-progress <file> <student_col> <grade_col> <title>": "Student progress bar chart",
            },
            "PRESENTATIONS (.pptx)": {
                "pptx-create <file> <title> [subtitle]": "Create new presentation",
                "pptx-add-slide <file> <title> [content] [layout]": "Add slide (title_only|content|blank|two_content|comparison)",
                "pptx-add-bullets <file> <title> <bullets>": "Bullet-point slide (\\n separated)",
                "pptx-read <file>": "Read all slides and content",
                "pptx-add-image <file> <title> <image_path>": "Add image slide",
                "pptx-add-table <file> <title> <headers> <data> [--coerce-rows]": "Add table slide",
                "pptx-delete-slide <file> <index>": "Delete slide by index (0-based)",
                "pptx-speaker-notes <file> <index> [text]": "Read or set speaker notes (omit text to read)",
                "pptx-update-text <file> <index> [--title <t>] [--body <b>]": "Update existing slide text",
                "pptx-slide-count <file>": "Get slide count and titles",
                "pptx-extract-images <file> <output_dir> [--slide <index>] [--format png|jpg]": "Extract all images from slides",
                "pptx-list-shapes <file> [--slide <index>]": "List all shapes with position/size/text (spatial map)",
                "pptx-add-textbox <file> <slide_index> <text> [--left <in>] [--top <in>] [--width <in>] [--height <in>] [--font-size <pt>] [--bold] [--italic] [--color <hex>] [--align <dir>]": "Add textbox at exact position",
                "pptx-modify-shape <file> <slide_index> <shape_name> [--left <in>] [--top <in>] [--width <in>] [--height <in>] [--text <text>] [--font-size <pt>] [--rotation <deg>]": "Move/resize/edit any shape by name",
                "pptx-preview <file> <output_dir> [--slide <index>] [--dpi <n>]": "Render slides as PNG images (requires OnlyOffice Docker)",
            },
            "PDF (.pdf) — Image Extraction": {
                "pdf-extract-images <file> <output_dir> [--format png|jpg] [--pages <range>]": "Extract embedded images from PDF (PyMuPDF)",
                "pdf-page-to-image <file> <output_dir> [--pages <range>] [--dpi <n>] [--format png|jpg]": "Render full PDF pages as images",
            },
            "RDF (Knowledge Graphs)": {
                "rdf-create <file> [--base <uri>] [--format turtle|xml|n3|json-ld] [--prefix <p>=<uri>]": "Create empty RDF graph with prefixes",
                "rdf-read <file> [--limit <n>]": "Read/parse RDF file and show triples",
                "rdf-add <file> <subject> <predicate> <object> [--type uri|literal|bnode] [--lang <l>] [--datatype <dt>]": "Add a triple",
                "rdf-remove <file> [--subject <s>] [--predicate <p>] [--object <o>] [--type uri|literal|bnode]": "Remove triples (None = wildcard)",
                "rdf-query <file> <sparql> [--limit <n>]": "Execute SPARQL query",
                "rdf-export <file> <output> [--format turtle|xml|n3|nt|json-ld|trig]": "Convert/export to different format",
                "rdf-merge <file_a> <file_b> [--output <file>] [--format turtle]": "Merge two RDF graphs",
                "rdf-stats <file>": "Graph statistics (subjects, predicates, types, etc.)",
                "rdf-namespace <file> [<prefix> <uri>]": "Add namespace prefix or list all",
                "rdf-validate <file> <shapes_file>": "SHACL validation (requires pyshacl)",
            },
            "GENERAL": {
                "list": "List recent documents, spreadsheets, presentations",
                "open <file> [gui|web]": "Open in OnlyOffice GUI or web viewer",
                "watch <file> [gui|web]": "Watch file for changes + auto-open",
                "info <file>": "File info (type, size, sheet/slide/paragraph counts)",
                "backup-list <file> [--limit <n>]": "List backups for file",
                "backup-prune [--file <f>] [--keep <n>] [--days <n>]": "Prune old backups",
                "backup-restore <file> [--backup <p>] [--latest] [--dry-run]": "Restore from backup",
                "status": "Check installation and all capabilities",
                "help": "Show this help",
            },
        },
        "capabilities": {
            "python_docx": DOCX_AVAILABLE,
            "openpyxl": OPENPYXL_AVAILABLE,
            "python_pptx": PPTX_AVAILABLE,
            "rdflib": rdflib_available,
        },
        "total_commands": 103,
        "examples": [
            "# Documents",
            "cli-anything-onlyoffice doc-create /tmp/essay.docx 'My Essay' 'Introduction paragraph here'",
            "cli-anything-onlyoffice doc-insert /tmp/essay.docx 'New first paragraph' 0 --style 'Heading 1'",
            "cli-anything-onlyoffice doc-search /tmp/essay.docx 'introduction'",
            "cli-anything-onlyoffice doc-add-hyperlink /tmp/essay.docx 'Click here' 'https://example.com'",
            "cli-anything-onlyoffice doc-add-list /tmp/essay.docx 'First item;Second item;Third item' --type bullet",
            "cli-anything-onlyoffice doc-word-count /tmp/essay.docx",
            "# Spreadsheets",
            "cli-anything-onlyoffice xlsx-write /tmp/data.xlsx 'Name,Score' 'Alice,90;Bob,85' --sheet Grades",
            "cli-anything-onlyoffice xlsx-cell-read /tmp/data.xlsx B2 --sheet Grades",
            "cli-anything-onlyoffice xlsx-cell-write /tmp/data.xlsx C2 95 --sheet Grades",
            "cli-anything-onlyoffice xlsx-sort /tmp/data.xlsx B --desc --numeric",
            "cli-anything-onlyoffice xlsx-filter /tmp/data.xlsx B gt 80 --sheet Grades",
            "cli-anything-onlyoffice xlsx-sheet-list /tmp/data.xlsx",
            "cli-anything-onlyoffice xlsx-format-cells /tmp/data.xlsx A1:B1 --bold --bg-color 4472C4 --color FFFFFF",
            "cli-anything-onlyoffice xlsx-csv-import /tmp/data.xlsx /tmp/raw.csv --sheet Imported",
            "# Presentations",
            "cli-anything-onlyoffice pptx-create /tmp/lecture.pptx 'Lecture 1' 'Introduction'",
            "cli-anything-onlyoffice pptx-speaker-notes /tmp/lecture.pptx 0 'Remember to introduce yourself'",
            "cli-anything-onlyoffice pptx-slide-count /tmp/lecture.pptx",
            "# RDF Knowledge Graphs",
            "cli-anything-onlyoffice rdf-create /tmp/knowledge.ttl --base http://example.org/",
            "cli-anything-onlyoffice rdf-add /tmp/knowledge.ttl http://example.org/Alice http://xmlns.com/foaf/0.1/name Alice --type literal",
            "cli-anything-onlyoffice rdf-query /tmp/knowledge.ttl 'SELECT ?s ?p ?o WHERE { ?s ?p ?o } LIMIT 10'",
            "cli-anything-onlyoffice rdf-stats /tmp/knowledge.ttl",
        ],
    }
    print_result(result, json_output)


# ==================== MAIN ====================


def main():
    parser = argparse.ArgumentParser(
        description="CLI-Anything OnlyOffice v4.0 - Documents, Spreadsheets, Presentations, RDF", add_help=False
    )
    parser.add_argument("command", nargs="?", default="help", help="Command")
    parser.add_argument("args", nargs="*", default=[], help="Arguments")
    parser.add_argument("--json", action="store_true", help="JSON output")

    # Parse known global flags only; keep raw command args order from sys.argv.
    args, unknown = parser.parse_known_args()
    args.args = [
        a for a in (sys.argv[2:] if len(sys.argv) > 2 else []) if a != "--json"
    ]
    json_output = args.json

    # Document commands
    if args.command == "doc-create":
        if len(args.args) < 3:
            print_result(
                {
                    "success": False,
                    "error": "Usage: doc-create <file> <title> <content>",
                },
                json_output,
            )
        else:
            cmd_doc_create(
                args.args[0], args.args[1], " ".join(args.args[2:]), json_output
            )

    elif args.command == "doc-read":
        if not args.args:
            print_result(
                {"success": False, "error": "Usage: doc-read <file>"}, json_output
            )
        else:
            cmd_doc_read(args.args[0], json_output)

    elif args.command == "doc-append":
        if len(args.args) < 2:
            print_result(
                {"success": False, "error": "Usage: doc-append <file> <content>"},
                json_output,
            )
        else:
            cmd_doc_append(args.args[0], " ".join(args.args[1:]), json_output)

    elif args.command == "doc-replace":
        if len(args.args) < 3:
            print_result(
                {
                    "success": False,
                    "error": "Usage: doc-replace <file> <search> <replace>",
                },
                json_output,
            )
        else:
            cmd_doc_replace(
                args.args[0], args.args[1], " ".join(args.args[2:]), json_output
            )

    elif args.command == "doc-format":
        if len(args.args) < 2:
            print_result(
                {
                    "success": False,
                    "error": "Usage: doc-format <file> <paragraph_index> [--bold] [--italic] [--underline] [--font-name <name>] [--font-size <n>] [--color <hex>] [--align <left|center|right|justify>]",
                },
                json_output,
            )
        else:
            file_path = args.args[0]
            paragraph_index = args.args[1]
            opts = {
                "bold": False,
                "italic": False,
                "underline": False,
                "font_name": None,
                "font_size": None,
                "color": None,
                "alignment": None,
            }
            i = 2
            while i < len(args.args):
                if args.args[i] == "--bold":
                    opts["bold"] = True
                    i += 1
                elif args.args[i] == "--italic":
                    opts["italic"] = True
                    i += 1
                elif args.args[i] == "--underline":
                    opts["underline"] = True
                    i += 1
                elif args.args[i] == "--font-name" and i + 1 < len(args.args):
                    opts["font_name"] = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--font-size" and i + 1 < len(args.args):
                    opts["font_size"] = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--color" and i + 1 < len(args.args):
                    opts["color"] = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--align" and i + 1 < len(args.args):
                    opts["alignment"] = args.args[i + 1]
                    i += 2
                else:
                    i += 1
            cmd_doc_format(file_path, paragraph_index, json_output=json_output, **opts)

    elif args.command == "doc-highlight":
        if len(args.args) < 2:
            print_result(
                {
                    "success": False,
                    "error": "Usage: doc-highlight <file> <search_text> [--color <name>]",
                },
                json_output,
            )
        else:
            file_path = args.args[0]
            search_text = args.args[1]
            color = "yellow"
            if len(args.args) >= 4 and args.args[2] == "--color":
                color = args.args[3]
            cmd_doc_highlight(
                file_path, search_text, color=color, json_output=json_output
            )

    elif args.command == "doc-comment":
        if len(args.args) < 2:
            print_result(
                {
                    "success": False,
                    "error": "Usage: doc-comment <file> <comment> [--paragraph <index>]",
                },
                json_output,
            )
        else:
            file_path = args.args[0]
            comment = args.args[1]
            paragraph = 0
            if len(args.args) >= 4 and args.args[2] == "--paragraph":
                paragraph = int(args.args[3])
            cmd_doc_comment(
                file_path, comment, paragraph_index=paragraph, json_output=json_output
            )

    elif args.command == "doc-add-reference":
        if len(args.args) < 2:
            print_result(
                {
                    "success": False,
                    "error": 'Usage: doc-add-reference <file> <ref_json>  (ref_json: {"author":"...","year":"...","title":"...","source":"...","type":"journal|book|website|report|chapter","doi":"..."})',
                },
                json_output,
            )
        else:
            result = doc_server.add_reference(args.args[0], args.args[1])
            print_result(result, json_output)

    elif args.command == "doc-build-references":
        if len(args.args) < 1:
            print_result(
                {
                    "success": False,
                    "error": "Usage: doc-build-references <file>  (reads <file>.refs.json, appends APA 7th formatted References section)",
                },
                json_output,
            )
        else:
            result = doc_server.build_references(args.args[0])
            print_result(result, json_output)

    elif args.command == "doc-add-table":
        if len(args.args) < 3:
            print_result(
                {
                    "success": False,
                    "error": "Usage: doc-add-table <file> <headers_csv> <data_csv>  (rows separated by ';')",
                },
                json_output,
            )
        else:
            result = doc_server.add_table(args.args[0], args.args[1], args.args[2])
            print_result(result, json_output)

    elif args.command == "doc-set-style":
        if len(args.args) < 3:
            print_result(
                {
                    "success": False,
                    "error": 'Usage: doc-set-style <file> <paragraph_index> <style>  (e.g., "Heading 1", "Heading 2", "Normal", "Title")',
                },
                json_output,
            )
        else:
            result = doc_server.set_paragraph_style(args.args[0], int(args.args[1]), args.args[2])
            print_result(result, json_output)

    elif args.command == "doc-add-image":
        if len(args.args) < 2:
            print_result(
                {
                    "success": False,
                    "error": "Usage: doc-add-image <file> <image_path> [--width <inches>] [--caption <text>]",
                },
                json_output,
            )
        else:
            width = 5.5
            caption = None
            i = 2
            while i < len(args.args):
                if args.args[i] == "--width" and i + 1 < len(args.args):
                    width = float(args.args[i + 1])
                    i += 2
                elif args.args[i] == "--caption" and i + 1 < len(args.args):
                    caption = args.args[i + 1]
                    i += 2
                else:
                    i += 1
            result = doc_server.add_image(args.args[0], args.args[1], width_inches=width, caption=caption)
            print_result(result, json_output)

    elif args.command == "doc-layout":
        if len(args.args) < 1:
            print_result(
                {
                    "success": False,
                    "error": "Usage: doc-layout <file> [--orientation <portrait|landscape>] [--margin-top <in>] [--margin-bottom <in>] [--margin-left <in>] [--margin-right <in>] [--header <text>] [--page-numbers]",
                },
                json_output,
            )
        else:
            file_path = args.args[0]
            orientation = "portrait"
            margins = {}
            header_text = None
            page_numbers = False
            i = 1
            while i < len(args.args):
                if args.args[i] == "--orientation" and i + 1 < len(args.args):
                    orientation = args.args[i + 1]
                    i += 2
                elif args.args[i].startswith("--margin-") and i + 1 < len(args.args):
                    side = args.args[i].replace("--margin-", "")
                    margins[side] = float(args.args[i + 1])
                    i += 2
                elif args.args[i] == "--header" and i + 1 < len(args.args):
                    header_text = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--page-numbers":
                    page_numbers = True
                    i += 1
                else:
                    i += 1
            result = doc_server.set_page_layout(
                file_path,
                orientation=orientation,
                margins=margins,
                header_text=header_text,
                page_numbers=page_numbers,
            )
            print_result(result, json_output)

    elif args.command == "doc-formatting-info":
        if not args.args:
            print_result(
                {"success": False, "error": "Usage: doc-formatting-info <file>"},
                json_output,
            )
        else:
            cmd_doc_formatting_info(args.args[0], json_output=json_output)

    elif args.command == "doc-search":
        if len(args.args) < 2:
            print_result({"success": False, "error": "Usage: doc-search <file> <text> [--case-sensitive]"}, json_output)
        else:
            case_sensitive = "--case-sensitive" in args.args[2:]
            result = doc_server.search_document(args.args[0], args.args[1], case_sensitive=case_sensitive)
            print_result(result, json_output)

    elif args.command == "doc-insert":
        if len(args.args) < 3:
            print_result({"success": False, "error": "Usage: doc-insert <file> <text> <index> [--style <name>]"}, json_output)
        else:
            style = None
            i = 3
            while i < len(args.args):
                if args.args[i] == "--style" and i + 1 < len(args.args):
                    style = args.args[i + 1]
                    i += 2
                else:
                    i += 1
            result = doc_server.insert_paragraph(args.args[0], args.args[1], int(args.args[2]), style=style)
            print_result(result, json_output)

    elif args.command == "doc-delete":
        if len(args.args) < 2:
            print_result({"success": False, "error": "Usage: doc-delete <file> <paragraph_index>"}, json_output)
        else:
            result = doc_server.delete_paragraph(args.args[0], int(args.args[1]))
            print_result(result, json_output)

    elif args.command == "doc-read-tables":
        if not args.args:
            print_result({"success": False, "error": "Usage: doc-read-tables <file>"}, json_output)
        else:
            result = doc_server.read_tables(args.args[0])
            print_result(result, json_output)

    elif args.command == "doc-add-hyperlink":
        if len(args.args) < 3:
            print_result({"success": False, "error": "Usage: doc-add-hyperlink <file> <text> <url> [--paragraph <index>]"}, json_output)
        else:
            para_idx = -1
            i = 3
            while i < len(args.args):
                if args.args[i] == "--paragraph" and i + 1 < len(args.args):
                    para_idx = int(args.args[i + 1])
                    i += 2
                else:
                    i += 1
            result = doc_server.add_hyperlink(args.args[0], args.args[1], args.args[2], paragraph_index=para_idx)
            print_result(result, json_output)

    elif args.command == "doc-add-page-break":
        if not args.args:
            print_result({"success": False, "error": "Usage: doc-add-page-break <file>"}, json_output)
        else:
            result = doc_server.add_page_break(args.args[0])
            print_result(result, json_output)

    elif args.command == "doc-list-styles":
        if not args.args:
            print_result({"success": False, "error": "Usage: doc-list-styles <file>"}, json_output)
        else:
            result = doc_server.list_styles(args.args[0])
            print_result(result, json_output)

    elif args.command == "doc-add-list":
        if len(args.args) < 2:
            print_result({"success": False, "error": "Usage: doc-add-list <file> <items> [--type bullet|number] (items separated by ;)"}, json_output)
        else:
            list_type = "bullet"
            i = 2
            while i < len(args.args):
                if args.args[i] == "--type" and i + 1 < len(args.args):
                    list_type = args.args[i + 1]
                    i += 2
                else:
                    i += 1
            items = [item.strip() for item in args.args[1].split(";") if item.strip()]
            result = doc_server.add_list(args.args[0], items, list_type=list_type)
            print_result(result, json_output)

    elif args.command == "doc-set-metadata":
        if not args.args:
            print_result({"success": False, "error": "Usage: doc-set-metadata <file> [--author <a>] [--title <t>] [--subject <s>] [--keywords <k>] [--comments <c>] [--category <cat>]"}, json_output)
        else:
            opts = {}
            i = 1
            while i < len(args.args):
                for key in ("--author", "--title", "--subject", "--keywords", "--comments", "--category"):
                    if args.args[i] == key and i + 1 < len(args.args):
                        opts[key.lstrip("-")] = args.args[i + 1]
                        i += 2
                        break
                else:
                    i += 1
            result = doc_server.set_metadata(args.args[0], **opts)
            print_result(result, json_output)

    elif args.command == "doc-get-metadata":
        if not args.args:
            print_result({"success": False, "error": "Usage: doc-get-metadata <file>"}, json_output)
        else:
            result = doc_server.get_metadata(args.args[0])
            print_result(result, json_output)

    elif args.command == "doc-word-count":
        if not args.args:
            print_result({"success": False, "error": "Usage: doc-word-count <file>"}, json_output)
        else:
            result = doc_server.word_count(args.args[0])
            print_result(result, json_output)

    # Spreadsheet commands
    elif args.command == "xlsx-create":
        if not args.args:
            print_result(
                {"success": False, "error": "Usage: xlsx-create <file> [sheet]"},
                json_output,
            )
        else:
            cmd_xlsx_create(
                args.args[0],
                args.args[1] if len(args.args) > 1 else "Sheet1",
                json_output,
            )

    elif args.command == "xlsx-write":
        if len(args.args) < 3:
            print_result(
                {
                    "success": False,
                    "error": "Usage: xlsx-write <file> <headers> <data> [--sheet <name>] [--overwrite] [--coerce-rows]",
                },
                json_output,
            )
        else:
            sheet = "Sheet1"
            overwrite = False
            coerce_rows = False
            text_columns = None
            i = 3
            while i < len(args.args):
                if args.args[i] == "--sheet" and i + 1 < len(args.args):
                    sheet = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--overwrite":
                    overwrite = True
                    i += 1
                elif args.args[i] == "--coerce-rows":
                    coerce_rows = True
                    i += 1
                elif args.args[i] == "--text-columns" and i + 1 < len(args.args):
                    text_columns = [c.strip() for c in args.args[i + 1].split(",")]
                    i += 2
                else:
                    i += 1
            cmd_xlsx_write(
                args.args[0],
                args.args[1],
                args.args[2],
                sheet_name=sheet,
                overwrite_workbook=overwrite,
                coerce_rows=coerce_rows,
                text_columns=text_columns,
                json_output=json_output,
            )

    elif args.command == "xlsx-read":
        if not args.args:
            print_result(
                {"success": False, "error": "Usage: xlsx-read <file> [sheet]"},
                json_output,
            )
        else:
            result = doc_server.read_spreadsheet(
                args.args[0], args.args[1] if len(args.args) > 1 else None
            )
            print_result(result, json_output)

    elif args.command == "xlsx-append":
        if len(args.args) < 2:
            print_result(
                {
                    "success": False,
                    "error": "Usage: xlsx-append <file> <row-data> [--sheet <name>]",
                },
                json_output,
            )
        else:
            sheet = None
            if len(args.args) > 3 and args.args[2] == "--sheet":
                sheet = args.args[3]
            cmd_xlsx_append(
                args.args[0], args.args[1], sheet_name=sheet, json_output=json_output
            )

    elif args.command == "xlsx-search":
        if len(args.args) < 2:
            print_result(
                {
                    "success": False,
                    "error": "Usage: xlsx-search <file> <search-text> [--sheet <name>]",
                },
                json_output,
            )
        else:
            sheet = None
            if len(args.args) > 3 and args.args[2] == "--sheet":
                sheet = args.args[3]
            cmd_xlsx_search(
                args.args[0], args.args[1], sheet_name=sheet, json_output=json_output
            )

    elif args.command == "xlsx-calc":
        if len(args.args) < 3:
            print_result(
                {
                    "success": False,
                    "error": "Usage: xlsx-calc <file> <column> <operation> [--sheet <name>] [--include-formulas] [--strict-formulas]",
                },
                json_output,
            )
        else:
            sheet = "Sheet1"
            include_formulas = False
            strict_formulas = False
            i = 3
            while i < len(args.args):
                if args.args[i] == "--sheet" and i + 1 < len(args.args):
                    sheet = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--include-formulas":
                    include_formulas = True
                    i += 1
                elif args.args[i] == "--strict-formulas":
                    strict_formulas = True
                    i += 1
                else:
                    i += 1
            cmd_xlsx_calc(
                args.args[0],
                args.args[1],
                args.args[2],
                sheet_name=sheet,
                include_formulas=include_formulas,
                strict_formulas=strict_formulas,
                json_output=json_output,
            )

    elif args.command == "xlsx-formula":
        if len(args.args) < 3:
            print_result(
                {
                    "success": False,
                    "error": "Usage: xlsx-formula <file> <cell> <formula> [--sheet <name>]",
                },
                json_output,
            )
        else:
            sheet = "Sheet1"
            if len(args.args) > 4 and args.args[3] == "--sheet":
                sheet = args.args[4]
            cmd_xlsx_formula(
                args.args[0],
                args.args[1],
                args.args[2],
                sheet_name=sheet,
                json_output=json_output,
            )

    elif args.command == "xlsx-freq":
        if len(args.args) < 2:
            print_result(
                {
                    "success": False,
                    "error": "Usage: xlsx-freq <file> <column> [--sheet <name>] [--valid <csv>]",
                },
                json_output,
            )
        else:
            sheet = "Sheet1"
            valid_values = None
            i = 2
            while i < len(args.args):
                if args.args[i] == "--sheet" and i + 1 < len(args.args):
                    sheet = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--valid" and i + 1 < len(args.args):
                    valid_values = [
                        v.strip() for v in args.args[i + 1].split(",") if v.strip()
                    ]
                    i += 2
                else:
                    i += 1
            cmd_xlsx_freq(
                args.args[0],
                args.args[1],
                sheet_name=sheet,
                valid_values=valid_values,
                json_output=json_output,
            )

    elif args.command == "xlsx-corr":
        if len(args.args) < 3:
            print_result(
                {
                    "success": False,
                    "error": "Usage: xlsx-corr <file> <x_col> <y_col> [--sheet <name>] [--method pearson|spearman]",
                },
                json_output,
            )
        else:
            sheet = "Sheet1"
            method = "pearson"
            i = 3
            while i < len(args.args):
                if args.args[i] == "--sheet" and i + 1 < len(args.args):
                    sheet = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--method" and i + 1 < len(args.args):
                    method = args.args[i + 1]
                    i += 2
                else:
                    i += 1
            cmd_xlsx_corr(
                args.args[0],
                args.args[1],
                args.args[2],
                sheet_name=sheet,
                method=method,
                json_output=json_output,
            )

    elif args.command == "xlsx-ttest":
        if len(args.args) < 5:
            print_result(
                {
                    "success": False,
                    "error": "Usage: xlsx-ttest <file> <value_col> <group_col> <group_a> <group_b> [--sheet <name>] [--equal-var]",
                },
                json_output,
            )
        else:
            sheet = "Sheet1"
            equal_var = False
            i = 5
            while i < len(args.args):
                if args.args[i] == "--sheet" and i + 1 < len(args.args):
                    sheet = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--equal-var":
                    equal_var = True
                    i += 1
                else:
                    i += 1
            cmd_xlsx_ttest(
                args.args[0],
                args.args[1],
                args.args[2],
                args.args[3],
                args.args[4],
                sheet_name=sheet,
                equal_var=equal_var,
                json_output=json_output,
            )

    elif args.command == "xlsx-mannwhitney":
        if len(args.args) < 5:
            print_result(
                {
                    "success": False,
                    "error": "Usage: xlsx-mannwhitney <file> <value_col> <group_col> <group_a> <group_b> [--sheet <name>]",
                },
                json_output,
            )
        else:
            sheet = "Sheet1"
            i = 5
            while i < len(args.args):
                if args.args[i] == "--sheet" and i + 1 < len(args.args):
                    sheet = args.args[i + 1]
                    i += 2
                else:
                    i += 1
            result = doc_server.mann_whitney_test(
                args.args[0],
                args.args[1],
                args.args[2],
                args.args[3],
                args.args[4],
                sheet_name=sheet,
            )
            print_result(result, json_output)

    elif args.command == "xlsx-chi2":
        if len(args.args) < 3:
            print_result(
                {
                    "success": False,
                    "error": "Usage: xlsx-chi2 <file> <row_col> <col_col> [--sheet <name>] [--row-valid <csv>] [--col-valid <csv>]",
                },
                json_output,
            )
        else:
            sheet = "Sheet1"
            row_valid_values = None
            col_valid_values = None
            i = 3
            while i < len(args.args):
                if args.args[i] == "--sheet" and i + 1 < len(args.args):
                    sheet = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--row-valid" and i + 1 < len(args.args):
                    row_valid_values = [
                        v.strip() for v in args.args[i + 1].split(",") if v.strip()
                    ]
                    i += 2
                elif args.args[i] == "--col-valid" and i + 1 < len(args.args):
                    col_valid_values = [
                        v.strip() for v in args.args[i + 1].split(",") if v.strip()
                    ]
                    i += 2
                else:
                    i += 1
            cmd_xlsx_chi2(
                args.args[0],
                args.args[1],
                args.args[2],
                sheet_name=sheet,
                row_valid_values=row_valid_values,
                col_valid_values=col_valid_values,
                json_output=json_output,
            )

    elif args.command == "xlsx-research-pack":
        if len(args.args) < 1:
            print_result(
                {
                    "success": False,
                    "error": "Usage: xlsx-research-pack <file> [--sheet <name>] [--profile hlth3112] [--require-formula-safe]",
                },
                json_output,
            )
        else:
            sheet = "Sheet0"
            profile = "hlth3112"
            require_formula_safe = False
            i = 1
            while i < len(args.args):
                if args.args[i] == "--sheet" and i + 1 < len(args.args):
                    sheet = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--profile" and i + 1 < len(args.args):
                    profile = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--require-formula-safe":
                    require_formula_safe = True
                    i += 1
                else:
                    i += 1
            cmd_xlsx_research_pack(
                args.args[0],
                sheet_name=sheet,
                profile=profile,
                require_formula_safe=require_formula_safe,
                json_output=json_output,
            )

    elif args.command == "xlsx-formula-audit":
        if len(args.args) < 1:
            print_result(
                {
                    "success": False,
                    "error": "Usage: xlsx-formula-audit <file> [--sheet <name>]",
                },
                json_output,
            )
        else:
            sheet = None
            i = 1
            while i < len(args.args):
                if args.args[i] == "--sheet" and i + 1 < len(args.args):
                    sheet = args.args[i + 1]
                    i += 2
                else:
                    i += 1
            cmd_xlsx_formula_audit(
                args.args[0],
                sheet_name=sheet,
                json_output=json_output,
            )

    elif args.command == "xlsx-text-extract":
        if len(args.args) < 2:
            print_result(
                {
                    "success": False,
                    "error": "Usage: xlsx-text-extract <file> <column> [--sheet <name>] [--limit <n>] [--min-length <n>]",
                },
                json_output,
            )
        else:
            sheet = "Sheet0"
            limit = 20
            min_length = 20
            i = 2
            while i < len(args.args):
                if args.args[i] == "--sheet" and i + 1 < len(args.args):
                    sheet = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--limit" and i + 1 < len(args.args):
                    limit = int(args.args[i + 1])
                    i += 2
                elif args.args[i] == "--min-length" and i + 1 < len(args.args):
                    min_length = int(args.args[i + 1])
                    i += 2
                else:
                    i += 1
            cmd_xlsx_text_extract(
                args.args[0],
                args.args[1],
                sheet_name=sheet,
                limit=limit,
                min_length=min_length,
                json_output=json_output,
            )

    elif args.command == "xlsx-text-keywords":
        if len(args.args) < 2:
            print_result(
                {
                    "success": False,
                    "error": "Usage: xlsx-text-keywords <file> <column> [--sheet <name>] [--top <n>] [--min-word-length <n>]",
                },
                json_output,
            )
        else:
            sheet = "Sheet0"
            top = 15
            min_word_length = 4
            i = 2
            while i < len(args.args):
                if args.args[i] == "--sheet" and i + 1 < len(args.args):
                    sheet = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--top" and i + 1 < len(args.args):
                    top = int(args.args[i + 1])
                    i += 2
                elif args.args[i] == "--min-word-length" and i + 1 < len(args.args):
                    min_word_length = int(args.args[i + 1])
                    i += 2
                else:
                    i += 1
            cmd_xlsx_text_keywords(
                args.args[0],
                args.args[1],
                sheet_name=sheet,
                top=top,
                min_word_length=min_word_length,
                json_output=json_output,
            )

    # Extended spreadsheet commands
    elif args.command == "xlsx-cell-read":
        if len(args.args) < 2:
            print_result({"success": False, "error": "Usage: xlsx-cell-read <file> <cell> [--sheet <name>]"}, json_output)
        else:
            sheet = None
            if len(args.args) > 3 and args.args[2] == "--sheet":
                sheet = args.args[3]
            result = doc_server.cell_read(args.args[0], args.args[1], sheet_name=sheet)
            print_result(result, json_output)

    elif args.command == "xlsx-cell-write":
        if len(args.args) < 3:
            print_result({"success": False, "error": "Usage: xlsx-cell-write <file> <cell> <value> [--sheet <name>] [--text]"}, json_output)
        else:
            sheet = None
            as_text = False
            i = 3
            while i < len(args.args):
                if args.args[i] == "--sheet" and i + 1 < len(args.args):
                    sheet = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--text":
                    as_text = True
                    i += 1
                else:
                    i += 1
            result = doc_server.cell_write(args.args[0], args.args[1], args.args[2], sheet_name=sheet, as_text=as_text)
            print_result(result, json_output)

    elif args.command == "xlsx-range-read":
        if len(args.args) < 2:
            print_result({"success": False, "error": "Usage: xlsx-range-read <file> <range> [--sheet <name>]"}, json_output)
        else:
            sheet = None
            if len(args.args) > 3 and args.args[2] == "--sheet":
                sheet = args.args[3]
            result = doc_server.range_read(args.args[0], args.args[1], sheet_name=sheet)
            print_result(result, json_output)

    elif args.command == "xlsx-delete-rows":
        if len(args.args) < 2:
            print_result({"success": False, "error": "Usage: xlsx-delete-rows <file> <start_row> [count] [--sheet <name>]"}, json_output)
        else:
            count = 1
            sheet = None
            i = 2
            while i < len(args.args):
                if args.args[i] == "--sheet" and i + 1 < len(args.args):
                    sheet = args.args[i + 1]
                    i += 2
                elif not args.args[i].startswith("--"):
                    count = int(args.args[i])
                    i += 1
                else:
                    i += 1
            result = doc_server.delete_rows(args.args[0], int(args.args[1]), count=count, sheet_name=sheet)
            print_result(result, json_output)

    elif args.command == "xlsx-delete-cols":
        if len(args.args) < 2:
            print_result({"success": False, "error": "Usage: xlsx-delete-cols <file> <start_col> [count] [--sheet <name>]"}, json_output)
        else:
            count = 1
            sheet = None
            i = 2
            while i < len(args.args):
                if args.args[i] == "--sheet" and i + 1 < len(args.args):
                    sheet = args.args[i + 1]
                    i += 2
                elif not args.args[i].startswith("--"):
                    count = int(args.args[i])
                    i += 1
                else:
                    i += 1
            result = doc_server.delete_columns(args.args[0], int(args.args[1]), count=count, sheet_name=sheet)
            print_result(result, json_output)

    elif args.command == "xlsx-sort":
        if len(args.args) < 2:
            print_result({"success": False, "error": "Usage: xlsx-sort <file> <column> [--sheet <name>] [--desc] [--numeric]"}, json_output)
        else:
            sheet = None
            descending = False
            numeric = False
            i = 2
            while i < len(args.args):
                if args.args[i] == "--sheet" and i + 1 < len(args.args):
                    sheet = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--desc":
                    descending = True
                    i += 1
                elif args.args[i] == "--numeric":
                    numeric = True
                    i += 1
                else:
                    i += 1
            result = doc_server.sort_sheet(args.args[0], args.args[1], sheet_name=sheet, descending=descending, numeric=numeric)
            print_result(result, json_output)

    elif args.command == "xlsx-filter":
        if len(args.args) < 4:
            print_result({"success": False, "error": "Usage: xlsx-filter <file> <column> <op> <value> [--sheet <name>] (op: eq|ne|gt|lt|ge|le|contains|startswith|endswith)"}, json_output)
        else:
            sheet = None
            i = 4
            while i < len(args.args):
                if args.args[i] == "--sheet" and i + 1 < len(args.args):
                    sheet = args.args[i + 1]
                    i += 2
                else:
                    i += 1
            result = doc_server.filter_rows(args.args[0], args.args[1], args.args[2], args.args[3], sheet_name=sheet)
            print_result(result, json_output)

    elif args.command == "xlsx-sheet-list":
        if not args.args:
            print_result({"success": False, "error": "Usage: xlsx-sheet-list <file>"}, json_output)
        else:
            result = doc_server.sheet_list(args.args[0])
            print_result(result, json_output)

    elif args.command == "xlsx-sheet-add":
        if len(args.args) < 2:
            print_result({"success": False, "error": "Usage: xlsx-sheet-add <file> <name> [--position <n>]"}, json_output)
        else:
            position = None
            if len(args.args) > 3 and args.args[2] == "--position":
                position = int(args.args[3])
            result = doc_server.sheet_add(args.args[0], args.args[1], position=position)
            print_result(result, json_output)

    elif args.command == "xlsx-sheet-delete":
        if len(args.args) < 2:
            print_result({"success": False, "error": "Usage: xlsx-sheet-delete <file> <name>"}, json_output)
        else:
            result = doc_server.sheet_delete(args.args[0], args.args[1])
            print_result(result, json_output)

    elif args.command == "xlsx-sheet-rename":
        if len(args.args) < 3:
            print_result({"success": False, "error": "Usage: xlsx-sheet-rename <file> <old_name> <new_name>"}, json_output)
        else:
            result = doc_server.sheet_rename(args.args[0], args.args[1], args.args[2])
            print_result(result, json_output)

    elif args.command == "xlsx-merge-cells":
        if len(args.args) < 2:
            print_result({"success": False, "error": "Usage: xlsx-merge-cells <file> <range> [--sheet <name>]"}, json_output)
        else:
            sheet = None
            if len(args.args) > 3 and args.args[2] == "--sheet":
                sheet = args.args[3]
            result = doc_server.merge_cells(args.args[0], args.args[1], sheet_name=sheet)
            print_result(result, json_output)

    elif args.command == "xlsx-unmerge-cells":
        if len(args.args) < 2:
            print_result({"success": False, "error": "Usage: xlsx-unmerge-cells <file> <range> [--sheet <name>]"}, json_output)
        else:
            sheet = None
            if len(args.args) > 3 and args.args[2] == "--sheet":
                sheet = args.args[3]
            result = doc_server.unmerge_cells(args.args[0], args.args[1], sheet_name=sheet)
            print_result(result, json_output)

    elif args.command == "xlsx-format-cells":
        if len(args.args) < 2:
            print_result({"success": False, "error": "Usage: xlsx-format-cells <file> <range> [--sheet <name>] [--bold] [--italic] [--font-name <n>] [--font-size <n>] [--color <hex>] [--bg-color <hex>] [--number-format <fmt>] [--align <l|c|r>] [--wrap]"}, json_output)
        else:
            sheet = None
            opts = {"bold": False, "italic": False, "font_name": None, "font_size": None,
                    "color": None, "bg_color": None, "number_format": None, "alignment": None, "wrap_text": False}
            i = 2
            while i < len(args.args):
                if args.args[i] == "--sheet" and i + 1 < len(args.args):
                    sheet = args.args[i + 1]; i += 2
                elif args.args[i] == "--bold":
                    opts["bold"] = True; i += 1
                elif args.args[i] == "--italic":
                    opts["italic"] = True; i += 1
                elif args.args[i] == "--wrap":
                    opts["wrap_text"] = True; i += 1
                elif args.args[i] == "--font-name" and i + 1 < len(args.args):
                    opts["font_name"] = args.args[i + 1]; i += 2
                elif args.args[i] == "--font-size" and i + 1 < len(args.args):
                    opts["font_size"] = int(args.args[i + 1]); i += 2
                elif args.args[i] == "--color" and i + 1 < len(args.args):
                    opts["color"] = args.args[i + 1]; i += 2
                elif args.args[i] == "--bg-color" and i + 1 < len(args.args):
                    opts["bg_color"] = args.args[i + 1]; i += 2
                elif args.args[i] == "--number-format" and i + 1 < len(args.args):
                    opts["number_format"] = args.args[i + 1]; i += 2
                elif args.args[i] == "--align" and i + 1 < len(args.args):
                    opts["alignment"] = args.args[i + 1]; i += 2
                else:
                    i += 1
            result = doc_server.format_cells(args.args[0], args.args[1], sheet_name=sheet, **opts)
            print_result(result, json_output)

    elif args.command == "xlsx-csv-import":
        if len(args.args) < 2:
            print_result({"success": False, "error": "Usage: xlsx-csv-import <xlsx_file> <csv_file> [--sheet <name>] [--delimiter <char>]"}, json_output)
        else:
            sheet = "Sheet1"
            delimiter = ","
            i = 2
            while i < len(args.args):
                if args.args[i] == "--sheet" and i + 1 < len(args.args):
                    sheet = args.args[i + 1]; i += 2
                elif args.args[i] == "--delimiter" and i + 1 < len(args.args):
                    delimiter = args.args[i + 1]; i += 2
                else:
                    i += 1
            result = doc_server.csv_import(args.args[0], args.args[1], sheet_name=sheet, delimiter=delimiter)
            print_result(result, json_output)

    elif args.command == "xlsx-csv-export":
        if len(args.args) < 2:
            print_result({"success": False, "error": "Usage: xlsx-csv-export <xlsx_file> <csv_file> [--sheet <name>] [--delimiter <char>]"}, json_output)
        else:
            sheet = None
            delimiter = ","
            i = 2
            while i < len(args.args):
                if args.args[i] == "--sheet" and i + 1 < len(args.args):
                    sheet = args.args[i + 1]; i += 2
                elif args.args[i] == "--delimiter" and i + 1 < len(args.args):
                    delimiter = args.args[i + 1]; i += 2
                else:
                    i += 1
            result = doc_server.csv_export(args.args[0], args.args[1], sheet_name=sheet, delimiter=delimiter)
            print_result(result, json_output)

    # ==================== SPREADSHEET DATA VALIDATION ====================

    elif args.command == "xlsx-add-validation":
        if len(args.args) < 3:
            print_result(
                {"success": False, "error": "Usage: xlsx-add-validation <file> <range> <type> [--operator <op>] [--formula1 <v>] [--formula2 <v>] [--sheet <name>] [--error <msg>] [--prompt <msg>] [--error-style stop|warning|information] [--allow-blank]"},
                json_output,
            )
        else:
            file_path = args.args[0]
            cell_range = args.args[1]
            vtype = args.args[2]
            operator = None
            formula1 = None
            formula2 = None
            sheet = None
            error_msg = None
            error_title = None
            prompt_msg = None
            prompt_title = None
            error_style = "stop"
            allow_blank = True
            i = 3
            while i < len(args.args):
                if args.args[i] == "--operator" and i + 1 < len(args.args):
                    operator = args.args[i + 1]; i += 2
                elif args.args[i] == "--formula1" and i + 1 < len(args.args):
                    formula1 = args.args[i + 1]; i += 2
                elif args.args[i] == "--formula2" and i + 1 < len(args.args):
                    formula2 = args.args[i + 1]; i += 2
                elif args.args[i] == "--sheet" and i + 1 < len(args.args):
                    sheet = args.args[i + 1]; i += 2
                elif args.args[i] == "--error" and i + 1 < len(args.args):
                    error_msg = args.args[i + 1]; i += 2
                elif args.args[i] == "--error-title" and i + 1 < len(args.args):
                    error_title = args.args[i + 1]; i += 2
                elif args.args[i] == "--prompt" and i + 1 < len(args.args):
                    prompt_msg = args.args[i + 1]; i += 2
                elif args.args[i] == "--prompt-title" and i + 1 < len(args.args):
                    prompt_title = args.args[i + 1]; i += 2
                elif args.args[i] == "--error-style" and i + 1 < len(args.args):
                    error_style = args.args[i + 1]; i += 2
                elif args.args[i] == "--no-blank":
                    allow_blank = False; i += 1
                else:
                    i += 1
            result = doc_server.add_validation(
                file_path, cell_range, vtype,
                operator=operator, formula1=formula1, formula2=formula2,
                allow_blank=allow_blank, sheet_name=sheet,
                error_message=error_msg, error_title=error_title,
                prompt_message=prompt_msg, prompt_title=prompt_title,
                error_style=error_style,
            )
            print_result(result, json_output)

    elif args.command == "xlsx-add-dropdown":
        if len(args.args) < 3:
            print_result(
                {"success": False, "error": "Usage: xlsx-add-dropdown <file> <range> <options_csv> [--sheet <name>] [--prompt <msg>] [--error <msg>]"},
                json_output,
            )
        else:
            sheet = None
            prompt = None
            error_msg = None
            i = 3
            while i < len(args.args):
                if args.args[i] == "--sheet" and i + 1 < len(args.args):
                    sheet = args.args[i + 1]; i += 2
                elif args.args[i] == "--prompt" and i + 1 < len(args.args):
                    prompt = args.args[i + 1]; i += 2
                elif args.args[i] == "--error" and i + 1 < len(args.args):
                    error_msg = args.args[i + 1]; i += 2
                else:
                    i += 1
            result = doc_server.add_dropdown(
                args.args[0], args.args[1], args.args[2],
                sheet_name=sheet, prompt=prompt, error_message=error_msg,
            )
            print_result(result, json_output)

    elif args.command == "xlsx-list-validations":
        if not args.args:
            print_result(
                {"success": False, "error": "Usage: xlsx-list-validations <file> [--sheet <name>]"},
                json_output,
            )
        else:
            sheet = None
            i = 1
            while i < len(args.args):
                if args.args[i] == "--sheet" and i + 1 < len(args.args):
                    sheet = args.args[i + 1]; i += 2
                else:
                    i += 1
            result = doc_server.list_validations(args.args[0], sheet_name=sheet)
            print_result(result, json_output)

    elif args.command == "xlsx-remove-validation":
        if not args.args:
            print_result(
                {"success": False, "error": "Usage: xlsx-remove-validation <file> [--range <range>] [--all] [--sheet <name>]"},
                json_output,
            )
        else:
            cell_range = None
            remove_all = False
            sheet = None
            i = 1
            while i < len(args.args):
                if args.args[i] == "--range" and i + 1 < len(args.args):
                    cell_range = args.args[i + 1]; i += 2
                elif args.args[i] == "--all":
                    remove_all = True; i += 1
                elif args.args[i] == "--sheet" and i + 1 < len(args.args):
                    sheet = args.args[i + 1]; i += 2
                else:
                    i += 1
            result = doc_server.remove_validation(args.args[0], cell_range=cell_range, sheet_name=sheet, remove_all=remove_all)
            print_result(result, json_output)

    elif args.command == "xlsx-validate-data":
        if not args.args:
            print_result(
                {"success": False, "error": "Usage: xlsx-validate-data <file> [--sheet <name>] [--max-rows <n>]"},
                json_output,
            )
        else:
            sheet = None
            max_rows = 1000
            i = 1
            while i < len(args.args):
                if args.args[i] == "--sheet" and i + 1 < len(args.args):
                    sheet = args.args[i + 1]; i += 2
                elif args.args[i] == "--max-rows" and i + 1 < len(args.args):
                    max_rows = int(args.args[i + 1]); i += 2
                else:
                    i += 1
            result = doc_server.validate_data(args.args[0], sheet_name=sheet, max_rows=max_rows)
            print_result(result, json_output)

    # Chart commands
    elif args.command == "chart-create":
        if len(args.args) < 5:
            print_result(
                {
                    "success": False,
                    "error": "Usage: chart-create <file> <type> <data_range> <cat_range> <title> [--sheet <name>] [--output-sheet <name>] [--x-label <text>] [--y-label <text>] [--labels] [--no-legend] [--legend-pos <pos>] [--colors <hex,hex>]",
                },
                json_output,
            )
        else:
            # Parse optional arguments
            sheet = None
            output_sheet = None
            x_label = None
            y_label = None
            show_labels = False
            show_legend = True
            legend_pos = "right"
            colors = None

            i = 5
            while i < len(args.args):
                if args.args[i] == "--sheet" and i + 1 < len(args.args):
                    sheet = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--output-sheet" and i + 1 < len(args.args):
                    output_sheet = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--x-label" and i + 1 < len(args.args):
                    x_label = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--y-label" and i + 1 < len(args.args):
                    y_label = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--labels":
                    show_labels = True
                    i += 1
                elif args.args[i] == "--no-legend":
                    show_legend = False
                    i += 1
                elif args.args[i] == "--legend-pos" and i + 1 < len(args.args):
                    legend_pos = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--colors" and i + 1 < len(args.args):
                    colors = args.args[i + 1]
                    i += 2
                else:
                    i += 1

            cmd_chart_create(
                args.args[0],
                args.args[1],
                args.args[2],
                args.args[3],
                args.args[4],
                sheet,
                output_sheet,
                x_label,
                y_label,
                show_labels,
                show_legend,
                legend_pos,
                colors,
                json_output,
            )

    elif args.command == "chart-comparison":
        if len(args.args) < 3:
            print_result(
                {
                    "success": False,
                    "error": "Usage: chart-comparison <file> <type> <title> [--sheet <name>] [--start-row <n>] [--start-col <n>] [--cats <n>] [--series <n>] [--cat-col <n>] [--value-cols <n,n,n>] [--output <cell>] [--labels] [--no-legend]",
                },
                json_output,
            )
        else:
            sheet = None
            start_row = None
            start_col = None
            num_cats = None
            num_series = None
            cat_col = None
            value_cols = None
            output_cell = "A10"
            show_labels = False
            show_legend = True

            i = 3
            while i < len(args.args):
                if args.args[i] == "--sheet" and i + 1 < len(args.args):
                    sheet = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--start-row" and i + 1 < len(args.args):
                    start_row = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--start-col" and i + 1 < len(args.args):
                    start_col = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--cats" and i + 1 < len(args.args):
                    num_cats = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--series" and i + 1 < len(args.args):
                    num_series = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--cat-col" and i + 1 < len(args.args):
                    cat_col = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--value-cols" and i + 1 < len(args.args):
                    value_cols = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--output" and i + 1 < len(args.args):
                    output_cell = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--labels":
                    show_labels = True
                    i += 1
                elif args.args[i] == "--no-legend":
                    show_legend = False
                    i += 1
                else:
                    i += 1

            cmd_chart_comparison(
                args.args[0],
                args.args[1],
                sheet,
                args.args[2],
                start_row,
                start_col,
                num_cats,
                num_series,
                cat_col,
                value_cols,
                output_cell,
                show_labels,
                show_legend,
                json_output,
            )

    elif args.command == "chart-grade-dist":
        if len(args.args) < 3:
            print_result(
                {
                    "success": False,
                    "error": "Usage: chart-grade-dist <file> <grade_col> <title> [--sheet <name>] [--output <cell>]",
                },
                json_output,
            )
        else:
            sheet = None
            output_cell = "F2"

            i = 3
            while i < len(args.args):
                if args.args[i] == "--sheet" and i + 1 < len(args.args):
                    sheet = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--output" and i + 1 < len(args.args):
                    output_cell = args.args[i + 1]
                    i += 2
                else:
                    i += 1

            cmd_chart_grade_dist(
                args.args[0],
                sheet,
                args.args[1],
                args.args[2],
                output_cell,
                json_output,
            )

    elif args.command == "chart-progress":
        if len(args.args) < 4:
            print_result(
                {
                    "success": False,
                    "error": "Usage: chart-progress <file> <student_col> <grade_col> <title> [--sheet <name>] [--output <cell>] [--labels]",
                },
                json_output,
            )
        else:
            sheet = None
            output_cell = "D2"
            show_labels = True

            i = 4
            while i < len(args.args):
                if args.args[i] == "--sheet" and i + 1 < len(args.args):
                    sheet = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--output" and i + 1 < len(args.args):
                    output_cell = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--labels":
                    show_labels = True
                    i += 1
                elif args.args[i] == "--no-labels":
                    show_labels = False
                    i += 1
                else:
                    i += 1

            cmd_chart_progress(
                args.args[0],
                args.args[1],
                args.args[2],
                sheet,
                args.args[3],
                output_cell,
                show_labels,
                json_output,
            )

    # Presentation commands
    elif args.command == "pptx-create":
        if len(args.args) < 2:
            print_result(
                {
                    "success": False,
                    "error": "Usage: pptx-create <file> <title> [subtitle]",
                },
                json_output,
            )
        else:
            cmd_pptx_create(
                args.args[0],
                args.args[1],
                args.args[2] if len(args.args) > 2 else "",
                json_output,
            )

    elif args.command == "pptx-add-slide":
        if len(args.args) < 2:
            print_result(
                {
                    "success": False,
                    "error": "Usage: pptx-add-slide <file> <title> [content] [layout]",
                },
                json_output,
            )
        else:
            cmd_pptx_add_slide(
                args.args[0],
                args.args[1],
                args.args[2] if len(args.args) > 2 else "",
                args.args[3] if len(args.args) > 3 else "content",
                json_output,
            )

    elif args.command == "pptx-add-bullets":
        if len(args.args) < 3:
            print_result(
                {
                    "success": False,
                    "error": "Usage: pptx-add-bullets <file> <title> <bullets>",
                },
                json_output,
            )
        else:
            cmd_pptx_add_bullets(args.args[0], args.args[1], args.args[2], json_output)

    elif args.command == "pptx-read":
        if not args.args:
            print_result(
                {"success": False, "error": "Usage: pptx-read <file>"}, json_output
            )
        else:
            cmd_pptx_read(args.args[0], json_output)

    elif args.command == "pptx-add-image":
        if len(args.args) < 3:
            print_result(
                {
                    "success": False,
                    "error": "Usage: pptx-add-image <file> <title> <image_path>",
                },
                json_output,
            )
        else:
            cmd_pptx_add_image(args.args[0], args.args[1], args.args[2], json_output)

    elif args.command == "pptx-add-table":
        if len(args.args) < 4:
            print_result(
                {
                    "success": False,
                    "error": "Usage: pptx-add-table <file> <title> <headers> <data> [--coerce-rows]",
                },
                json_output,
            )
        else:
            coerce_rows = "--coerce-rows" in args.args[4:]
            cmd_pptx_add_table(
                args.args[0],
                args.args[1],
                args.args[2],
                args.args[3],
                coerce_rows=coerce_rows,
                json_output=json_output,
            )

    # Extended presentation commands
    elif args.command == "pptx-delete-slide":
        if len(args.args) < 2:
            print_result({"success": False, "error": "Usage: pptx-delete-slide <file> <index>"}, json_output)
        else:
            result = doc_server.delete_slide(args.args[0], int(args.args[1]))
            print_result(result, json_output)

    elif args.command == "pptx-speaker-notes":
        if len(args.args) < 2:
            print_result({"success": False, "error": "Usage: pptx-speaker-notes <file> <slide_index> [notes_text]"}, json_output)
        else:
            notes = " ".join(args.args[2:]) if len(args.args) > 2 else None
            result = doc_server.speaker_notes(args.args[0], int(args.args[1]), notes_text=notes)
            print_result(result, json_output)

    elif args.command == "pptx-update-text":
        if len(args.args) < 2:
            print_result({"success": False, "error": "Usage: pptx-update-text <file> <slide_index> [--title <t>] [--body <b>]"}, json_output)
        else:
            title = None
            body = None
            i = 2
            while i < len(args.args):
                if args.args[i] == "--title" and i + 1 < len(args.args):
                    title = args.args[i + 1]; i += 2
                elif args.args[i] == "--body" and i + 1 < len(args.args):
                    body = args.args[i + 1]; i += 2
                else:
                    i += 1
            result = doc_server.update_slide_text(args.args[0], int(args.args[1]), title=title, body=body)
            print_result(result, json_output)

    elif args.command == "pptx-slide-count":
        if not args.args:
            print_result({"success": False, "error": "Usage: pptx-slide-count <file>"}, json_output)
        else:
            result = doc_server.slide_count(args.args[0])
            print_result(result, json_output)

    # RDF commands
    elif args.command == "rdf-create":
        if not args.args:
            print_result({"success": False, "error": "Usage: rdf-create <file> [--base <uri>] [--format turtle|xml|n3|json-ld] [--prefix <p>=<uri>]"}, json_output)
        else:
            base_uri = None
            fmt = "turtle"
            prefixes = {}
            i = 1
            while i < len(args.args):
                if args.args[i] == "--base" and i + 1 < len(args.args):
                    base_uri = args.args[i + 1]; i += 2
                elif args.args[i] == "--format" and i + 1 < len(args.args):
                    fmt = args.args[i + 1]; i += 2
                elif args.args[i] == "--prefix" and i + 1 < len(args.args):
                    parts = args.args[i + 1].split("=", 1)
                    if len(parts) == 2:
                        prefixes[parts[0]] = parts[1]
                    i += 2
                else:
                    i += 1
            result = doc_server.rdf_create(args.args[0], base_uri=base_uri, format=fmt, prefixes=prefixes)
            print_result(result, json_output)

    elif args.command == "rdf-read":
        if not args.args:
            print_result({"success": False, "error": "Usage: rdf-read <file> [--limit <n>]"}, json_output)
        else:
            limit = 100
            if len(args.args) > 2 and args.args[1] == "--limit":
                limit = int(args.args[2])
            result = doc_server.rdf_read(args.args[0], limit=limit)
            print_result(result, json_output)

    elif args.command == "rdf-add":
        if len(args.args) < 4:
            print_result({"success": False, "error": "Usage: rdf-add <file> <subject> <predicate> <object> [--type uri|literal|bnode] [--lang <l>] [--datatype <dt>] [--format <f>]"}, json_output)
        else:
            obj_type = "uri"
            lang = None
            datatype = None
            fmt = "turtle"
            i = 4
            while i < len(args.args):
                if args.args[i] == "--type" and i + 1 < len(args.args):
                    obj_type = args.args[i + 1]; i += 2
                elif args.args[i] == "--lang" and i + 1 < len(args.args):
                    lang = args.args[i + 1]; i += 2
                elif args.args[i] == "--datatype" and i + 1 < len(args.args):
                    datatype = args.args[i + 1]; i += 2
                elif args.args[i] == "--format" and i + 1 < len(args.args):
                    fmt = args.args[i + 1]; i += 2
                else:
                    i += 1
            result = doc_server.rdf_add(args.args[0], args.args[1], args.args[2], args.args[3],
                                         object_type=obj_type, lang=lang, datatype=datatype, format=fmt)
            print_result(result, json_output)

    elif args.command == "rdf-remove":
        if not args.args:
            print_result({"success": False, "error": "Usage: rdf-remove <file> [--subject <s>] [--predicate <p>] [--object <o>] [--type uri|literal|bnode] [--format <f>]"}, json_output)
        else:
            subject = None
            predicate = None
            object_val = None
            object_type = "uri"
            fmt = None
            i = 1
            while i < len(args.args):
                if args.args[i] == "--subject" and i + 1 < len(args.args):
                    subject = args.args[i + 1]; i += 2
                elif args.args[i] == "--predicate" and i + 1 < len(args.args):
                    predicate = args.args[i + 1]; i += 2
                elif args.args[i] == "--object" and i + 1 < len(args.args):
                    object_val = args.args[i + 1]; i += 2
                elif args.args[i] == "--type" and i + 1 < len(args.args):
                    object_type = args.args[i + 1]; i += 2
                elif args.args[i] == "--format" and i + 1 < len(args.args):
                    fmt = args.args[i + 1]; i += 2
                else:
                    i += 1
            result = doc_server.rdf_remove(args.args[0], subject=subject, predicate=predicate,
                                            object_val=object_val, object_type=object_type, format=fmt)
            print_result(result, json_output)

    elif args.command == "rdf-query":
        if len(args.args) < 2:
            print_result({"success": False, "error": "Usage: rdf-query <file> <sparql_query> [--limit <n>]"}, json_output)
        else:
            limit = 100
            i = 2
            while i < len(args.args):
                if args.args[i] == "--limit" and i + 1 < len(args.args):
                    limit = int(args.args[i + 1]); i += 2
                else:
                    i += 1
            result = doc_server.rdf_query(args.args[0], args.args[1], limit=limit)
            print_result(result, json_output)

    elif args.command == "rdf-export":
        if len(args.args) < 2:
            print_result({"success": False, "error": "Usage: rdf-export <file> <output_file> [--format turtle|xml|n3|nt|json-ld|trig]"}, json_output)
        else:
            fmt = "turtle"
            if len(args.args) > 3 and args.args[2] == "--format":
                fmt = args.args[3]
            result = doc_server.rdf_export(args.args[0], args.args[1], output_format=fmt)
            print_result(result, json_output)

    elif args.command == "rdf-merge":
        if len(args.args) < 2:
            print_result({"success": False, "error": "Usage: rdf-merge <file_a> <file_b> [--output <file>] [--format turtle]"}, json_output)
        else:
            output = None
            fmt = "turtle"
            i = 2
            while i < len(args.args):
                if args.args[i] == "--output" and i + 1 < len(args.args):
                    output = args.args[i + 1]; i += 2
                elif args.args[i] == "--format" and i + 1 < len(args.args):
                    fmt = args.args[i + 1]; i += 2
                else:
                    i += 1
            result = doc_server.rdf_merge(args.args[0], args.args[1], output_path=output, format=fmt)
            print_result(result, json_output)

    elif args.command == "rdf-stats":
        if not args.args:
            print_result({"success": False, "error": "Usage: rdf-stats <file>"}, json_output)
        else:
            result = doc_server.rdf_stats(args.args[0])
            print_result(result, json_output)

    elif args.command == "rdf-namespace":
        if not args.args:
            print_result({"success": False, "error": "Usage: rdf-namespace <file> [<prefix> <uri>] [--format <f>]"}, json_output)
        else:
            prefix = None
            uri = None
            fmt = "turtle"
            if len(args.args) >= 3 and not args.args[1].startswith("--"):
                prefix = args.args[1]
                uri = args.args[2]
            i = 1
            while i < len(args.args):
                if args.args[i] == "--format" and i + 1 < len(args.args):
                    fmt = args.args[i + 1]; i += 2
                else:
                    i += 1
            result = doc_server.rdf_namespace(args.args[0], prefix=prefix, uri=uri, format=fmt)
            print_result(result, json_output)

    elif args.command == "rdf-validate":
        if len(args.args) < 2:
            print_result({"success": False, "error": "Usage: rdf-validate <data_file> <shapes_file>"}, json_output)
        else:
            result = doc_server.rdf_validate(args.args[0], args.args[1])
            print_result(result, json_output)

    # General commands
    elif args.command == "list":
        cmd_list(json_output)

    elif args.command == "open":
        if not args.args:
            print_result(
                {"success": False, "error": "Usage: open <file> [gui|web]"}, json_output
            )
        else:
            mode = args.args[1] if len(args.args) > 1 else "gui"
            cmd_open(args.args[0], mode, json_output)

    elif args.command == "watch":
        if not args.args:
            print_result(
                {"success": False, "error": "Usage: watch <file> [gui|web]"},
                json_output,
            )
        else:
            mode = args.args[1] if len(args.args) > 1 else "gui"
            cmd_watch(args.args[0], mode, json_output)

    elif args.command == "info":
        if not args.args:
            print_result({"success": False, "error": "Usage: info <file>"}, json_output)
        else:
            cmd_info(args.args[0], json_output)

    elif args.command == "backup-list":
        if not args.args:
            print_result(
                {"success": False, "error": "Usage: backup-list <file> [--limit <n>]"},
                json_output,
            )
        else:
            file_path = None
            limit = 20
            i = 0
            while i < len(args.args):
                if args.args[i] == "--limit" and i + 1 < len(args.args):
                    limit = int(args.args[i + 1])
                    i += 2
                elif not args.args[i].startswith("--") and file_path is None:
                    file_path = args.args[i]
                    i += 1
                else:
                    i += 1
            if not file_path:
                print_result(
                    {
                        "success": False,
                        "error": "Usage: backup-list <file> [--limit <n>]",
                    },
                    json_output,
                )
            else:
                cmd_backup_list(file_path, limit=limit, json_output=json_output)

    elif args.command == "backup-prune":
        file_path = None
        keep = 20
        days = None
        i = 0
        while i < len(args.args):
            if args.args[i] == "--file" and i + 1 < len(args.args):
                file_path = args.args[i + 1]
                i += 2
            elif args.args[i] == "--keep" and i + 1 < len(args.args):
                keep = int(args.args[i + 1])
                i += 2
            elif args.args[i] == "--days" and i + 1 < len(args.args):
                days = int(args.args[i + 1])
                i += 2
            elif not args.args[i].startswith("--") and file_path is None:
                file_path = args.args[i]
                i += 1
            else:
                i += 1
        cmd_backup_prune(
            file_path=file_path,
            keep=keep,
            days=days,
            json_output=json_output,
        )

    elif args.command == "backup-restore":
        if not args.args:
            print_result(
                {
                    "success": False,
                    "error": "Usage: backup-restore <file> [--backup <name|path>] [--latest] [--dry-run]",
                },
                json_output,
            )
        else:
            file_path = args.args[0]
            backup = None
            latest = False
            dry_run = False
            i = 1
            while i < len(args.args):
                if args.args[i] == "--backup" and i + 1 < len(args.args):
                    backup = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--latest":
                    latest = True
                    i += 1
                elif args.args[i] == "--dry-run":
                    dry_run = True
                    i += 1
                else:
                    i += 1
            cmd_backup_restore(
                file_path=file_path,
                backup=backup,
                latest=latest,
                dry_run=dry_run,
                json_output=json_output,
            )

    # ==================== IMAGE EXTRACTION ====================

    elif args.command == "doc-extract-images":
        if len(args.args) < 2:
            print_result(
                {"success": False, "error": "Usage: doc-extract-images <file> <output_dir> [--format png|jpg] [--prefix <name>]"},
                json_output,
            )
        else:
            fmt = "png"
            prefix = "image"
            i = 2
            while i < len(args.args):
                if args.args[i] == "--format" and i + 1 < len(args.args):
                    fmt = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--prefix" and i + 1 < len(args.args):
                    prefix = args.args[i + 1]
                    i += 2
                else:
                    i += 1
            result = doc_server.extract_images_from_docx(args.args[0], args.args[1], fmt=fmt, prefix=prefix)
            print_result(result, json_output)

    elif args.command == "doc-to-pdf":
        if len(args.args) < 1:
            print_result(
                {"success": False, "error": "Usage: doc-to-pdf <file> [output_path]"},
                json_output,
            )
        else:
            out_path = args.args[1] if len(args.args) >= 2 else None
            result = doc_server.doc_to_pdf(args.args[0], output_path=out_path)
            print_result(result, json_output)

    elif args.command == "pptx-extract-images":
        if len(args.args) < 2:
            print_result(
                {"success": False, "error": "Usage: pptx-extract-images <file> <output_dir> [--slide <index>] [--format png|jpg]"},
                json_output,
            )
        else:
            fmt = "png"
            slide_idx = None
            prefix = "slide"
            i = 2
            while i < len(args.args):
                if args.args[i] == "--format" and i + 1 < len(args.args):
                    fmt = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--slide" and i + 1 < len(args.args):
                    slide_idx = int(args.args[i + 1])
                    i += 2
                elif args.args[i] == "--prefix" and i + 1 < len(args.args):
                    prefix = args.args[i + 1]
                    i += 2
                else:
                    i += 1
            result = doc_server.extract_images_from_pptx(args.args[0], args.args[1], slide_index=slide_idx, fmt=fmt, prefix=prefix)
            print_result(result, json_output)

    elif args.command == "pdf-extract-images":
        if len(args.args) < 2:
            print_result(
                {"success": False, "error": "Usage: pdf-extract-images <file> <output_dir> [--format png|jpg] [--pages <range>]"},
                json_output,
            )
        else:
            fmt = "png"
            pages = None
            i = 2
            while i < len(args.args):
                if args.args[i] == "--format" and i + 1 < len(args.args):
                    fmt = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--pages" and i + 1 < len(args.args):
                    pages = args.args[i + 1]
                    i += 2
                else:
                    i += 1
            result = doc_server.pdf_extract_images(args.args[0], args.args[1], fmt=fmt, pages=pages)
            print_result(result, json_output)

    elif args.command == "pdf-page-to-image":
        if len(args.args) < 2:
            print_result(
                {"success": False, "error": "Usage: pdf-page-to-image <file> <output_dir> [--pages <range>] [--dpi <n>] [--format png|jpg]"},
                json_output,
            )
        else:
            fmt = "png"
            pages = None
            dpi = 150
            i = 2
            while i < len(args.args):
                if args.args[i] == "--format" and i + 1 < len(args.args):
                    fmt = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--pages" and i + 1 < len(args.args):
                    pages = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--dpi" and i + 1 < len(args.args):
                    dpi = int(args.args[i + 1])
                    i += 2
                else:
                    i += 1
            result = doc_server.pdf_page_to_image(args.args[0], args.args[1], pages=pages, dpi=dpi, fmt=fmt)
            print_result(result, json_output)

    # ==================== PPTX SPATIAL / TEXTBOX ====================

    elif args.command == "pptx-list-shapes":
        if not args.args:
            print_result(
                {"success": False, "error": "Usage: pptx-list-shapes <file> [--slide <index>]"},
                json_output,
            )
        else:
            slide_idx = None
            i = 1
            while i < len(args.args):
                if args.args[i] == "--slide" and i + 1 < len(args.args):
                    slide_idx = int(args.args[i + 1])
                    i += 2
                else:
                    i += 1
            result = doc_server.list_shapes(args.args[0], slide_index=slide_idx)
            print_result(result, json_output)

    elif args.command == "pptx-add-textbox":
        if len(args.args) < 3:
            print_result(
                {"success": False, "error": "Usage: pptx-add-textbox <file> <slide_index> <text> [--left <in>] [--top <in>] [--width <in>] [--height <in>] [--font-size <pt>] [--font-name <name>] [--bold] [--italic] [--color <hex>] [--align <left|center|right>]"},
                json_output,
            )
        else:
            file_path = args.args[0]
            slide_idx = int(args.args[1])
            text = args.args[2]
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
            i = 3
            while i < len(args.args):
                if args.args[i] == "--left" and i + 1 < len(args.args):
                    left = float(args.args[i + 1])
                    i += 2
                elif args.args[i] == "--top" and i + 1 < len(args.args):
                    top = float(args.args[i + 1])
                    i += 2
                elif args.args[i] == "--width" and i + 1 < len(args.args):
                    width = float(args.args[i + 1])
                    i += 2
                elif args.args[i] == "--height" and i + 1 < len(args.args):
                    height = float(args.args[i + 1])
                    i += 2
                elif args.args[i] == "--font-size" and i + 1 < len(args.args):
                    font_size = float(args.args[i + 1])
                    i += 2
                elif args.args[i] == "--font-name" and i + 1 < len(args.args):
                    font_name = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--bold":
                    bold = True
                    i += 1
                elif args.args[i] == "--italic":
                    italic = True
                    i += 1
                elif args.args[i] == "--color" and i + 1 < len(args.args):
                    color = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--align" and i + 1 < len(args.args):
                    align = args.args[i + 1]
                    i += 2
                else:
                    i += 1
            result = doc_server.add_textbox(
                file_path, slide_idx, text,
                left=left, top=top, width=width, height=height,
                font_size=font_size, font_name=font_name,
                bold=bold, italic=italic, color=color, align=align,
            )
            print_result(result, json_output)

    elif args.command == "pptx-modify-shape":
        if len(args.args) < 3:
            print_result(
                {"success": False, "error": "Usage: pptx-modify-shape <file> <slide_index> <shape_name> [--left <in>] [--top <in>] [--width <in>] [--height <in>] [--text <text>] [--font-size <pt>] [--rotation <deg>]"},
                json_output,
            )
        else:
            file_path = args.args[0]
            slide_idx = int(args.args[1])
            shape_name = args.args[2]
            left = None
            top = None
            width = None
            height = None
            text = None
            font_size = None
            rotation = None
            i = 3
            while i < len(args.args):
                if args.args[i] == "--left" and i + 1 < len(args.args):
                    left = float(args.args[i + 1])
                    i += 2
                elif args.args[i] == "--top" and i + 1 < len(args.args):
                    top = float(args.args[i + 1])
                    i += 2
                elif args.args[i] == "--width" and i + 1 < len(args.args):
                    width = float(args.args[i + 1])
                    i += 2
                elif args.args[i] == "--height" and i + 1 < len(args.args):
                    height = float(args.args[i + 1])
                    i += 2
                elif args.args[i] == "--text" and i + 1 < len(args.args):
                    text = args.args[i + 1]
                    i += 2
                elif args.args[i] == "--font-size" and i + 1 < len(args.args):
                    font_size = float(args.args[i + 1])
                    i += 2
                elif args.args[i] == "--rotation" and i + 1 < len(args.args):
                    rotation = float(args.args[i + 1])
                    i += 2
                else:
                    i += 1
            result = doc_server.modify_shape(
                file_path, slide_idx, shape_name,
                left=left, top=top, width=width, height=height,
                text=text, font_size=font_size, rotation=rotation,
            )
            print_result(result, json_output)

    elif args.command == "pptx-preview":
        if len(args.args) < 2:
            print_result(
                {"success": False, "error": "Usage: pptx-preview <file> <output_dir> [--slide <index>] [--dpi <n>]"},
                json_output,
            )
        else:
            slide_idx = None
            dpi = 150
            i = 2
            while i < len(args.args):
                if args.args[i] == "--slide" and i + 1 < len(args.args):
                    slide_idx = int(args.args[i + 1])
                    i += 2
                elif args.args[i] == "--dpi" and i + 1 < len(args.args):
                    dpi = int(args.args[i + 1])
                    i += 2
                else:
                    i += 1
            result = doc_server.preview_slide(args.args[0], args.args[1], slide_index=slide_idx, dpi=dpi)
            print_result(result, json_output)

    # ==================== GENERAL ====================

    elif args.command == "status":
        cmd_status(json_output)

    elif args.command in ["help", "--help", "-h"]:
        cmd_help(json_output)

    else:
        print_result(
            {"success": False, "error": f"Unknown command: {args.command}"}, json_output
        )


if __name__ == "__main__":
    main()
