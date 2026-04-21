#!/usr/bin/env python3
"""XLSX and chart CLI command handlers."""

from __future__ import annotations

from typing import Any, Callable, List, Optional

from cli_anything.onlyoffice.core.command_registry import command_usage


def _split_csv(value: str) -> List[str]:
    return [item.strip() for item in value.split(",") if item.strip()]


def handle_xlsx_command(
    command: str,
    raw_args: List[str],
    doc_server: Any,
    json_output: bool,
    print_result: Callable[[dict, bool], None],
) -> bool:
    """Handle XLSX and chart commands and return True when recognised."""
    if command == "xlsx-create":
        if not raw_args:
            print_result(
                {"success": False, "error": command_usage("xlsx-create")},
                json_output,
            )
            return True
        print_result(
            doc_server.create_spreadsheet(
                raw_args[0],
                raw_args[1] if len(raw_args) > 1 else "Sheet1",
            ),
            json_output,
        )
        return True

    if command == "xlsx-write":
        if len(raw_args) < 3:
            print_result(
                {
                    "success": False,
                    "error": command_usage("xlsx-write"),
                },
                json_output,
            )
            return True
        sheet = "Sheet1"
        overwrite = False
        coerce_rows = False
        text_columns = None
        index = 3
        while index < len(raw_args):
            if raw_args[index] == "--sheet" and index + 1 < len(raw_args):
                sheet = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--overwrite":
                overwrite = True
                index += 1
            elif raw_args[index] == "--coerce-rows":
                coerce_rows = True
                index += 1
            elif raw_args[index] == "--text-columns" and index + 1 < len(raw_args):
                text_columns = _split_csv(raw_args[index + 1])
                index += 2
            else:
                index += 1

        headers = [header.strip() for header in raw_args[1].split(",")]
        rows = []
        for row_str in raw_args[2].split(";"):
            if row_str.strip():
                rows.append([cell.strip() for cell in row_str.split(",")])

        print_result(
            doc_server.write_spreadsheet(
                raw_args[0],
                headers,
                rows,
                sheet_name=sheet,
                overwrite_workbook=overwrite,
                coerce_rows=coerce_rows,
                text_columns=text_columns,
            ),
            json_output,
        )
        return True

    if command == "xlsx-read":
        if not raw_args:
            print_result(
                {"success": False, "error": command_usage("xlsx-read")},
                json_output,
            )
            return True
        print_result(
            doc_server.read_spreadsheet(raw_args[0], raw_args[1] if len(raw_args) > 1 else None),
            json_output,
        )
        return True

    if command == "xlsx-append":
        if len(raw_args) < 2:
            print_result(
                {
                    "success": False,
                    "error": command_usage("xlsx-append"),
                },
                json_output,
            )
            return True
        sheet = None
        if len(raw_args) > 3 and raw_args[2] == "--sheet":
            sheet = raw_args[3]
        row_data = [cell.strip() for cell in raw_args[1].split(",")]
        print_result(
            doc_server.append_to_spreadsheet(raw_args[0], row_data, sheet_name=sheet),
            json_output,
        )
        return True

    if command == "xlsx-search":
        if len(raw_args) < 2:
            print_result(
                {
                    "success": False,
                    "error": command_usage("xlsx-search"),
                },
                json_output,
            )
            return True
        sheet = None
        if len(raw_args) > 3 and raw_args[2] == "--sheet":
            sheet = raw_args[3]
        print_result(
            doc_server.search_spreadsheet(raw_args[0], raw_args[1], sheet_name=sheet),
            json_output,
        )
        return True

    if command == "xlsx-calc":
        if len(raw_args) < 3:
            print_result(
                {
                    "success": False,
                    "error": command_usage("xlsx-calc"),
                },
                json_output,
            )
            return True
        sheet = "Sheet1"
        include_formulas = False
        strict_formulas = False
        index = 3
        while index < len(raw_args):
            if raw_args[index] == "--sheet" and index + 1 < len(raw_args):
                sheet = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--include-formulas":
                include_formulas = True
                index += 1
            elif raw_args[index] == "--strict-formulas":
                strict_formulas = True
                index += 1
            else:
                index += 1
        print_result(
            doc_server.calculate_column(
                raw_args[0],
                raw_args[1],
                raw_args[2],
                sheet_name=sheet,
                include_formulas=include_formulas,
                strict_formula_safety=strict_formulas,
            ),
            json_output,
        )
        return True

    if command == "xlsx-formula":
        if len(raw_args) < 3:
            print_result(
                {
                    "success": False,
                    "error": command_usage("xlsx-formula"),
                },
                json_output,
            )
            return True
        sheet = "Sheet1"
        if len(raw_args) > 4 and raw_args[3] == "--sheet":
            sheet = raw_args[4]
        print_result(
            doc_server.add_formula(raw_args[0], raw_args[1], raw_args[2], sheet_name=sheet),
            json_output,
        )
        return True

    if command == "xlsx-freq":
        if len(raw_args) < 2:
            print_result(
                {
                    "success": False,
                    "error": command_usage("xlsx-freq"),
                },
                json_output,
            )
            return True
        sheet = "Sheet1"
        valid_values = None
        index = 2
        while index < len(raw_args):
            if raw_args[index] == "--sheet" and index + 1 < len(raw_args):
                sheet = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--valid" and index + 1 < len(raw_args):
                valid_values = _split_csv(raw_args[index + 1])
                index += 2
            else:
                index += 1
        print_result(
            doc_server.frequencies(
                raw_args[0],
                raw_args[1],
                sheet_name=sheet,
                allowed_values=valid_values,
            ),
            json_output,
        )
        return True

    if command == "xlsx-corr":
        if len(raw_args) < 3:
            print_result(
                {
                    "success": False,
                    "error": command_usage("xlsx-corr"),
                },
                json_output,
            )
            return True
        sheet = "Sheet1"
        method = "pearson"
        index = 3
        while index < len(raw_args):
            if raw_args[index] == "--sheet" and index + 1 < len(raw_args):
                sheet = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--method" and index + 1 < len(raw_args):
                method = raw_args[index + 1]
                index += 2
            else:
                index += 1
        print_result(
            doc_server.correlation_test(
                file_path=raw_args[0],
                x_column=raw_args[1],
                y_column=raw_args[2],
                sheet_name=sheet,
                method=method,
            ),
            json_output,
        )
        return True

    if command == "xlsx-ttest":
        if len(raw_args) < 5:
            print_result(
                {
                    "success": False,
                    "error": command_usage("xlsx-ttest"),
                },
                json_output,
            )
            return True
        sheet = "Sheet1"
        equal_var = False
        index = 5
        while index < len(raw_args):
            if raw_args[index] == "--sheet" and index + 1 < len(raw_args):
                sheet = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--equal-var":
                equal_var = True
                index += 1
            else:
                index += 1
        print_result(
            doc_server.ttest_independent(
                file_path=raw_args[0],
                value_column=raw_args[1],
                group_column=raw_args[2],
                group_a=raw_args[3],
                group_b=raw_args[4],
                sheet_name=sheet,
                equal_var=equal_var,
            ),
            json_output,
        )
        return True

    if command == "xlsx-mannwhitney":
        if len(raw_args) < 5:
            print_result(
                {
                    "success": False,
                    "error": command_usage("xlsx-mannwhitney"),
                },
                json_output,
            )
            return True
        sheet = "Sheet1"
        index = 5
        while index < len(raw_args):
            if raw_args[index] == "--sheet" and index + 1 < len(raw_args):
                sheet = raw_args[index + 1]
                index += 2
            else:
                index += 1
        print_result(
            doc_server.mann_whitney_test(
                raw_args[0],
                raw_args[1],
                raw_args[2],
                raw_args[3],
                raw_args[4],
                sheet_name=sheet,
            ),
            json_output,
        )
        return True

    if command == "xlsx-chi2":
        if len(raw_args) < 3:
            print_result(
                {
                    "success": False,
                    "error": command_usage("xlsx-chi2"),
                },
                json_output,
            )
            return True
        sheet = "Sheet1"
        row_valid_values = None
        col_valid_values = None
        index = 3
        while index < len(raw_args):
            if raw_args[index] == "--sheet" and index + 1 < len(raw_args):
                sheet = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--row-valid" and index + 1 < len(raw_args):
                row_valid_values = _split_csv(raw_args[index + 1])
                index += 2
            elif raw_args[index] == "--col-valid" and index + 1 < len(raw_args):
                col_valid_values = _split_csv(raw_args[index + 1])
                index += 2
            else:
                index += 1
        print_result(
            doc_server.chi_square_test(
                file_path=raw_args[0],
                row_column=raw_args[1],
                col_column=raw_args[2],
                sheet_name=sheet,
                row_allowed_values=row_valid_values,
                col_allowed_values=col_valid_values,
            ),
            json_output,
        )
        return True

    if command == "xlsx-research-pack":
        if len(raw_args) < 1:
            print_result(
                {
                    "success": False,
                    "error": command_usage("xlsx-research-pack"),
                },
                json_output,
            )
            return True
        sheet = "Sheet0"
        profile = "hlth3112"
        require_formula_safe = False
        index = 1
        while index < len(raw_args):
            if raw_args[index] == "--sheet" and index + 1 < len(raw_args):
                sheet = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--profile" and index + 1 < len(raw_args):
                profile = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--require-formula-safe":
                require_formula_safe = True
                index += 1
            else:
                index += 1
        print_result(
            doc_server.research_analysis_pack(
                file_path=raw_args[0],
                sheet_name=sheet,
                profile=profile,
                require_formula_safe=require_formula_safe,
            ),
            json_output,
        )
        return True

    if command == "xlsx-formula-audit":
        if len(raw_args) < 1:
            print_result(
                {
                    "success": False,
                    "error": command_usage("xlsx-formula-audit"),
                },
                json_output,
            )
            return True
        sheet = None
        index = 1
        while index < len(raw_args):
            if raw_args[index] == "--sheet" and index + 1 < len(raw_args):
                sheet = raw_args[index + 1]
                index += 2
            else:
                index += 1
        print_result(
            doc_server.audit_spreadsheet_formulas(raw_args[0], sheet_name=sheet),
            json_output,
        )
        return True

    if command == "xlsx-text-extract":
        if len(raw_args) < 2:
            print_result(
                {
                    "success": False,
                    "error": command_usage("xlsx-text-extract"),
                },
                json_output,
            )
            return True
        sheet = "Sheet0"
        limit = 20
        min_length = 20
        index = 2
        while index < len(raw_args):
            if raw_args[index] == "--sheet" and index + 1 < len(raw_args):
                sheet = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--limit" and index + 1 < len(raw_args):
                limit = int(raw_args[index + 1])
                index += 2
            elif raw_args[index] == "--min-length" and index + 1 < len(raw_args):
                min_length = int(raw_args[index + 1])
                index += 2
            else:
                index += 1
        print_result(
            doc_server.open_text_extract(
                file_path=raw_args[0],
                column_letter=raw_args[1],
                sheet_name=sheet,
                limit=limit,
                min_length=min_length,
            ),
            json_output,
        )
        return True

    if command == "xlsx-text-keywords":
        if len(raw_args) < 2:
            print_result(
                {
                    "success": False,
                    "error": command_usage("xlsx-text-keywords"),
                },
                json_output,
            )
            return True
        sheet = "Sheet0"
        top = 15
        min_word_length = 4
        index = 2
        while index < len(raw_args):
            if raw_args[index] == "--sheet" and index + 1 < len(raw_args):
                sheet = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--top" and index + 1 < len(raw_args):
                top = int(raw_args[index + 1])
                index += 2
            elif raw_args[index] == "--min-word-length" and index + 1 < len(raw_args):
                min_word_length = int(raw_args[index + 1])
                index += 2
            else:
                index += 1
        print_result(
            doc_server.open_text_keywords(
                file_path=raw_args[0],
                column_letter=raw_args[1],
                sheet_name=sheet,
                top_n=top,
                min_word_length=min_word_length,
            ),
            json_output,
        )
        return True

    if command == "xlsx-cell-read":
        if len(raw_args) < 2:
            print_result(
                {"success": False, "error": command_usage("xlsx-cell-read")},
                json_output,
            )
            return True
        sheet = None
        if len(raw_args) > 3 and raw_args[2] == "--sheet":
            sheet = raw_args[3]
        print_result(doc_server.cell_read(raw_args[0], raw_args[1], sheet_name=sheet), json_output)
        return True

    if command == "xlsx-cell-write":
        if len(raw_args) < 3:
            print_result(
                {"success": False, "error": command_usage("xlsx-cell-write")},
                json_output,
            )
            return True
        sheet = None
        as_text = False
        index = 3
        while index < len(raw_args):
            if raw_args[index] == "--sheet" and index + 1 < len(raw_args):
                sheet = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--text":
                as_text = True
                index += 1
            else:
                index += 1
        print_result(
            doc_server.cell_write(
                raw_args[0],
                raw_args[1],
                raw_args[2],
                sheet_name=sheet,
                as_text=as_text,
            ),
            json_output,
        )
        return True

    if command == "xlsx-range-read":
        if len(raw_args) < 2:
            print_result(
                {"success": False, "error": command_usage("xlsx-range-read")},
                json_output,
            )
            return True
        sheet = None
        if len(raw_args) > 3 and raw_args[2] == "--sheet":
            sheet = raw_args[3]
        print_result(doc_server.range_read(raw_args[0], raw_args[1], sheet_name=sheet), json_output)
        return True

    if command == "xlsx-delete-rows":
        if len(raw_args) < 2:
            print_result(
                {"success": False, "error": command_usage("xlsx-delete-rows")},
                json_output,
            )
            return True
        count = 1
        sheet = None
        index = 2
        while index < len(raw_args):
            if raw_args[index] == "--sheet" and index + 1 < len(raw_args):
                sheet = raw_args[index + 1]
                index += 2
            elif not raw_args[index].startswith("--"):
                count = int(raw_args[index])
                index += 1
            else:
                index += 1
        print_result(
            doc_server.delete_rows(raw_args[0], int(raw_args[1]), count=count, sheet_name=sheet),
            json_output,
        )
        return True

    if command == "xlsx-delete-cols":
        if len(raw_args) < 2:
            print_result(
                {"success": False, "error": command_usage("xlsx-delete-cols")},
                json_output,
            )
            return True
        count = 1
        sheet = None
        index = 2
        while index < len(raw_args):
            if raw_args[index] == "--sheet" and index + 1 < len(raw_args):
                sheet = raw_args[index + 1]
                index += 2
            elif not raw_args[index].startswith("--"):
                count = int(raw_args[index])
                index += 1
            else:
                index += 1
        print_result(
            doc_server.delete_columns(raw_args[0], int(raw_args[1]), count=count, sheet_name=sheet),
            json_output,
        )
        return True

    if command == "xlsx-sort":
        if len(raw_args) < 2:
            print_result(
                {"success": False, "error": command_usage("xlsx-sort")},
                json_output,
            )
            return True
        sheet = None
        descending = False
        numeric = False
        index = 2
        while index < len(raw_args):
            if raw_args[index] == "--sheet" and index + 1 < len(raw_args):
                sheet = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--desc":
                descending = True
                index += 1
            elif raw_args[index] == "--numeric":
                numeric = True
                index += 1
            else:
                index += 1
        print_result(
            doc_server.sort_sheet(
                raw_args[0],
                raw_args[1],
                sheet_name=sheet,
                descending=descending,
                numeric=numeric,
            ),
            json_output,
        )
        return True

    if command == "xlsx-filter":
        if len(raw_args) < 4:
            print_result(
                {
                    "success": False,
                    "error": command_usage("xlsx-filter"),
                },
                json_output,
            )
            return True
        sheet = None
        index = 4
        while index < len(raw_args):
            if raw_args[index] == "--sheet" and index + 1 < len(raw_args):
                sheet = raw_args[index + 1]
                index += 2
            else:
                index += 1
        print_result(
            doc_server.filter_rows(
                raw_args[0],
                raw_args[1],
                raw_args[2],
                raw_args[3],
                sheet_name=sheet,
            ),
            json_output,
        )
        return True

    if command == "xlsx-sheet-list":
        if not raw_args:
            print_result(
                {"success": False, "error": command_usage("xlsx-sheet-list")},
                json_output,
            )
            return True
        print_result(doc_server.sheet_list(raw_args[0]), json_output)
        return True

    if command == "xlsx-sheet-add":
        if len(raw_args) < 2:
            print_result(
                {"success": False, "error": command_usage("xlsx-sheet-add")},
                json_output,
            )
            return True
        position = None
        if len(raw_args) > 3 and raw_args[2] == "--position":
            position = int(raw_args[3])
        print_result(doc_server.sheet_add(raw_args[0], raw_args[1], position=position), json_output)
        return True

    if command == "xlsx-sheet-delete":
        if len(raw_args) < 2:
            print_result(
                {"success": False, "error": command_usage("xlsx-sheet-delete")},
                json_output,
            )
            return True
        print_result(doc_server.sheet_delete(raw_args[0], raw_args[1]), json_output)
        return True

    if command == "xlsx-sheet-rename":
        if len(raw_args) < 3:
            print_result(
                {"success": False, "error": command_usage("xlsx-sheet-rename")},
                json_output,
            )
            return True
        print_result(doc_server.sheet_rename(raw_args[0], raw_args[1], raw_args[2]), json_output)
        return True

    if command == "xlsx-merge-cells":
        if len(raw_args) < 2:
            print_result(
                {"success": False, "error": command_usage("xlsx-merge-cells")},
                json_output,
            )
            return True
        sheet = None
        if len(raw_args) > 3 and raw_args[2] == "--sheet":
            sheet = raw_args[3]
        print_result(doc_server.merge_cells(raw_args[0], raw_args[1], sheet_name=sheet), json_output)
        return True

    if command == "xlsx-unmerge-cells":
        if len(raw_args) < 2:
            print_result(
                {"success": False, "error": command_usage("xlsx-unmerge-cells")},
                json_output,
            )
            return True
        sheet = None
        if len(raw_args) > 3 and raw_args[2] == "--sheet":
            sheet = raw_args[3]
        print_result(doc_server.unmerge_cells(raw_args[0], raw_args[1], sheet_name=sheet), json_output)
        return True

    if command == "xlsx-format-cells":
        if len(raw_args) < 2:
            print_result(
                {
                    "success": False,
                    "error": command_usage("xlsx-format-cells"),
                },
                json_output,
            )
            return True
        sheet = None
        options = {
            "bold": False,
            "italic": False,
            "font_name": None,
            "font_size": None,
            "color": None,
            "bg_color": None,
            "number_format": None,
            "alignment": None,
            "wrap_text": False,
        }
        index = 2
        while index < len(raw_args):
            if raw_args[index] == "--sheet" and index + 1 < len(raw_args):
                sheet = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--bold":
                options["bold"] = True
                index += 1
            elif raw_args[index] == "--italic":
                options["italic"] = True
                index += 1
            elif raw_args[index] == "--wrap":
                options["wrap_text"] = True
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
            elif raw_args[index] == "--bg-color" and index + 1 < len(raw_args):
                options["bg_color"] = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--number-format" and index + 1 < len(raw_args):
                options["number_format"] = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--align" and index + 1 < len(raw_args):
                options["alignment"] = raw_args[index + 1]
                index += 2
            else:
                index += 1
        print_result(
            doc_server.format_cells(raw_args[0], raw_args[1], sheet_name=sheet, **options),
            json_output,
        )
        return True

    if command == "xlsx-csv-import":
        if len(raw_args) < 2:
            print_result(
                {"success": False, "error": command_usage("xlsx-csv-import")},
                json_output,
            )
            return True
        sheet = "Sheet1"
        delimiter = ","
        index = 2
        while index < len(raw_args):
            if raw_args[index] == "--sheet" and index + 1 < len(raw_args):
                sheet = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--delimiter" and index + 1 < len(raw_args):
                delimiter = raw_args[index + 1]
                index += 2
            else:
                index += 1
        print_result(
            doc_server.csv_import(raw_args[0], raw_args[1], sheet_name=sheet, delimiter=delimiter),
            json_output,
        )
        return True

    if command == "xlsx-csv-export":
        if len(raw_args) < 2:
            print_result(
                {"success": False, "error": command_usage("xlsx-csv-export")},
                json_output,
            )
            return True
        sheet = None
        delimiter = ","
        index = 2
        while index < len(raw_args):
            if raw_args[index] == "--sheet" and index + 1 < len(raw_args):
                sheet = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--delimiter" and index + 1 < len(raw_args):
                delimiter = raw_args[index + 1]
                index += 2
            else:
                index += 1
        print_result(
            doc_server.csv_export(raw_args[0], raw_args[1], sheet_name=sheet, delimiter=delimiter),
            json_output,
        )
        return True

    if command == "xlsx-add-validation":
        if len(raw_args) < 3:
            print_result(
                {
                    "success": False,
                    "error": command_usage("xlsx-add-validation"),
                },
                json_output,
            )
            return True
        file_path = raw_args[0]
        cell_range = raw_args[1]
        validation_type = raw_args[2]
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
        index = 3
        while index < len(raw_args):
            if raw_args[index] == "--operator" and index + 1 < len(raw_args):
                operator = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--formula1" and index + 1 < len(raw_args):
                formula1 = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--formula2" and index + 1 < len(raw_args):
                formula2 = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--sheet" and index + 1 < len(raw_args):
                sheet = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--error" and index + 1 < len(raw_args):
                error_msg = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--error-title" and index + 1 < len(raw_args):
                error_title = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--prompt" and index + 1 < len(raw_args):
                prompt_msg = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--prompt-title" and index + 1 < len(raw_args):
                prompt_title = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--error-style" and index + 1 < len(raw_args):
                error_style = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--no-blank":
                allow_blank = False
                index += 1
            else:
                index += 1
        print_result(
            doc_server.add_validation(
                file_path,
                cell_range,
                validation_type,
                operator=operator,
                formula1=formula1,
                formula2=formula2,
                allow_blank=allow_blank,
                sheet_name=sheet,
                error_message=error_msg,
                error_title=error_title,
                prompt_message=prompt_msg,
                prompt_title=prompt_title,
                error_style=error_style,
            ),
            json_output,
        )
        return True

    if command == "xlsx-add-dropdown":
        if len(raw_args) < 3:
            print_result(
                {
                    "success": False,
                    "error": command_usage("xlsx-add-dropdown"),
                },
                json_output,
            )
            return True
        sheet = None
        prompt = None
        error_msg = None
        index = 3
        while index < len(raw_args):
            if raw_args[index] == "--sheet" and index + 1 < len(raw_args):
                sheet = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--prompt" and index + 1 < len(raw_args):
                prompt = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--error" and index + 1 < len(raw_args):
                error_msg = raw_args[index + 1]
                index += 2
            else:
                index += 1
        print_result(
            doc_server.add_dropdown(
                raw_args[0],
                raw_args[1],
                raw_args[2],
                sheet_name=sheet,
                prompt=prompt,
                error_message=error_msg,
            ),
            json_output,
        )
        return True

    if command == "xlsx-list-validations":
        if not raw_args:
            print_result(
                {"success": False, "error": command_usage("xlsx-list-validations")},
                json_output,
            )
            return True
        sheet = None
        index = 1
        while index < len(raw_args):
            if raw_args[index] == "--sheet" and index + 1 < len(raw_args):
                sheet = raw_args[index + 1]
                index += 2
            else:
                index += 1
        print_result(doc_server.list_validations(raw_args[0], sheet_name=sheet), json_output)
        return True

    if command == "xlsx-remove-validation":
        if not raw_args:
            print_result(
                {"success": False, "error": command_usage("xlsx-remove-validation")},
                json_output,
            )
            return True
        cell_range = None
        remove_all = False
        sheet = None
        index = 1
        while index < len(raw_args):
            if raw_args[index] == "--range" and index + 1 < len(raw_args):
                cell_range = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--all":
                remove_all = True
                index += 1
            elif raw_args[index] == "--sheet" and index + 1 < len(raw_args):
                sheet = raw_args[index + 1]
                index += 2
            else:
                index += 1
        print_result(
            doc_server.remove_validation(
                raw_args[0],
                cell_range=cell_range,
                sheet_name=sheet,
                remove_all=remove_all,
            ),
            json_output,
        )
        return True

    if command == "xlsx-validate-data":
        if not raw_args:
            print_result(
                {"success": False, "error": command_usage("xlsx-validate-data")},
                json_output,
            )
            return True
        sheet = None
        max_rows = 1000
        index = 1
        while index < len(raw_args):
            if raw_args[index] == "--sheet" and index + 1 < len(raw_args):
                sheet = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--max-rows" and index + 1 < len(raw_args):
                max_rows = int(raw_args[index + 1])
                index += 2
            else:
                index += 1
        print_result(
            doc_server.validate_data(raw_args[0], sheet_name=sheet, max_rows=max_rows),
            json_output,
        )
        return True

    if command == "xlsx-to-pdf":
        if len(raw_args) < 1:
            print_result(
                {"success": False, "error": command_usage("xlsx-to-pdf")},
                json_output,
            )
            return True
        print_result(
            doc_server.spreadsheet_to_pdf(
                raw_args[0],
                output_path=raw_args[1] if len(raw_args) >= 2 else None,
            ),
            json_output,
        )
        return True

    if command == "xlsx-preview":
        if len(raw_args) < 2:
            print_result(
                {
                    "success": False,
                    "error": command_usage("xlsx-preview"),
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
            doc_server.preview_spreadsheet(raw_args[0], raw_args[1], pages=pages, dpi=dpi, fmt=fmt),
            json_output,
        )
        return True

    if command == "chart-create":
        if len(raw_args) < 5:
            print_result(
                {
                    "success": False,
                    "error": command_usage("chart-create"),
                },
                json_output,
            )
            return True
        sheet = None
        output_sheet = None
        x_label = None
        y_label = None
        show_labels = False
        show_legend = True
        legend_pos = "right"
        colors = None
        index = 5
        while index < len(raw_args):
            if raw_args[index] == "--sheet" and index + 1 < len(raw_args):
                sheet = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--output-sheet" and index + 1 < len(raw_args):
                output_sheet = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--x-label" and index + 1 < len(raw_args):
                x_label = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--y-label" and index + 1 < len(raw_args):
                y_label = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--labels":
                show_labels = True
                index += 1
            elif raw_args[index] == "--no-legend":
                show_legend = False
                index += 1
            elif raw_args[index] == "--legend-pos" and index + 1 < len(raw_args):
                legend_pos = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--colors" and index + 1 < len(raw_args):
                colors = raw_args[index + 1]
                index += 2
            else:
                index += 1
        print_result(
            doc_server.create_chart(
                file_path=raw_args[0],
                chart_type=raw_args[1],
                data_range=raw_args[2],
                categories_range=raw_args[3],
                title=raw_args[4],
                sheet_name=sheet,
                output_sheet=output_sheet,
                x_label=x_label,
                y_label=y_label,
                show_data_labels=show_labels,
                show_legend=show_legend,
                legend_pos=legend_pos,
                colors=colors.split(",") if colors else None,
            ),
            json_output,
        )
        return True

    if command == "chart-comparison":
        if len(raw_args) < 3:
            print_result(
                {
                    "success": False,
                    "error": command_usage("chart-comparison"),
                },
                json_output,
            )
            return True
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
        index = 3
        while index < len(raw_args):
            if raw_args[index] == "--sheet" and index + 1 < len(raw_args):
                sheet = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--start-row" and index + 1 < len(raw_args):
                start_row = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--start-col" and index + 1 < len(raw_args):
                start_col = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--cats" and index + 1 < len(raw_args):
                num_cats = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--series" and index + 1 < len(raw_args):
                num_series = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--cat-col" and index + 1 < len(raw_args):
                cat_col = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--value-cols" and index + 1 < len(raw_args):
                value_cols = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--output" and index + 1 < len(raw_args):
                output_cell = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--labels":
                show_labels = True
                index += 1
            elif raw_args[index] == "--no-legend":
                show_legend = False
                index += 1
            else:
                index += 1
        print_result(
            doc_server.create_comparison_chart(
                file_path=raw_args[0],
                chart_type=raw_args[1],
                sheet_name=sheet,
                title=raw_args[2],
                start_row=int(start_row) if start_row else 1,
                start_col=int(start_col) if start_col else 1,
                num_categories=int(num_cats) if num_cats else None,
                num_series=int(num_series) if num_series else None,
                category_col=int(cat_col) if cat_col else 1,
                value_cols=[int(value) for value in value_cols.split(",")] if value_cols else None,
                output_cell=output_cell,
                show_data_labels=show_labels,
                show_legend=show_legend,
            ),
            json_output,
        )
        return True

    if command == "chart-grade-dist":
        if len(raw_args) < 3:
            print_result(
                {
                    "success": False,
                    "error": command_usage("chart-grade-dist"),
                },
                json_output,
            )
            return True
        sheet = None
        output_cell = "F2"
        index = 3
        while index < len(raw_args):
            if raw_args[index] == "--sheet" and index + 1 < len(raw_args):
                sheet = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--output" and index + 1 < len(raw_args):
                output_cell = raw_args[index + 1]
                index += 2
            else:
                index += 1
        print_result(
            doc_server.create_grade_distribution_chart(
                file_path=raw_args[0],
                sheet_name=sheet,
                grade_column=raw_args[1],
                title=raw_args[2],
                output_cell=output_cell,
            ),
            json_output,
        )
        return True

    if command == "chart-progress":
        if len(raw_args) < 4:
            print_result(
                {
                    "success": False,
                    "error": command_usage("chart-progress"),
                },
                json_output,
            )
            return True
        sheet = None
        output_cell = "D2"
        show_labels = True
        index = 4
        while index < len(raw_args):
            if raw_args[index] == "--sheet" and index + 1 < len(raw_args):
                sheet = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--output" and index + 1 < len(raw_args):
                output_cell = raw_args[index + 1]
                index += 2
            elif raw_args[index] == "--labels":
                show_labels = True
                index += 1
            elif raw_args[index] == "--no-labels":
                show_labels = False
                index += 1
            else:
                index += 1
        print_result(
            doc_server.create_progress_chart(
                file_path=raw_args[0],
                student_column=raw_args[1],
                grade_column=raw_args[2],
                sheet_name=sheet,
                title=raw_args[3],
                output_cell=output_cell,
                show_data_labels=show_labels,
            ),
            json_output,
        )
        return True

    return False
