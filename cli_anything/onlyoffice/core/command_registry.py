#!/usr/bin/env python3
"""Central command/help registry for the OnlyOffice CLI."""

from __future__ import annotations

from copy import deepcopy
from typing import Dict, List


VERSION = "4.4.11"


COMMAND_CATEGORIES: Dict[str, Dict[str, str]] = {
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
        "doc-layout <file> [--size A4|Letter] [--orientation portrait|landscape] [--margin-* <in>] [--header <text>] [--page-numbers]": "Set page size/layout/margins/header/footer",
        "doc-formatting-info <file>": "Inspect paragraph/section formatting",
        "doc-set-style <file> <index> <style>": "Set paragraph style (Heading 1, Normal, etc.)",
        "doc-list-styles <file>": "List all available paragraph/character styles",
        "doc-add-table <file> <headers_csv> <data_csv>": "Add table (rows separated by ;)",
        "doc-read-tables <file>": "Read all tables from document",
        "doc-add-image <file> <image_path> [--width <in>] [--caption <text>] [--paragraph <index>] [--position before|after]": "Add image with optional caption near a paragraph",
        "doc-add-hyperlink <file> <text> <url> [--paragraph <index>]": "Add hyperlink (-1 = new paragraph)",
        "doc-add-page-break <file>": "Insert page break",
        "doc-add-list <file> <items_csv> [--type bullet|number]": "Add bulleted or numbered list (items separated by ;)",
        "doc-add-reference <file> <ref_json>": "Add reference to sidecar .refs.json",
        "doc-build-references <file>": "Build APA 7th References section from sidecar",
        "doc-set-metadata <file> [--author <a>] [--title <t>] [--subject <s>] [--keywords <k>]": "Set document properties",
        "doc-get-metadata <file>": "Read document properties",
        "doc-inspect-hidden-data <file>": "Inspect hidden DOCX metadata, comments, revisions, and custom XML",
        "doc-sanitize <file> [output_path] [--remove-comments] [--accept-revisions] [--clear-metadata] [--remove-custom-xml] [--author <a>]": "Sanitize DOCX for submission",
        "doc-preflight <file> [--expected-page-size <A4|Letter>] [--expected-font <name>] [--expected-font-size <pt>]": "Run submission-oriented DOCX preflight checks",
        "doc-word-count <file>": "Word/character/paragraph counts",
        "doc-extract-images <file> <output_dir> [--format png|jpg] [--prefix <name>]": "Extract all embedded images from .docx",
        "doc-to-pdf <file> [output_path]": "Convert .docx to PDF via OnlyOffice",
        "doc-preview <file> <output_dir> [--pages <range>] [--dpi <n>] [--format png|jpg]": "Render DOCX pages as images via OnlyOffice",
        "doc-render-map <file>": "Map DOCX paragraphs/table cells to OnlyOffice-rendered PDF pages and bounding boxes",
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
        "xlsx-to-pdf <file> [output_path]": "Convert a spreadsheet to PDF via OnlyOffice",
        "xlsx-preview <file> <output_dir> [--pages <range>] [--dpi <n>] [--format png|jpg]": "Render spreadsheet pages as images via OnlyOffice",
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
        "pptx-extract-images <file> <output_dir> [--slide <index>] [--format png|jpg] [--prefix <name>]": "Extract all images from slides",
        "pptx-list-shapes <file> [--slide <index>]": "List all shapes with position/size/text (spatial map)",
        "pptx-add-textbox <file> <slide_index> <text> [--left <in>] [--top <in>] [--width <in>] [--height <in>] [--font-size <pt>] [--font-name <name>] [--bold] [--italic] [--color <hex>] [--align <dir>]": "Add textbox at exact position",
        "pptx-modify-shape <file> <slide_index> <shape_name> [--left <in>] [--top <in>] [--width <in>] [--height <in>] [--text <text>] [--font-size <pt>] [--rotation <deg>]": "Move/resize/edit any shape by name",
        "pptx-preview <file> <output_dir> [--slide <index>] [--dpi <n>]": "Render slides as PNG images (requires OnlyOffice Docker)",
    },
    "PDF (.pdf)": {
        "pdf-extract-images <file> <output_dir> [--format png|jpg] [--pages <range>]": "Extract embedded images from PDF (PyMuPDF)",
        "pdf-page-to-image <file> <output_dir> [--pages <range>] [--dpi <n>] [--format png|jpg]": "Render full PDF pages as images",
        "pdf-read-blocks <file> [--pages <range>] [--no-spans] [--no-images] [--include-empty]": "Read native PDF text blocks/lines/spans with bounding boxes",
        "pdf-search-blocks <file> <query> [--pages <range>] [--case-sensitive] [--no-spans]": "Search native PDF blocks/spans and return exact block/span anchors",
        "pdf-inspect-hidden-data <file>": "Inspect hidden PDF metadata, annotations, embedded files, and page-size consistency",
        "pdf-sanitize <file> [output_path] [--clear-metadata] [--remove-xml-metadata] [--author <a>]": "Sanitize PDF metadata/XMP for submission",
    },
    "RDF (Knowledge Graphs)": {
        "rdf-create <file> [--base <uri>] [--format turtle|xml|n3|json-ld] [--prefix <p>=<uri>]": "Create empty RDF graph with prefixes",
        "rdf-read <file> [--limit <n>]": "Read/parse RDF file and show triples",
        "rdf-add <file> <subject> <predicate> <object> [--type uri|literal|bnode] [--lang <l>] [--datatype <dt>]": "Add a triple",
        "rdf-remove <file> [--subject <s>] [--predicate <p>] [--object <o>] [--type uri|literal|bnode] [--lang <l>] [--datatype <dt>]": "Remove triples (None = wildcard)",
        "rdf-query <file> <sparql> [--limit <n>]": "Execute SPARQL query",
        "rdf-export <file> <output> [--format turtle|xml|n3|nt|json-ld|trig]": "Convert/export to different format",
        "rdf-merge <file_a> <file_b> [--output <file>] [--format turtle]": "Merge two RDF graphs",
        "rdf-stats <file>": "Graph statistics (subjects, predicates, types, etc.)",
        "rdf-namespace <file> [<prefix> <uri>]": "Add namespace prefix or list all",
        "rdf-validate <file> <shapes_file>": "SHACL validation (requires pyshacl)",
    },
    "GENERAL": {
        "list": "List recent documents, spreadsheets, presentations",
        "open <file> [gui|web]": "Open in OnlyOffice GUI or web viewer (aliases: document.open, spreadsheet.open, presentation.open, pdf.open)",
        "watch <file> [gui|web]": "Watch file for changes + auto-open (aliases: document.watch, spreadsheet.watch, presentation.watch, pdf.watch)",
        "info <file>": "File info (type, size, sheet/slide/paragraph counts; aliases: document.info, spreadsheet.info, presentation.info, pdf.info)",
        "backup-list <file> [--limit <n>]": "List backups for file",
        "backup-prune [--file <f>] [--keep <n>] [--days <n>]": "Prune old backups",
        "backup-restore <file> [--backup <p>] [--latest] [--dry-run]": "Restore from backup",
        "editor-session <file> [--open] [--wait <sec>] [--activate]": "Inspect or open a native OnlyOffice desktop editor session",
        "editor-capture <file> <output_image> [--backend auto|desktop|rendered] [--page <n>] [--range <A1:D20>] [--slide <n>] [--zoom-reset] [--zoom-in <n>] [--zoom-out <n>] [--crop x,y,w,h]": "Capture a live editor viewport or rendered fallback image",
        "status": "Check installation and all capabilities",
        "help": "Show this help",
    },
}


HELP_EXAMPLES: List[str] = [
    "# Documents",
    "cli-anything-onlyoffice doc-create /tmp/essay.docx 'My Essay' 'Introduction paragraph here'",
    "cli-anything-onlyoffice doc-insert /tmp/essay.docx 'New first paragraph' 0 --style 'Heading 1'",
    "cli-anything-onlyoffice doc-search /tmp/essay.docx 'introduction'",
    "cli-anything-onlyoffice doc-add-hyperlink /tmp/essay.docx 'Click here' 'https://example.com'",
    "cli-anything-onlyoffice doc-add-list /tmp/essay.docx 'First item;Second item;Third item' --type bullet",
    "cli-anything-onlyoffice doc-preview /tmp/essay.docx /tmp/doc-preview --pages 1-2",
    "cli-anything-onlyoffice doc-render-map /tmp/essay.docx",
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
    "cli-anything-onlyoffice xlsx-preview /tmp/data.xlsx /tmp/xlsx-preview --pages 0",
    "cli-anything-onlyoffice editor-capture /tmp/data.xlsx /tmp/current-view.png --backend desktop --range Sheet0!A1:F20 --crop 100,120,1400,800",
    "# Presentations",
    "cli-anything-onlyoffice pptx-create /tmp/lecture.pptx 'Lecture 1' 'Introduction'",
    "cli-anything-onlyoffice pptx-speaker-notes /tmp/lecture.pptx 0 'Remember to introduce yourself'",
    "cli-anything-onlyoffice pptx-slide-count /tmp/lecture.pptx",
    "# PDFs",
    "cli-anything-onlyoffice pdf-read-blocks /tmp/paper.pdf --pages 0-1",
    "cli-anything-onlyoffice pdf-search-blocks /tmp/paper.pdf 'Results' --pages 2-3",
    "# RDF Knowledge Graphs",
    "cli-anything-onlyoffice rdf-create /tmp/knowledge.ttl --base http://example.org/",
    "cli-anything-onlyoffice rdf-add /tmp/knowledge.ttl http://example.org/Alice http://xmlns.com/foaf/0.1/name Alice --type literal",
    "cli-anything-onlyoffice rdf-query /tmp/knowledge.ttl 'SELECT ?s ?p ?o WHERE { ?s ?p ?o } LIMIT 10'",
    "cli-anything-onlyoffice rdf-stats /tmp/knowledge.ttl",
]


CATEGORY_COUNTS: Dict[str, int] = {
    category: len(commands) for category, commands in COMMAND_CATEGORIES.items()
}
TOTAL_COMMANDS = sum(CATEGORY_COUNTS.values())
COMMAND_SIGNATURES: Dict[str, str] = {
    signature.split()[0]: signature
    for commands in COMMAND_CATEGORIES.values()
    for signature in commands
}
USAGE_OVERRIDES: Dict[str, str] = {
    "chart-comparison": "chart-comparison <file> <type> <title> [--sheet <name>] [--start-row <n>] [--start-col <n>] [--cats <n>] [--series <n>] [--cat-col <n>] [--value-cols <n,n,n>] [--output <cell>] [--labels] [--no-legend]",
    "chart-create": "chart-create <file> <type> <data_range> <cat_range> <title> [--sheet <name>] [--output-sheet <name>] [--x-label <text>] [--y-label <text>] [--labels] [--no-legend] [--legend-pos <pos>] [--colors <hex,hex>]",
    "chart-grade-dist": "chart-grade-dist <file> <grade_col> <title> [--sheet <name>] [--output <cell>]",
    "chart-progress": "chart-progress <file> <student_col> <grade_col> <title> [--sheet <name>] [--output <cell>] [--labels]",
    "doc-add-image": "doc-add-image <file> <image_path> [--width <inches>] [--caption <text>] [--paragraph <index>] [--position before|after]",
    "doc-add-list": "doc-add-list <file> <items> [--type bullet|number] (items separated by ;)",
    "doc-add-reference": 'doc-add-reference <file> <ref_json>  (ref_json: {"author":"...","year":"...","title":"...","source":"...","type":"journal|book|website|report|chapter","doi":"..."})',
    "doc-add-table": "doc-add-table <file> <headers_csv> <data_csv>  (rows separated by ';')",
    "doc-build-references": "doc-build-references <file>  (reads <file>.refs.json, appends APA 7th formatted References section)",
    "doc-delete": "doc-delete <file> <paragraph_index>",
    "doc-layout": "doc-layout <file> [--size <A4|Letter>] [--orientation <portrait|landscape>] [--margin-top <in>] [--margin-bottom <in>] [--margin-left <in>] [--margin-right <in>] [--header <text>] [--page-numbers]",
    "doc-sanitize": "doc-sanitize <file> [output_path] [--remove-comments] [--accept-revisions] [--clear-metadata] [--remove-custom-xml] [--set-remove-personal-information] [--author <a>] [--title <t>] [--subject <s>] [--keywords <k>]",
    "doc-set-metadata": "doc-set-metadata <file> [--author <a>] [--title <t>] [--subject <s>] [--keywords <k>] [--comments <c>] [--category <cat>]",
    'doc-set-style': 'doc-set-style <file> <paragraph_index> <style>  (e.g., "Heading 1", "Heading 2", "Normal", "Title")',
    "pdf-sanitize": "pdf-sanitize <file> [output_path] [--clear-metadata] [--remove-xml-metadata] [--author <a>] [--title <t>] [--subject <s>] [--keywords <k>] [--creator <c>] [--producer <p>]",
    "pptx-add-textbox": "pptx-add-textbox <file> <slide_index> <text> [--left <in>] [--top <in>] [--width <in>] [--height <in>] [--font-size <pt>] [--font-name <name>] [--bold] [--italic] [--color <hex>] [--align <left|center|right>]",
    "pptx-speaker-notes": "pptx-speaker-notes <file> <slide_index> [notes_text]",
    "pptx-update-text": "pptx-update-text <file> <slide_index> [--title <t>] [--body <b>]",
    "rdf-add": "rdf-add <file> <subject> <predicate> <object> [--type uri|literal|bnode] [--lang <l>] [--datatype <dt>] [--format <f>]",
    "rdf-export": "rdf-export <file> <output_file> [--format turtle|xml|n3|nt|json-ld|trig]",
    "rdf-namespace": "rdf-namespace <file> [<prefix> <uri>] [--format <f>]",
    "rdf-query": "rdf-query <file> <sparql_query> [--limit <n>]",
    "rdf-remove": "rdf-remove <file> [--subject <s>] [--predicate <p>] [--object <o>] [--type uri|literal|bnode] [--lang <l>] [--datatype <dt>] [--format <f>]",
    "rdf-validate": "rdf-validate <data_file> <shapes_file>",
    "xlsx-add-dropdown": "xlsx-add-dropdown <file> <range> <options_csv> [--sheet <name>] [--prompt <msg>] [--error <msg>]",
    "xlsx-add-validation": "xlsx-add-validation <file> <range> <type> [--operator <op>] [--formula1 <v>] [--formula2 <v>] [--sheet <name>] [--error <msg>] [--prompt <msg>] [--error-style stop|warning|information] [--allow-blank]",
    "xlsx-calc": "xlsx-calc <file> <column> <operation> [--sheet <name>] [--include-formulas] [--strict-formulas]",
    "xlsx-chi2": "xlsx-chi2 <file> <row_col> <col_col> [--sheet <name>] [--row-valid <csv>] [--col-valid <csv>]",
    "xlsx-corr": "xlsx-corr <file> <x_col> <y_col> [--sheet <name>] [--method pearson|spearman]",
    "xlsx-filter": "xlsx-filter <file> <column> <op> <value> [--sheet <name>] (op: eq|ne|gt|lt|ge|le|contains|startswith|endswith)",
    "xlsx-format-cells": "xlsx-format-cells <file> <range> [--sheet <name>] [--bold] [--italic] [--font-name <n>] [--font-size <n>] [--color <hex>] [--bg-color <hex>] [--number-format <fmt>] [--align <l|c|r>] [--wrap]",
    "xlsx-mannwhitney": "xlsx-mannwhitney <file> <value_col> <group_col> <group_a> <group_b> [--sheet <name>]",
    "xlsx-remove-validation": "xlsx-remove-validation <file> [--range <range>] [--all] [--sheet <name>]",
    "xlsx-research-pack": "xlsx-research-pack <file> [--sheet <name>] [--profile hlth3112] [--require-formula-safe]",
    "xlsx-search": "xlsx-search <file> <search-text> [--sheet <name>]",
    "xlsx-sheet-rename": "xlsx-sheet-rename <file> <old_name> <new_name>",
    "xlsx-text-extract": "xlsx-text-extract <file> <column> [--sheet <name>] [--limit <n>] [--min-length <n>]",
    "xlsx-text-keywords": "xlsx-text-keywords <file> <column> [--sheet <name>] [--top <n>] [--min-word-length <n>]",
    "xlsx-ttest": "xlsx-ttest <file> <value_col> <group_col> <group_a> <group_b> [--sheet <name>] [--equal-var]",
    "open": "open <file> [gui|web]",
    "watch": "watch <file> [gui|web]",
    "info": "info <file>",
    "editor-session": "editor-session <file> [--open] [--wait <sec>] [--activate]",
    "editor-capture": "editor-capture <file> <output_image> [--backend auto|desktop|rendered] [--open] [--page <n>] [--range <A1:D20>] [--slide <n>] [--zoom-reset] [--zoom-in <n>] [--zoom-out <n>] [--crop x,y,w,h] [--wait <sec>] [--settle-ms <n>] [--dpi <n>] [--format png|jpg]",
    "backup-list": "backup-list <file> [--limit <n>]",
    "backup-restore": "backup-restore <file> [--backup <name|path>] [--latest] [--dry-run]",
}


def get_command_categories() -> Dict[str, Dict[str, str]]:
    """Return a deep copy of the command catalogue."""
    return deepcopy(COMMAND_CATEGORIES)


def get_help_examples() -> List[str]:
    """Return a copy of the CLI help examples."""
    return list(HELP_EXAMPLES)


def command_signature(command: str) -> str:
    """Return the canonical command signature without the leading Usage: prefix."""
    return USAGE_OVERRIDES.get(command, COMMAND_SIGNATURES.get(command, command))


def command_usage(command: str) -> str:
    """Return the command usage string with the leading Usage: prefix."""
    return f"Usage: {command_signature(command)}"


def usage_error(command: str) -> Dict[str, str]:
    """Build a standard usage error payload for a command."""
    return {"success": False, "error": command_usage(command)}


def build_help_payload(capabilities: Dict[str, bool]) -> Dict[str, object]:
    """Build the public help payload from the registry."""
    return {
        "success": True,
        "version": VERSION,
        "categories": get_command_categories(),
        "capabilities": dict(capabilities),
        "total_commands": TOTAL_COMMANDS,
        "examples": get_help_examples(),
        "category_counts": dict(CATEGORY_COUNTS),
    }
