#!/usr/bin/env python3
"""Central command/help registry for the OnlyOffice CLI."""

from __future__ import annotations

from copy import deepcopy
from typing import Dict, List


VERSION = "4.4.19"
CLI_SCHEMA_VERSION = "1.0"


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
        "doc-normalize-format <file> [output_path] [--font <name>] [--body-size <pt>] [--title-size <pt>] [--line-spacing <single|1.5|double>] [--paragraph-after <pt>] [--clear-theme-fonts] [--skip-header-footer] [--remove-style-borders] [--reference-hanging <in>]": "Normalize whole-document academic formatting",
        "doc-formatting-info <file> [--all] [--start <n>] [--limit <n>]": "Inspect paragraph/section formatting, including indents, tabs, and page breaks",
        "doc-font-audit <file> [--expected-font <name>] [--expected-font-size <pt>] [--rendered] [--pdf <path>]": "Audit DOCX-declared and optionally PDF-rendered fonts",
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
        "doc-citation-audit <file> [--include-sidecar]": "Audit APA-like in-text citations against the DOCX reference list",
        "doc-set-metadata <file> [--author <a>] [--title <t>] [--subject <s>] [--keywords <k>]": "Set document properties",
        "doc-get-metadata <file>": "Read document properties",
        "doc-inspect-hidden-data <file>": "Inspect hidden DOCX metadata, comments, revisions, and custom XML",
        "doc-sanitize <file> [output_path] [--remove-comments] [--accept-revisions] [--clear-metadata] [--remove-custom-xml] [--set-remove-personal-information] [--canonicalize-ooxml] [--author <a>] [--title <t>] [--subject <s>] [--keywords <k>]": "Sanitize DOCX for submission",
        "doc-preflight <file> [--expected-page-size <A4|Letter>] [--expected-font <name>] [--expected-font-size <pt>] [--rendered-layout] [--profile auto|generic|apa-references]": "Run submission-oriented DOCX preflight checks",
        "doc-submission-pack <file> <output_dir> [--basename <name>] [--expected-page-size <A4|Letter>] [--expected-font <name>] [--expected-font-size <pt>] [--profile auto|generic|apa-references] [--skip-docx-sanitize] [--skip-pdf-sanitize] [--skip-rendered-layout]": "Create cleaned DOCX/PDF submission pack with manifest",
        "doc-word-count <file>": "Word/character/paragraph counts",
        "doc-extract-images <file> <output_dir> [--format png|jpg] [--prefix <name>]": "Extract all embedded images from .docx",
        "doc-to-pdf <file> [output_path] [--layout-warnings] [--profile auto|generic|apa-references]": "Convert .docx to PDF via OnlyOffice",
        "doc-preview <file> <output_dir> [--pages <range>] [--dpi <n>] [--format png|jpg]": "Render DOCX pages as images via OnlyOffice",
        "doc-render-map <file>": "Map DOCX paragraphs/table cells to OnlyOffice-rendered PDF pages and bounding boxes",
        "doc-render-audit <file> [--pdf <path>] [--tolerance-points <n>] [--profile auto|generic|apa-references]": "Audit rendered PDF margins, page breaks, and hanging indents against DOCX intent",
    },
    "SPREADSHEETS (.xlsx)": {
        "xlsx-create <file> [sheet]": "Create new .xlsx spreadsheet",
        "xlsx-write <file> <headers> <data> [--sheet <name>] [--overwrite] [--coerce-rows] [--text-columns <csv>]": "Write data to sheet (safe update default)",
        "xlsx-read <file> [sheet]": "Read spreadsheet data (all sheets if no sheet specified)",
        "xlsx-append <file> <row-data> [--sheet <name>]": "Append row to spreadsheet",
        "xlsx-search <file> <search-text> [--sheet <name>]": "Search for text in spreadsheet",
        "xlsx-cell-read <file> <cell> [--sheet <name>]": "Read a single cell (e.g., B5)",
        "xlsx-cell-write <file> <cell> <value> [--sheet <name>] [--text]": "Write to a single cell",
        "xlsx-range-read <file> <range> [--sheet <name>]": "Read a cell range (e.g., A1:D10)",
        "xlsx-delete-rows <file> <start_row> [count] [--sheet <name>]": "Delete rows (1-indexed)",
        "xlsx-delete-cols <file> <start_col> [count] [--sheet <name>]": "Delete columns (1-indexed)",
        "xlsx-sort <file> <column> [--sheet <name>] [--desc] [--numeric]": "Sort by column (preserves header)",
        "xlsx-filter <file> <column> <op> <value> [--sheet <name>] (op: eq|ne|gt|lt|ge|le|contains|startswith|endswith)": "Filter rows (op: eq|ne|gt|lt|ge|le|contains|startswith|endswith)",
        "xlsx-calc <file> <column> <operation> [--sheet <name>] [--include-formulas] [--strict-formulas]": "Column statistics (sum|avg|min|max|all)",
        "xlsx-formula <file> <cell> <formula> [--sheet <name>]": "Add formula to cell",
        "xlsx-formula-audit <file> [--sheet <name>]": "Audit formula complexity/risk",
        "xlsx-freq <file> <column> [--sheet <name>] [--valid <csv>]": "Frequency + percentage table",
        "xlsx-corr <file> <x_col> <y_col> [--sheet <name>] [--method pearson|spearman]": "Correlation test with APA output",
        "xlsx-ttest <file> <value_col> <group_col> <group_a> <group_b> [--sheet <name>] [--equal-var]": "Independent t-test (Welch default, Cohen's d)",
        "xlsx-mannwhitney <file> <value_col> <group_col> <group_a> <group_b> [--sheet <name>]": "Mann-Whitney U test (non-parametric)",
        "xlsx-chi2 <file> <row_col> <col_col> [--sheet <name>] [--row-valid <csv>] [--col-valid <csv>]": "Chi-square test (Cramer's V)",
        "xlsx-research-pack <file> [--sheet <name>] [--profile hlth3112] [--require-formula-safe]": "Bundled analysis pack",
        "xlsx-text-extract <file> <column> [--sheet <name>] [--limit <n>] [--min-length <n>]": "Extract open-text responses",
        "xlsx-text-keywords <file> <column> [--sheet <name>] [--top <n>] [--min-word-length <n>]": "Keyword frequency for themes",
        "xlsx-sheet-list <file>": "List all sheets with row/column counts",
        "xlsx-sheet-add <file> <name> [--position <n>]": "Add a new sheet",
        "xlsx-sheet-delete <file> <name>": "Delete a sheet",
        "xlsx-sheet-rename <file> <old_name> <new_name>": "Rename a sheet",
        "xlsx-merge-cells <file> <range> [--sheet <name>]": "Merge cells (e.g., A1:C1)",
        "xlsx-unmerge-cells <file> <range> [--sheet <name>]": "Unmerge cells",
        "xlsx-format-cells <file> <range> [--sheet <name>] [--bold] [--italic] [--font-name <n>] [--font-size <n>] [--color <hex>] [--bg-color <hex>] [--number-format <fmt>] [--align <l|c|r>] [--wrap]": "Format cell range",
        "xlsx-csv-import <xlsx_file> <csv_file> [--sheet <name>] [--delimiter <char>]": "Import CSV into xlsx",
        "xlsx-csv-export <xlsx_file> <csv_file> [--sheet <name>] [--delimiter <char>]": "Export sheet to CSV",
        "xlsx-add-validation <file> <range> <type> [--operator <op>] [--formula1 <v>] [--formula2 <v>] [--sheet <name>] [--error <msg>] [--error-title <title>] [--prompt <msg>] [--prompt-title <title>] [--error-style stop|warning|information] [--no-blank]": "Add data validation (list|whole|decimal|textLength|date|custom)",
        "xlsx-add-dropdown <file> <range> <options_csv> [--sheet <name>] [--prompt <msg>] [--error <msg>]": "Add dropdown list validation (shortcut)",
        "xlsx-list-validations <file> [--sheet <name>]": "List all validation rules on a sheet",
        "xlsx-remove-validation <file> [--range <range>] [--all] [--sheet <name>]": "Remove validation rules",
        "xlsx-validate-data <file> [--sheet <name>] [--max-rows <n>]": "Audit data against validation rules (pass/fail per cell)",
        "xlsx-to-pdf <file> [output_path]": "Convert a spreadsheet to PDF via OnlyOffice",
        "xlsx-preview <file> <output_dir> [--pages <range>] [--dpi <n>] [--format png|jpg]": "Render spreadsheet pages as images via OnlyOffice",
    },
    "CHARTS (.xlsx)": {
        "chart-create <file> <type> <data_range> <cat_range> <title> [--sheet <name>] [--output-sheet <name>] [--x-label <text>] [--y-label <text>] [--labels] [--no-legend] [--legend-pos <pos>] [--colors <hex,hex>]": "Create chart (bar|line|pie|scatter|bar_horizontal)",
        "chart-comparison <file> <type> <title> [--sheet <name>] [--start-row <n>] [--start-col <n>] [--cats <n>] [--series <n>] [--cat-col <n>] [--value-cols <n,n,n>] [--output <cell>] [--labels] [--no-legend]": "Comparison chart from structured data",
        "chart-grade-dist <file> <grade_col> <title> [--sheet <name>] [--output <cell>]": "Grade distribution pie chart",
        "chart-progress <file> <student_col> <grade_col> <title> [--sheet <name>] [--output <cell>] [--labels] [--no-labels]": "Student progress bar chart",
    },
    "PRESENTATIONS (.pptx)": {
        "pptx-create <file> <title> [subtitle]": "Create new presentation",
        "pptx-add-slide <file> <title> [content] [layout]": "Add slide (title_only|content|blank|two_content|comparison)",
        "pptx-add-bullets <file> <title> <bullets>": "Bullet-point slide (\\n separated)",
        "pptx-read <file>": "Read all slides and content",
        "pptx-add-image <file> <title> <image_path>": "Add image slide",
        "pptx-add-table <file> <title> <headers> <data> [--coerce-rows]": "Add table slide",
        "pptx-delete-slide <file> <index>": "Delete slide by index (0-based)",
        "pptx-speaker-notes <file> <slide_index> [notes_text]": "Read or set speaker notes (omit text to read)",
        "pptx-update-text <file> <slide_index> [--title <t>] [--body <b>]": "Update existing slide text",
        "pptx-slide-count <file>": "Get slide count and titles",
        "pptx-extract-images <file> <output_dir> [--slide <index>] [--format png|jpg] [--prefix <name>]": "Extract all images from slides",
        "pptx-list-shapes <file> [--slide <index>]": "List all shapes with position/size/text (spatial map)",
        "pptx-add-textbox <file> <slide_index> <text> [--left <in>] [--top <in>] [--width <in>] [--height <in>] [--font-size <pt>] [--font-name <name>] [--bold] [--italic] [--color <hex>] [--align <left|center|right>]": "Add textbox at exact position",
        "pptx-modify-shape <file> <slide_index> <shape_name> [--left <in>] [--top <in>] [--width <in>] [--height <in>] [--text <text>] [--font-size <pt>] [--rotation <deg>]": "Move/resize/edit any shape by name",
        "pptx-preview <file> <output_dir> [--slide <index>] [--dpi <n>]": "Render slides as PNG images (requires OnlyOffice Docker)",
    },
    "PDF (.pdf)": {
        "pdf-extract-images <file> <output_dir> [--format png|jpg] [--pages <range>]": "Extract embedded images from PDF (PyMuPDF)",
        "pdf-page-to-image <file> <output_dir> [--pages <range>] [--dpi <n>] [--format png|jpg]": "Render full PDF pages as images",
        "pdf-map-page <file> <page> <output_image> [--dpi <n>] [--format png|jpg] [--no-labels] [--no-images]": "Render a PDF page with selectable block-id overlays",
        "pdf-read-blocks <file> [--pages <range>] [--no-spans] [--no-images] [--include-empty]": "Read native PDF text blocks/lines/spans with bounding boxes",
        "pdf-search-blocks <file> <query> [--pages <range>] [--case-sensitive] [--no-spans]": "Search native PDF blocks/spans and return exact block/span anchors",
        "pdf-inspect-hidden-data <file>": "Inspect hidden PDF metadata, annotations, embedded files, and page-size consistency",
        "pdf-sanitize <file> [output_path] [--clear-metadata] [--remove-xml-metadata] [--remove-annotations] [--remove-embedded-files] [--flatten-forms] [--author <a>] [--title <t>] [--subject <s>] [--keywords <k>] [--creator <c>] [--producer <p>]": "Sanitize PDF metadata/XMP and explicitly requested hidden objects for submission",
        "pdf-compact <file> [output_path] [--garbage <0-4>] [--no-deflate] [--no-clean] [--linearize]": "Explicitly compact/optimize a PDF; never applied by default",
        "pdf-merge <input_file> <input_file> [input_file ...] --output <file>": "Stitch multiple PDFs into one output PDF",
        "pdf-split <file> <output_dir> [--pages <range>] [--prefix <name>]": "Split selected PDF pages into one-page PDFs",
        "pdf-reorder <file> <page_order> [output_path]": "Create a PDF with pages in explicit order, preserving duplicates",
        "pdf-add-text <file> <page> <text> [--output <file>] [--x <pt>] [--y <pt>] [--width <pt>] [--height <pt>] [--font-size <pt>] [--font <name>] [--color <hex>] [--rotation <0|90|180|270>]": "Overlay bounded text onto a PDF page",
        "pdf-add-image <file> <page> <image_path> [--output <file>] [--x <pt>] [--y <pt>] [--width <pt>] [--height <pt>] [--no-keep-proportion]": "Overlay an image onto a PDF page",
        "pdf-redact <file> [output_path] (--text <query> | --rect <page,left,top,right,bottom>) [--pages <range>] [--case-sensitive] [--fill <hex>] [--dry-run]": "Apply true PDF redactions by exact text match or rectangle",
        "pdf-redact-block <file> <block_id> [output_path] [--fill <hex>] [--dry-run]": "Apply true PDF redaction to a native block id from pdf-map-page/pdf-read-blocks",
    },
    "RDF (Knowledge Graphs)": {
        "rdf-create <file> [--base <uri>] [--format turtle|xml|n3|json-ld] [--prefix <p>=<uri>]": "Create empty RDF graph with prefixes",
        "rdf-read <file> [--limit <n>]": "Read/parse RDF file and show triples",
        "rdf-add <file> <subject> <predicate> <object> [--type uri|literal|bnode] [--lang <l>] [--datatype <dt>] [--format <f>]": "Add a triple",
        "rdf-remove <file> (--all | [--subject <s>] [--predicate <p>] [--object <o>]) [--type uri|literal|bnode] [--lang <l>] [--datatype <dt>] [--format <f>] [--dry-run]": "Remove triples matching selectors, or all triples with explicit --all",
        "rdf-query <file> <sparql> [--limit <n>]": "Execute SPARQL query",
        "rdf-export <file> <output> [--format turtle|xml|n3|nt|json-ld|trig]": "Convert/export to different format",
        "rdf-merge <file_a> <file_b> [--output <file>] [--format turtle]": "Merge two RDF graphs",
        "rdf-stats <file>": "Graph statistics (subjects, predicates, types, etc.)",
        "rdf-namespace <file> [<prefix> <uri>] [--format <f>]": "Add namespace prefix or list all",
        "rdf-validate <file> <shapes_file>": "SHACL validation (requires pyshacl)",
    },
    "GENERAL": {
        "list": "List recent documents, spreadsheets, presentations",
        "open <file> [gui|web]": "Open in OnlyOffice GUI or web viewer (aliases: document.open, spreadsheet.open, presentation.open, pdf.open)",
        "watch <file> [gui|web]": "Watch file for changes + auto-open (aliases: document.watch, spreadsheet.watch, presentation.watch, pdf.watch)",
        "info <file>": "File info (type, size, sheet/slide/paragraph counts; aliases: document.info, spreadsheet.info, presentation.info, pdf.info)",
        "backup-list <file> [--limit <n>]": "List backups for file",
        "backup-prune [--file <f>] [--keep <n>] [--days <n>]": "Prune old backups",
        "backup-restore <file> [--backup <name|path>] [--latest] [--dry-run]": "Restore from backup",
        "editor-session <file> [--open] [--wait <sec>] [--activate]": "Inspect or open a native OnlyOffice desktop editor session",
        "editor-capture <file> <output_image> [--backend auto|desktop|rendered] [--open] [--page <n>] [--range <A1:D20>] [--slide <n>] [--zoom-reset] [--zoom-in <n>] [--zoom-out <n>] [--crop x,y,w,h] [--wait <sec>] [--settle-ms <n>] [--dpi <n>] [--format png|jpg]": "Capture a live editor viewport or rendered fallback image",
        "setup-check": "Strict install/update dependency check for cloned checkouts",
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
    "cli-anything-onlyoffice doc-citation-audit /tmp/essay.docx --json",
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
    "chart-progress": "chart-progress <file> <student_col> <grade_col> <title> [--sheet <name>] [--output <cell>] [--labels] [--no-labels]",
    "doc-add-image": "doc-add-image <file> <image_path> [--width <inches>] [--caption <text>] [--paragraph <index>] [--position before|after]",
    "doc-add-list": "doc-add-list <file> <items> [--type bullet|number] (items separated by ;)",
    "doc-add-reference": 'doc-add-reference <file> <ref_json>  (ref_json: {"author":"...","year":"...","title":"...","source":"...","type":"journal|book|website|report|chapter","doi":"..."})',
    "doc-add-table": "doc-add-table <file> <headers_csv> <data_csv>  (rows separated by ';')",
    "doc-build-references": "doc-build-references <file>  (reads <file>.refs.json, appends APA 7th formatted References section)",
    "doc-citation-audit": "doc-citation-audit <file> [--include-sidecar]",
    "doc-delete": "doc-delete <file> <paragraph_index>",
    "doc-formatting-info": "doc-formatting-info <file> [--all] [--start <n>] [--limit <n>]",
    "doc-font-audit": "doc-font-audit <file> [--expected-font <name>] [--expected-font-size <pt>] [--rendered] [--pdf <path>]",
    "doc-layout": "doc-layout <file> [--size <A4|Letter>] [--orientation <portrait|landscape>] [--margin-top <in>] [--margin-bottom <in>] [--margin-left <in>] [--margin-right <in>] [--header <text>] [--page-numbers]",
    "doc-normalize-format": "doc-normalize-format <file> [output_path] [--font <name>] [--body-size <pt>] [--title-size <pt>] [--line-spacing <single|1.5|double>] [--paragraph-after <pt>] [--clear-theme-fonts] [--skip-header-footer] [--remove-style-borders] [--reference-hanging <in>]",
    "doc-preflight": "doc-preflight <file> [--expected-page-size <A4|Letter>] [--expected-font <name>] [--expected-font-size <pt>] [--rendered-layout] [--profile auto|generic|apa-references]",
    "doc-render-audit": "doc-render-audit <file> [--pdf <path>] [--tolerance-points <n>] [--profile auto|generic|apa-references]",
    "doc-sanitize": "doc-sanitize <file> [output_path] [--remove-comments] [--accept-revisions] [--clear-metadata] [--remove-custom-xml] [--set-remove-personal-information] [--canonicalize-ooxml] [--author <a>] [--title <t>] [--subject <s>] [--keywords <k>]",
    "doc-set-metadata": "doc-set-metadata <file> [--author <a>] [--title <t>] [--subject <s>] [--keywords <k>] [--comments <c>] [--category <cat>]",
    'doc-set-style': 'doc-set-style <file> <paragraph_index> <style>  (e.g., "Heading 1", "Heading 2", "Normal", "Title")',
    "doc-submission-pack": "doc-submission-pack <file> <output_dir> [--basename <name>] [--expected-page-size <A4|Letter>] [--expected-font <name>] [--expected-font-size <pt>] [--profile auto|generic|apa-references] [--skip-docx-sanitize] [--skip-pdf-sanitize] [--skip-rendered-layout]",
    "doc-to-pdf": "doc-to-pdf <file> [output_path] [--layout-warnings] [--profile auto|generic|apa-references]",
    "pdf-add-image": "pdf-add-image <file> <page> <image_path> [--output <file>] [--x <pt>] [--y <pt>] [--width <pt>] [--height <pt>] [--no-keep-proportion]",
    "pdf-add-text": "pdf-add-text <file> <page> <text> [--output <file>] [--x <pt>] [--y <pt>] [--width <pt>] [--height <pt>] [--font-size <pt>] [--font <name>] [--color <hex>] [--rotation <0|90|180|270>]",
    "pdf-compact": "pdf-compact <file> [output_path] [--garbage <0-4>] [--no-deflate] [--no-clean] [--linearize]",
    "pdf-merge": "pdf-merge <input_file> <input_file> [input_file ...] --output <file>",
    "pdf-map-page": "pdf-map-page <file> <page> <output_image> [--dpi <n>] [--format png|jpg] [--no-labels] [--no-images]",
    "pdf-redact": "pdf-redact <file> [output_path] (--text <query> | --rect <page,left,top,right,bottom>) [--pages <range>] [--case-sensitive] [--fill <hex>] [--dry-run]",
    "pdf-redact-block": "pdf-redact-block <file> <block_id> [output_path] [--fill <hex>] [--dry-run]",
    "pdf-reorder": "pdf-reorder <file> <page_order> [output_path]",
    "pdf-sanitize": "pdf-sanitize <file> [output_path] [--clear-metadata] [--remove-xml-metadata] [--remove-annotations] [--remove-embedded-files] [--flatten-forms] [--author <a>] [--title <t>] [--subject <s>] [--keywords <k>] [--creator <c>] [--producer <p>]",
    "pdf-split": "pdf-split <file> <output_dir> [--pages <range>] [--prefix <name>]",
    "pptx-add-textbox": "pptx-add-textbox <file> <slide_index> <text> [--left <in>] [--top <in>] [--width <in>] [--height <in>] [--font-size <pt>] [--font-name <name>] [--bold] [--italic] [--color <hex>] [--align <left|center|right>]",
    "pptx-speaker-notes": "pptx-speaker-notes <file> <slide_index> [notes_text]",
    "pptx-update-text": "pptx-update-text <file> <slide_index> [--title <t>] [--body <b>]",
    "rdf-add": "rdf-add <file> <subject> <predicate> <object> [--type uri|literal|bnode] [--lang <l>] [--datatype <dt>] [--format <f>]",
    "rdf-export": "rdf-export <file> <output_file> [--format turtle|xml|n3|nt|json-ld|trig]",
    "rdf-namespace": "rdf-namespace <file> [<prefix> <uri>] [--format <f>]",
    "rdf-query": "rdf-query <file> <sparql_query> [--limit <n>]",
    "rdf-remove": "rdf-remove <file> (--all | [--subject <s>] [--predicate <p>] [--object <o>]) [--type uri|literal|bnode] [--lang <l>] [--datatype <dt>] [--format <f>] [--dry-run]",
    "rdf-validate": "rdf-validate <data_file> <shapes_file>",
    "xlsx-add-dropdown": "xlsx-add-dropdown <file> <range> <options_csv> [--sheet <name>] [--prompt <msg>] [--error <msg>]",
    "xlsx-add-validation": "xlsx-add-validation <file> <range> <type> [--operator <op>] [--formula1 <v>] [--formula2 <v>] [--sheet <name>] [--error <msg>] [--error-title <title>] [--prompt <msg>] [--prompt-title <title>] [--error-style stop|warning|information] [--no-blank]",
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
    "setup-check": "setup-check",
    "backup-list": "backup-list <file> [--limit <n>]",
    "backup-restore": "backup-restore <file> [--backup <name|path>] [--latest] [--dry-run]",
}


CAPABILITY_DETAILS: Dict[str, Dict[str, object]] = {
    "python_docx": {
        "label": "python-docx",
        "category": "dependency",
        "description": "DOCX document creation, reading, editing, formatting, metadata, and preflight support.",
        "commands": ["doc-*"],
    },
    "openpyxl": {
        "label": "openpyxl",
        "category": "dependency",
        "description": "XLSX workbook reading, writing, validation, charts, CSV, and statistics support.",
        "commands": ["xlsx-*", "chart-*"],
    },
    "python_pptx": {
        "label": "python-pptx",
        "category": "dependency",
        "description": "PPTX presentation creation, reading, slide, shape, notes, and media support.",
        "commands": ["pptx-*"],
    },
    "rdflib": {
        "label": "rdflib",
        "category": "dependency",
        "description": "RDF graph parsing, mutation, querying, export, and namespace support.",
        "commands": ["rdf-create", "rdf-read", "rdf-add", "rdf-remove", "rdf-query", "rdf-export", "rdf-merge", "rdf-stats", "rdf-namespace"],
    },
    "pyshacl": {
        "label": "pyshacl",
        "category": "dependency",
        "description": "SHACL validation support for rdf-validate.",
        "commands": ["rdf-validate"],
    },
    "docx_create": {
        "label": "DOCX create",
        "category": "DOCUMENTS (.docx)",
        "requires": ["python_docx"],
        "commands": ["doc-create"],
    },
    "docx_read": {
        "label": "DOCX read",
        "category": "DOCUMENTS (.docx)",
        "requires": ["python_docx"],
        "commands": ["doc-read", "doc-search", "doc-word-count"],
    },
    "docx_edit": {
        "label": "DOCX edit",
        "category": "DOCUMENTS (.docx)",
        "requires": ["python_docx"],
        "commands": ["doc-append", "doc-replace", "doc-insert", "doc-delete"],
    },
    "docx_tables": {
        "label": "DOCX tables",
        "category": "DOCUMENTS (.docx)",
        "requires": ["python_docx"],
        "commands": ["doc-add-table", "doc-read-tables"],
    },
    "docx_formatting": {
        "label": "DOCX formatting",
        "category": "DOCUMENTS (.docx)",
        "requires": ["python_docx"],
        "commands": [
            "doc-format",
            "doc-layout",
            "doc-normalize-format",
            "doc-formatting-info",
            "doc-font-audit",
            "doc-render-audit",
        ],
    },
    "docx_submission": {
        "label": "DOCX submission readiness",
        "category": "DOCUMENTS (.docx)",
        "requires": ["python_docx"],
        "commands": [
            "doc-inspect-hidden-data",
            "doc-sanitize",
            "doc-preflight",
            "doc-submission-pack",
        ],
    },
    "docx_references": {
        "label": "DOCX references",
        "category": "DOCUMENTS (.docx)",
        "requires": ["python_docx"],
        "commands": ["doc-add-reference", "doc-build-references", "doc-citation-audit"],
    },
    "xlsx_create": {
        "label": "XLSX create",
        "category": "SPREADSHEETS (.xlsx)",
        "requires": ["openpyxl"],
        "commands": ["xlsx-create", "xlsx-write"],
    },
    "xlsx_read": {
        "label": "XLSX read",
        "category": "SPREADSHEETS (.xlsx)",
        "requires": ["openpyxl"],
        "commands": ["xlsx-read", "xlsx-cell-read", "xlsx-range-read", "xlsx-search"],
    },
    "xlsx_edit": {
        "label": "XLSX edit",
        "category": "SPREADSHEETS (.xlsx)",
        "requires": ["openpyxl"],
        "commands": ["xlsx-cell-write", "xlsx-append", "xlsx-delete-rows", "xlsx-delete-cols", "xlsx-format-cells"],
    },
    "xlsx_formulas": {
        "label": "XLSX formulas",
        "category": "SPREADSHEETS (.xlsx)",
        "requires": ["openpyxl"],
        "commands": ["xlsx-formula", "xlsx-formula-audit", "xlsx-calc"],
    },
    "xlsx_charts": {
        "label": "XLSX charts",
        "category": "CHARTS (.xlsx)",
        "requires": ["openpyxl"],
        "commands": ["chart-create", "chart-comparison", "chart-grade-dist", "chart-progress"],
    },
    "xlsx_stats": {
        "label": "XLSX statistics",
        "category": "SPREADSHEETS (.xlsx)",
        "requires": ["openpyxl"],
        "commands": ["xlsx-freq", "xlsx-corr", "xlsx-ttest", "xlsx-mannwhitney", "xlsx-chi2", "xlsx-research-pack"],
    },
    "xlsx_csv": {
        "label": "XLSX CSV import/export",
        "category": "SPREADSHEETS (.xlsx)",
        "requires": ["openpyxl"],
        "commands": ["xlsx-csv-import", "xlsx-csv-export"],
    },
    "pptx_create": {
        "label": "PPTX create",
        "category": "PRESENTATIONS (.pptx)",
        "requires": ["python_pptx"],
        "commands": ["pptx-create", "pptx-add-slide", "pptx-add-bullets"],
    },
    "pptx_read": {
        "label": "PPTX read",
        "category": "PRESENTATIONS (.pptx)",
        "requires": ["python_pptx"],
        "commands": ["pptx-read", "pptx-slide-count", "pptx-list-shapes"],
    },
    "pptx_edit": {
        "label": "PPTX edit",
        "category": "PRESENTATIONS (.pptx)",
        "requires": ["python_pptx"],
        "commands": ["pptx-update-text", "pptx-delete-slide", "pptx-add-textbox", "pptx-modify-shape"],
    },
    "pptx_notes": {
        "label": "PPTX speaker notes",
        "category": "PRESENTATIONS (.pptx)",
        "requires": ["python_pptx"],
        "commands": ["pptx-speaker-notes"],
    },
    "rdf_create": {
        "label": "RDF create/edit",
        "category": "RDF (Knowledge Graphs)",
        "requires": ["rdflib"],
        "commands": ["rdf-create", "rdf-add", "rdf-remove"],
    },
    "rdf_query": {
        "label": "RDF query",
        "category": "RDF (Knowledge Graphs)",
        "requires": ["rdflib"],
        "commands": ["rdf-read", "rdf-query", "rdf-stats", "rdf-namespace"],
    },
    "rdf_validate": {
        "label": "RDF validate",
        "category": "RDF (Knowledge Graphs)",
        "requires": ["pyshacl"],
        "commands": ["rdf-validate"],
    },
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


def get_usage_map() -> Dict[str, str]:
    """Return command names mapped to stable usage strings."""
    return {
        command: command_usage(command)
        for command in sorted(COMMAND_SIGNATURES)
    }


def get_command_metadata() -> Dict[str, Dict[str, object]]:
    """Return additive command metadata keyed by canonical command name."""
    commands: Dict[str, Dict[str, object]] = {}
    for category, signatures in COMMAND_CATEGORIES.items():
        for raw_signature, description in signatures.items():
            command = raw_signature.split()[0]
            commands[command] = {
                "name": command,
                "category": category,
                "signature": command_signature(command),
                "usage": command_usage(command),
                "description": description,
            }
    return {command: commands[command] for command in sorted(commands)}


def build_capability_metadata(
    capabilities: Dict[str, bool]
) -> Dict[str, Dict[str, object]]:
    """Annotate capability booleans with stable dependency/command metadata."""
    metadata: Dict[str, Dict[str, object]] = {}
    for name in sorted(capabilities):
        details = deepcopy(CAPABILITY_DETAILS.get(name, {}))
        details.setdefault("label", name.replace("_", " "))
        details["available"] = bool(capabilities[name])
        metadata[name] = details
    return metadata


def usage_error(command: str) -> Dict[str, str]:
    """Build a standard usage error payload for a command."""
    return {"success": False, "error": command_usage(command)}


def build_help_payload(capabilities: Dict[str, bool]) -> Dict[str, object]:
    """Build the public help payload from the registry."""
    return {
        "success": True,
        "schema_version": CLI_SCHEMA_VERSION,
        "version": VERSION,
        "categories": get_command_categories(),
        "capabilities": dict(capabilities),
        "capability_metadata": build_capability_metadata(capabilities),
        "total_commands": TOTAL_COMMANDS,
        "command_count": TOTAL_COMMANDS,
        "examples": get_help_examples(),
        "category_counts": dict(CATEGORY_COUNTS),
        "commands": get_command_metadata(),
        "usage": get_usage_map(),
    }
