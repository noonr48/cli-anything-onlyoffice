---
name: onlyoffice
version: 4.4.11
author: SLOANE OS
description: 117-command CLI for Documents (.docx), Spreadsheets (.xlsx), Presentations (.pptx), PDFs, and RDF Knowledge Graphs
tags: [productivity, documents, office, onlyoffice, charts, spreadsheets, rdf, apa, pdf, image-extraction, spatial, data-validation]
---

# CLI-Anything OnlyOffice v4.4.11

Programmatic control over Office documents designed for AI agents. Full JSON output. Production-safe with atomic writes, two-layer file locking, and automatic backups.

## What's New in v4.4.11
- **General CLI Split** — alias handling plus the non-prefixed `open/watch/info/editor-* / backup-* / status / help / list` command layer now lives in `core/general_cli.py`, so `core/cli.py` is reduced to bootstrap and modality routing
- **Direct General-Handler Tests** — added dedicated `test_general_cli.py` coverage for alias normalization, editor-session flag parsing, backup-prune parsing, usage errors, and unknown-command fallthrough
- **Carry Forward** — the registry-driven help/status surface, registry-driven usage strings, and all earlier modality/backend splits remain in place from `v4.4.10`, `v4.4.9`, `v4.4.8`, `v4.4.7`, `v4.4.6`, `v4.4.5`, `v4.4.4`, `v4.4.3`, `v4.4.2`, and `v4.4.1`

## Carry Forward from v4.1.0
- **Image Extraction** — pull images out of PDFs, .docx, and .pptx files
- **PDF Rendering** — render PDF pages as PNG/JPG images (PyMuPDF)
- **Spatial Awareness** — list all shapes with exact positions/sizes for layout control
- **Textbox Control** — add/modify textboxes at exact coordinates on slides
- **Slide Preview** — render slides as PNG images via OnlyOffice x2t
- **Data Validation** — Excel-style cell validation (dropdowns, number ranges, text length) + data audit

## CRITICAL: How to Call This Tool

```bash
# Via SLOANE skill router (preferred for subject agents):
cli_anything_run --app onlyoffice <command> [args] --json

# Direct binary (Claude Code / shell):
/home/benbi/cli-anything/onlyoffice/agent-harness/.venv/bin/cli-anything-onlyoffice <command> [args] --json
```

**NEVER do any of these — they will fail:**
- `which onlyoffice-cli` — wrong name
- `pip install python-pptx` / `pip3 install` — already in venv, don't touch system python
- `python3 -c "import pptx"` — wrong interpreter, use venv python
- Installing anything — the tool is already installed

Verify the tool is working:
```bash
cli_anything_run --app onlyoffice status --json
```
Check the `python` field in the response — it must point to `.venv/bin/python3`.

## CRITICAL: Office files are not plain text

`.docx`, `.xlsx`, and `.pptx` are OOXML containers (zip/XML packages), so generic text Read/Write tools will often fail or report them as binary.

**Do not treat that as a capability gap.** Use this tool's semantic commands instead:
- DOCX: `doc-read`, `doc-append`, `doc-replace`, `doc-search`, `doc-read-tables`, `doc-add-image`, `doc-to-pdf`, `doc-preview`, `doc-render-map`
- XLSX: `xlsx-read`, `xlsx-cell-read`, `xlsx-range-read`, `xlsx-write`, `xlsx-cell-write`, `xlsx-preview`
- PPTX: `pptx-read`, `pptx-add-slide`, `pptx-update-text`, `pptx-preview`

If you need the rendered visual layout, use preview/export commands rather than raw file reads.

---

## Document Defaults (doc-create)

| Setting | Value |
|---------|-------|
| Page size | A4 (210 x 297 mm) |
| Margins | 1.0" all sides |
| Font | Calibri 11pt |
| Line spacing | Double (APA 7th) |
| Space after | 0pt |

## Spreadsheet Defaults (xlsx-write)
- Column widths auto-fit to content
- A4 paper size set on all sheets

## Presentation Defaults (pptx-create)
- Slide size: 16:9 widescreen (13.333" x 7.5")

---

## DOCUMENTS (.docx) — 32 commands

### Core CRUD
- `doc-create <file> <title> <content>`
- `doc-read <file>`
- `doc-append <file> <content>`
- `doc-replace <file> <search> <replace>`
- `doc-search <file> <text> [--case-sensitive]`
- `doc-insert <file> <text> <index> [--style <name>]`
- `doc-delete <file> <index>`
- `doc-word-count <file>`

### Formatting
- `doc-format <file> <paragraph_index> [--bold] [--italic] [--underline] [--font-name <n>] [--font-size <n>] [--color <hex>] [--align <left|center|right|justify>]`
- `doc-set-style <file> <index> <style>`
- `doc-list-styles <file>`
- `doc-highlight <file> <text> [--color yellow|cyan|green|pink]`
- `doc-formatting-info <file>`

### Page Layout
- `doc-layout <file> [--size A4|Letter] [--orientation portrait|landscape] [--margin-top <in>] [--margin-bottom <in>] [--margin-left <in>] [--margin-right <in>] [--header <text>] [--page-numbers]`

### Rich Content
- `doc-add-table <file> <headers_csv> <data_csv>` — rows separated by `;`
- `doc-read-tables <file>`
- `doc-add-image <file> <image_path> [--width <inches>] [--caption <text>]`
- `doc-add-hyperlink <file> <text> <url> [--paragraph <index>]`
- `doc-add-page-break <file>`
- `doc-add-list <file> <items> [--type bullet|number]` — items separated by `;`

### Image Extraction
- `doc-extract-images <file> <output_dir> [--format png|jpg] [--prefix <name>]` — extract all embedded images

### Metadata & Annotations
- `doc-set-metadata <file> [--author] [--title] [--subject] [--keywords] [--comments] [--category]`
- `doc-get-metadata <file>`
- `doc-inspect-hidden-data <file>`
- `doc-preflight <file> [--expected-page-size <A4|Letter>] [--expected-font <name>] [--expected-font-size <pt>]`
- `doc-sanitize <file> [output_path] [--remove-comments] [--accept-revisions] [--clear-metadata] [--remove-custom-xml] [--author <a>]`
- `doc-comment <file> <comment> [--paragraph <index>]`

### References (APA 7th)
- `doc-add-reference <file> <ref_json>` — types: journal, book, website, report, chapter
- `doc-build-references <file>`

### Rendered Output
- `doc-to-pdf <file> [output_path]`
- `doc-preview <file> <output_dir> [--pages <range>] [--dpi <n>] [--format png|jpg]`
- `doc-render-map <file>`

---

## SPREADSHEETS (.xlsx) — 37 commands

### Core CRUD
- `xlsx-create <file> [sheet_name]`
- `xlsx-write <file> <headers_csv> <data_csv> [--sheet <name>] [--overwrite] [--coerce-rows] [--text-columns <csv>]`
- `xlsx-read <file> [sheet_name]`
- `xlsx-append <file> <row_data_csv> [--sheet <name>]`
- `xlsx-search <file> <text> [--sheet <name>]`

### Cell & Range
- `xlsx-cell-read <file> <cell> [--sheet <name>]`
- `xlsx-cell-write <file> <cell> <value> [--sheet <name>] [--text]`
- `xlsx-range-read <file> <range> [--sheet <name>]`
- `xlsx-delete-rows <file> <start_row> [count] [--sheet <name>]`
- `xlsx-delete-cols <file> <start_col> [count] [--sheet <name>]`

### Sorting & Filtering
- `xlsx-sort <file> <column> [--sheet <name>] [--desc] [--numeric]`
- `xlsx-filter <file> <column> <op> <value> [--sheet <name>]` — ops: eq|ne|gt|lt|ge|le|contains|startswith|endswith

### Formulas & Stats
- `xlsx-formula <file> <cell> <formula> [--sheet <name>]`
- `xlsx-calc <file> <column> <op> [--sheet <name>] [--include-formulas] [--strict-formulas]` — ops: sum|avg|min|max|all
- `xlsx-formula-audit <file> [--sheet <name>]`
- `xlsx-freq <file> <column> [--sheet <name>] [--valid <csv>]`
- `xlsx-corr <file> <x_col> <y_col> [--sheet <name>] [--method pearson|spearman]`
- `xlsx-ttest <file> <val_col> <grp_col> <group_a> <group_b> [--sheet <name>] [--equal-var]`
- `xlsx-mannwhitney <file> <val_col> <grp_col> <group_a> <group_b> [--sheet <name>]`
- `xlsx-chi2 <file> <row_col> <col_col> [--sheet <name>] [--row-valid <csv>] [--col-valid <csv>]`
- `xlsx-text-extract <file> <column> [--sheet <name>] [--limit <n>] [--min-length <n>]`
- `xlsx-text-keywords <file> <column> [--sheet <name>] [--top <n>]`
- `xlsx-research-pack <file> [--sheet <name>] [--profile hlth3112]`

### Sheet Management
- `xlsx-sheet-list <file>`
- `xlsx-sheet-add <file> <name> [--position <n>]`
- `xlsx-sheet-delete <file> <name>`
- `xlsx-sheet-rename <file> <old_name> <new_name>`

### Cell Formatting
- `xlsx-merge-cells <file> <range> [--sheet <name>]`
- `xlsx-unmerge-cells <file> <range> [--sheet <name>]`
- `xlsx-format-cells <file> <range> [--sheet <name>] [--bold] [--italic] [--font-name <n>] [--font-size <n>] [--color <hex>] [--bg-color <hex>] [--number-format <fmt>] [--wrap] [--align <left|center|right>]`

### CSV
- `xlsx-csv-import <xlsx_file> <csv_file> [--sheet <name>] [--delimiter <char>]`
- `xlsx-csv-export <xlsx_file> <csv_file> [--sheet <name>] [--delimiter <char>]`

### Data Validation
- `xlsx-add-validation <file> <range> <type> [--operator <op>] [--formula1 <v>] [--formula2 <v>] [--sheet <name>] [--error <msg>] [--prompt <msg>] [--error-style stop|warning|information] [--no-blank]`
  - Types: `list`, `whole`, `decimal`, `date`, `time`, `textLength`, `custom`
  - Operators: `between`, `notBetween`, `equal`, `notEqual`, `lessThan`, `lessThanOrEqual`, `greaterThan`, `greaterThanOrEqual`
- `xlsx-add-dropdown <file> <range> <options_csv> [--sheet <name>] [--prompt <msg>] [--error <msg>]` — shortcut for dropdown list
- `xlsx-list-validations <file> [--sheet <name>]` — list all validation rules
- `xlsx-remove-validation <file> [--range <range>] [--all] [--sheet <name>]` — remove rules
- `xlsx-validate-data <file> [--sheet <name>] [--max-rows <n>]` — audit existing data against rules (pass/fail per cell)

**Validation Workflow:**
1. Write data with `xlsx-write`
2. Add rules with `xlsx-add-validation` or `xlsx-add-dropdown`
3. Run `xlsx-validate-data` to audit — returns every failing cell with reason
4. Fix failing cells with `xlsx-cell-write`
5. Re-run `xlsx-validate-data` to confirm all clean

---

## CHARTS (.xlsx) — 4 commands

Types: bar, column, bar_horizontal, line, pie, scatter

- `chart-create <file> <type> <data_range> <categories_range> <title> [--sheet <name>] [--output-sheet <name>] [--x-label <text>] [--y-label <text>] [--labels] [--no-legend] [--legend-pos right|top|bottom|left] [--colors <hex,hex>]`
- `chart-comparison <file> <type> <title> [--sheet <name>] [--start-row <n>] [--cat-col <n>] [--value-cols <n,n,n>] [--output <cell>] [--labels]`
- `chart-grade-dist <file> <grade_col> <title> [--sheet <name>] [--output <cell>]`
- `chart-progress <file> <student_col> <grade_col> <title> [--sheet <name>] [--output <cell>] [--labels]`

---

## PRESENTATIONS (.pptx) — 16 commands

### Core
- `pptx-create <file> <title> [subtitle]` — 16:9 widescreen by default
- `pptx-add-slide <file> <title> [content] [layout]` — layouts: title_only|content|blank|two_content|comparison
- `pptx-add-bullets <file> <title> <bullets>` — bullets separated by `\n`
- `pptx-add-table <file> <title> <headers_csv> <data_csv> [--coerce-rows]`
- `pptx-add-image <file> <title> <image_path>`
- `pptx-read <file>`
- `pptx-slide-count <file>`
- `pptx-delete-slide <file> <index>`
- `pptx-speaker-notes <file> <slide_index> [notes_text]`
- `pptx-update-text <file> <slide_index> [--title <t>] [--body <b>]`

### Image Extraction
- `pptx-extract-images <file> <output_dir> [--slide <index>] [--format png|jpg] [--prefix <name>]` — extract all images from slides

### Spatial Awareness & Layout Control
- `pptx-list-shapes <file> [--slide <index>]` — list ALL shapes with exact position, size, text, type
- `pptx-add-textbox <file> <slide_index> <text> [--left <in>] [--top <in>] [--width <in>] [--height <in>] [--font-size <pt>] [--font-name <name>] [--bold] [--italic] [--color <hex>] [--align <left|center|right>]`
- `pptx-modify-shape <file> <slide_index> <shape_name> [--left <in>] [--top <in>] [--width <in>] [--height <in>] [--text <text>] [--font-size <pt>] [--rotation <deg>]`

### Visual Preview
- `pptx-preview <file> <output_dir> [--slide <index>] [--dpi <n>]` — render slides as PNG (requires OnlyOffice Docker)

**Spatial Workflow (recommended for quality presentations):**
1. Create slides with content
2. Run `pptx-list-shapes` to see exact positions of all elements
3. Use `pptx-modify-shape` to fix overlaps or reposition elements
4. Use `pptx-add-textbox` for custom positioned text
5. Run `pptx-preview` to render a visual check — view the PNG to verify layout
6. Iterate if needed

**Slide coordinate system:** Origin (0,0) = top-left. Slide is 13.333" wide x 7.5" tall (16:9).

---

## PDF (.pdf) — 6 commands

- `pdf-extract-images <file> <output_dir> [--format png|jpg] [--pages <range>]` — extract embedded image objects (PyMuPDF)
- `pdf-page-to-image <file> <output_dir> [--pages <range>] [--dpi <n>] [--format png|jpg]` — render full pages as images
- `pdf-read-blocks <file> [--pages <range>] [--no-spans] [--no-images] [--include-empty]` — read native PDF blocks/lines/spans with bbox metadata
- `pdf-search-blocks <file> <query> [--pages <range>] [--case-sensitive] [--no-spans]` — search exact PDF blocks/spans and return native anchors
- `pdf-inspect-hidden-data <file>` — inspect metadata, XMP/XML metadata, annotations, embedded files, form usage, and page-size consistency
- `pdf-sanitize <file> [output_path] [--clear-metadata] [--remove-xml-metadata] [--author <a>]` — clear PDF metadata/XMP for submission

Page ranges: `0-3` (pages 0 through 3), `1,3,5` (specific pages), omit for all pages.
Default DPI: 150. Use 300 for print quality.

---

## RDF KNOWLEDGE GRAPHS — 10 commands

- `rdf-create <file> [--base <uri>] [--format turtle|xml|n3|json-ld] [--prefix <p>=<uri>]`
- `rdf-read <file> [--limit <n>]`
- `rdf-add <file> <subject> <predicate> <object> [--type uri|literal|bnode] [--lang <tag>] [--datatype <xsd_uri>]`
- `rdf-remove <file> [--subject <uri>] [--predicate <uri>] [--object <value>] [--type uri|literal|bnode] [--lang <tag>] [--datatype <xsd_uri>]`
- `rdf-query <file> <sparql_query> [--limit <n>]`
- `rdf-export <file> <output_file> [--format <format>]`
- `rdf-merge <file_a> <file_b> [--output <file>] [--format <f>]`
- `rdf-stats <file>`
- `rdf-namespace <file> [<prefix> <uri>]`
- `rdf-validate <data_file> <shapes_file>`

---

## GENERAL — 11 commands

- `list` — List recent office files
- `open <file> [gui|web]`
- `watch <file> [gui|web]`
- `info <file>`
- `editor-session <file> [--open] [--wait <sec>] [--activate]`
- `editor-capture <file> <output_image> [--backend auto|desktop|rendered] [--page <n>] [--range <A1:D20>] [--slide <n>] [--zoom-reset] [--zoom-in <n>] [--zoom-out <n>] [--crop x,y,w,h]`
- `status` — Check installation (includes `python` field showing active interpreter)
- `help`
- `backup-list <file> [--limit <n>]`
- `backup-prune [--file <f>] [--keep <n>] [--days <n>]`
- `backup-restore <file> [--latest] [--dry-run]`

---

## Agent Rules

1. **Always `--json`** — every response is `{"success": true/false, ...}`
2. **Check `success` first**
3. **Run `status --json` at session start** — verify `python` field = `.venv/bin/python3`
4. **Never install anything** — tool is fully installed, all deps in venv
5. **Never use system python3/pip3/pip**
6. **Backups are automatic** — `backup-restore --latest` on any write error
7. **Use `pptx-list-shapes` before modifying slides** — know exact positions to avoid overlaps
8. **Use `pptx-preview` after building slides** — visually verify the layout looks correct

---

**Last Updated:** 2026-04-21
**Version:** 4.4.11
