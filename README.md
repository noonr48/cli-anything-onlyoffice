# CLI-Anything OnlyOffice v4.0.2

> Programmatic control over Documents (.docx), Spreadsheets (.xlsx), Presentations (.pptx), and RDF Knowledge Graphs — designed for AI agents.

**90 commands. 4 modes. Full JSON output. Production-safe.**

Part of the [SLOANE OS](https://github.com/sloane-os) agent stack. Agents call this via `cli_anything_run(tool='onlyoffice', ...)`.

---

## Installation

```bash
git clone https://github.com/noonr48/cli-anything-onlyoffice.git
cd cli-anything-onlyoffice

python3 -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate

pip install -e .
# or with optional SHACL validation support
pip install -e ".[shacl]"

# Verify
cli-anything-onlyoffice status --json
```

To re-activate in a future shell session:
```bash
source /path/to/cli-anything-onlyoffice/.venv/bin/activate
cli-anything-onlyoffice status --json
```

When calling from SLOANE OS or another agent, invoke via the venv binary directly so you don't depend on the shell's active environment:
```bash
/path/to/cli-anything-onlyoffice/.venv/bin/cli-anything-onlyoffice status --json
```

### Dependencies

| Library | Purpose | Required |
|---------|---------|---------|
| `python-docx>=1.1.0` | .docx manipulation | Core |
| `openpyxl>=3.1.2` | .xlsx manipulation + charts | Core |
| `python-pptx>=0.6.23` | .pptx manipulation | Core |
| `rdflib>=7.0.0` | RDF graph / SPARQL | Core |
| `lxml>=4.9.0` | XML parsing for OOXML | Core |
| `scipy>=1.11.0` | Statistical tests | Core |
| `pyshacl>=0.25.0` | SHACL validation | Optional (`[shacl]`) |

---

## Architecture

```
cli-anything-onlyoffice <command> [args] [--json]
        │
        ▼
  core/cli.py          ← Arg parsing + dispatch router
        │
        ▼
utils/docserver.py     ← Backend engine (4,859 lines)
  ├── Documents        ← python-docx wrapper
  ├── Spreadsheets     ← openpyxl wrapper + scipy stats
  ├── Presentations    ← python-pptx wrapper
  └── RDF              ← rdflib 7 wrapper
```

### Production-Safety Features

Every write operation goes through four safeguards:

- **Atomic saves** — writes to a temp file, then `os.replace()`. No partial writes.
- **Two-layer file locking** — `threading.Lock` (per-path, intra-process) + `fcntl.flock(LOCK_EX)` (cross-process). Both layers are required: `fcntl.flock` is per-process on Linux and does not serialise threads within the same process.
- **Automatic backup snapshots** — pre-save copy written to `~/.cli-anything/backups/` before every mutation.

---

## Universal Flags

| Flag | Effect |
|------|--------|
| `--json` | Machine-readable JSON output (always use this with agents) |

All responses have `{"success": true, ...}` or `{"success": false, "error": "..."}`.

---

## Mode 1 — Documents (.docx)

**25 commands** — full lifecycle from creation to APA references.

### Document Defaults

Every document created with `doc-create` is pre-configured for academic/APA use:

| Setting | Value |
|---------|-------|
| Page size | A4 (210 × 297 mm) |
| Margins | 1.0" all sides (top, bottom, left, right) |
| Font | Calibri 11pt |
| Line spacing | Double (APA 7th edition) |
| Space after paragraph | 0pt |

These defaults apply to all body paragraphs including those added via `doc-append`. Use `doc-layout` to override page size or margins on an existing file.

### Core CRUD

#### `doc-create <file> <title> <content>`
Create a new .docx document.

```bash
cli-anything-onlyoffice doc-create /tmp/essay.docx "My Essay" "Introduction paragraph here" --json
```
```json
{"success": true, "file": "/tmp/essay.docx", "title": "My Essay", "size": 8192}
```

#### `doc-read <file>`
Read all content — paragraphs, tables, full text.

```bash
cli-anything-onlyoffice doc-read /tmp/essay.docx --json
```
```json
{
  "success": true,
  "file": "/tmp/essay.docx",
  "paragraphs": ["Introduction paragraph here"],
  "paragraph_count": 1,
  "full_text": "Introduction paragraph here"
}
```

#### `doc-append <file> <content>`
Append a paragraph to the end.

```bash
cli-anything-onlyoffice doc-append /tmp/essay.docx "Body paragraph with more detail." --json
```

#### `doc-replace <file> <search> <replace>`
Find and replace text (cross-run safe, preserves formatting).

```bash
cli-anything-onlyoffice doc-replace /tmp/essay.docx "draft" "final version" --json
```

#### `doc-search <file> <text> [--case-sensitive]`
Search paragraphs and tables for text, returns match locations.

```bash
cli-anything-onlyoffice doc-search /tmp/essay.docx "introduction" --json
cli-anything-onlyoffice doc-search /tmp/essay.docx "Introduction" --case-sensitive --json
```

#### `doc-insert <file> <text> <index> [--style <name>]`
Insert a paragraph at a specific position (0-based index).

```bash
cli-anything-onlyoffice doc-insert /tmp/essay.docx "New first paragraph" 0 --style "Heading 1" --json
cli-anything-onlyoffice doc-insert /tmp/essay.docx "A middle paragraph" 2 --json
```

#### `doc-delete <file> <index>`
Delete a paragraph by index (0-based).

```bash
cli-anything-onlyoffice doc-delete /tmp/essay.docx 3 --json
```

#### `doc-word-count <file>`
Word, character, and paragraph counts.

```bash
cli-anything-onlyoffice doc-word-count /tmp/essay.docx --json
```
```json
{"success": true, "words": 450, "characters": 2780, "paragraphs": 8}
```

---

### Formatting

#### `doc-format <file> <paragraph_index> [options]`
Apply rich formatting to a paragraph.

Options: `--bold`, `--italic`, `--underline`, `--font-name <name>`, `--font-size <pts>`, `--color <RRGGBB>`, `--align <left|center|right|justify>`

```bash
cli-anything-onlyoffice doc-format /tmp/essay.docx 0 --bold --font-size 18 --align center --json
cli-anything-onlyoffice doc-format /tmp/essay.docx 1 --italic --color FF0000 --json
```

#### `doc-set-style <file> <index> <style>`
Set a paragraph style by name.

```bash
cli-anything-onlyoffice doc-set-style /tmp/essay.docx 0 "Heading 1" --json
cli-anything-onlyoffice doc-set-style /tmp/essay.docx 1 "Normal" --json
```

Style names: `Heading 1`, `Heading 2`, `Heading 3`, `Normal`, `Title`, `Subtitle`, `Quote`, etc.

#### `doc-list-styles <file>`
List all available paragraph/character styles in the document.

```bash
cli-anything-onlyoffice doc-list-styles /tmp/essay.docx --json
```

#### `doc-highlight <file> <search_text> [--color <name>]`
Highlight matching text runs. Colors: `yellow` (default), `cyan`, `green`, `pink`, etc.

```bash
cli-anything-onlyoffice doc-highlight /tmp/essay.docx "important term" --color yellow --json
```

#### `doc-formatting-info <file>`
Inspect all paragraph and section formatting details.

```bash
cli-anything-onlyoffice doc-formatting-info /tmp/essay.docx --json
```

---

### Page Layout

#### `doc-layout <file> [options]`
Set page orientation, margins, header, and page numbers.

Options: `--orientation portrait|landscape`, `--margin-top <in>`, `--margin-bottom <in>`, `--margin-left <in>`, `--margin-right <in>`, `--header <text>`, `--page-numbers`

```bash
cli-anything-onlyoffice doc-layout /tmp/essay.docx --orientation landscape --json
cli-anything-onlyoffice doc-layout /tmp/report.docx \
  --margin-top 1.0 --margin-bottom 1.0 --margin-left 1.25 --margin-right 1.25 \
  --header "Research Report 2026" --page-numbers --json
```

---

### Rich Content

#### `doc-add-table <file> <headers_csv> <data_csv>`
Add a table. Rows are separated by `;`.

```bash
cli-anything-onlyoffice doc-add-table /tmp/essay.docx \
  "Name,Score,Grade" \
  "Alice,92,A;Bob,78,B;Charlie,85,B+" --json
```

#### `doc-read-tables <file>`
Read all tables from the document.

```bash
cli-anything-onlyoffice doc-read-tables /tmp/essay.docx --json
```
```json
{
  "success": true,
  "tables": [
    {"rows": [["Name", "Score"], ["Alice", "92"]], "row_count": 2, "col_count": 2}
  ],
  "table_count": 1
}
```

#### `doc-add-image <file> <image_path> [--width <inches>] [--caption <text>]`
Embed an image with optional caption.

```bash
cli-anything-onlyoffice doc-add-image /tmp/essay.docx /tmp/figure1.png --width 5.0 --caption "Figure 1: Overview" --json
```

#### `doc-add-hyperlink <file> <text> <url> [--paragraph <index>]`
Insert a hyperlink. Use `--paragraph -1` (default) to add to a new paragraph.

```bash
cli-anything-onlyoffice doc-add-hyperlink /tmp/essay.docx "Click here" "https://example.com" --json
cli-anything-onlyoffice doc-add-hyperlink /tmp/essay.docx "Source" "https://doi.org/..." --paragraph 3 --json
```

#### `doc-add-page-break <file>`
Insert a page break at the end of the document.

```bash
cli-anything-onlyoffice doc-add-page-break /tmp/essay.docx --json
```

#### `doc-add-list <file> <items> [--type bullet|number]`
Add a bulleted or numbered list. Items separated by `;`.

```bash
cli-anything-onlyoffice doc-add-list /tmp/essay.docx "First point;Second point;Third point" --type bullet --json
cli-anything-onlyoffice doc-add-list /tmp/essay.docx "Step one;Step two;Step three" --type number --json
```

---

### Metadata & Annotations

#### `doc-set-metadata <file> [options]`
Set document properties: `--author`, `--title`, `--subject`, `--keywords`, `--comments`, `--category`

```bash
cli-anything-onlyoffice doc-set-metadata /tmp/essay.docx \
  --author "SLOANE Agent" --title "Research Essay" --keywords "health,survey,2026" --json
```

#### `doc-get-metadata <file>`
Read all document properties.

```bash
cli-anything-onlyoffice doc-get-metadata /tmp/essay.docx --json
```
```json
{"success": true, "author": "SLOANE Agent", "title": "Research Essay", "created": "2026-04-07T10:00:00"}
```

#### `doc-comment <file> <comment> [--paragraph <index>]`
Attach an OOXML comment annotation to a paragraph.

```bash
cli-anything-onlyoffice doc-comment /tmp/essay.docx "Review this section" --paragraph 2 --json
```

---

### References (APA 7th)

#### `doc-add-reference <file> <ref_json>`
Add a reference to the sidecar `.refs.json` file.

```bash
cli-anything-onlyoffice doc-add-reference /tmp/essay.docx \
  '{"author": "Smith, J.", "year": "2023", "title": "Health Outcomes", "source": "Journal of Health", "type": "journal", "doi": "10.1234/jh.2023"}' --json
```

Supported types: `journal`, `book`, `website`, `report`, `chapter`

#### `doc-build-references <file>`
Build a formatted APA 7th edition References section from the sidecar `.refs.json` and append it to the document.

```bash
cli-anything-onlyoffice doc-build-references /tmp/essay.docx --json
```

---

## Mode 2 — Spreadsheets (.xlsx)

**32 commands** — cell-level access, sheets, stats, CSV I/O, charts.

### Spreadsheet Defaults

Every sheet written with `xlsx-write` automatically:
- **Auto-fits column widths** to content (min 12 chars, max 50 chars)
- **Sets A4 paper size** (paperSize=9) for printing

### Core CRUD

#### `xlsx-create <file> [sheet_name]`
Create a new empty spreadsheet.

```bash
cli-anything-onlyoffice xlsx-create /tmp/data.xlsx "Grades" --json
```

#### `xlsx-write <file> <headers_csv> <data_csv> [options]`
Write headers and rows. Row values separated by `,`, rows by `;`. Values starting with `=` become formulas.

Options: `--sheet <name>`, `--overwrite` (replace entire workbook), `--coerce-rows` (pad/trim row lengths), `--text-columns <A,B>` (force columns as text)

```bash
cli-anything-onlyoffice xlsx-write /tmp/grades.xlsx \
  "Student,Assignment1,Assignment2,Total" \
  "Alice,85,90,=B2+C2;Bob,78,82,=B3+C3;Charlie,92,88,=B4+C4" \
  --sheet Grades --json
```

#### `xlsx-read <file> [sheet_name]`
Read all data from a sheet (or all sheets if none specified).

```bash
cli-anything-onlyoffice xlsx-read /tmp/grades.xlsx Grades --json
cli-anything-onlyoffice xlsx-read /tmp/grades.xlsx --json  # reads all sheets
```

#### `xlsx-append <file> <row_data_csv> [--sheet <name>]`
Append a row to a sheet.

```bash
cli-anything-onlyoffice xlsx-append /tmp/grades.xlsx "Diana,91,87" --sheet Grades --json
```

#### `xlsx-search <file> <text> [--sheet <name>]`
Search for text across cells, returns exact cell addresses.

```bash
cli-anything-onlyoffice xlsx-search /tmp/grades.xlsx "Alice" --json
```
```json
{"success": true, "matches": [{"sheet": "Grades", "cell": "A2", "value": "Alice"}], "count": 1}
```

---

### Cell & Range Operations

#### `xlsx-cell-read <file> <cell> [--sheet <name>]`
Read the value of a single cell.

```bash
cli-anything-onlyoffice xlsx-cell-read /tmp/grades.xlsx B2 --sheet Grades --json
```
```json
{"success": true, "cell": "B2", "value": 85, "type": "int"}
```

#### `xlsx-cell-write <file> <cell> <value> [--sheet <name>] [--text]`
Write a value to a single cell. `--text` forces the value to be stored as text (not parsed as number/formula).

```bash
cli-anything-onlyoffice xlsx-cell-write /tmp/grades.xlsx C2 95 --sheet Grades --json
cli-anything-onlyoffice xlsx-cell-write /tmp/grades.xlsx A1 "Student Name" --text --json
```

#### `xlsx-range-read <file> <range> [--sheet <name>]`
Read a rectangular range of cells.

```bash
cli-anything-onlyoffice xlsx-range-read /tmp/grades.xlsx A1:D4 --sheet Grades --json
```
```json
{"success": true, "range": "A1:D4", "data": [["Student","A1","A2","Total"], ["Alice",85,90,175]]}
```

#### `xlsx-delete-rows <file> <start_row> [count] [--sheet <name>]`
Delete rows (1-indexed). `count` defaults to 1.

```bash
cli-anything-onlyoffice xlsx-delete-rows /tmp/grades.xlsx 3 --json  # delete row 3
cli-anything-onlyoffice xlsx-delete-rows /tmp/grades.xlsx 3 2 --json  # delete rows 3-4
```

#### `xlsx-delete-cols <file> <start_col> [count] [--sheet <name>]`
Delete columns (1-indexed).

```bash
cli-anything-onlyoffice xlsx-delete-cols /tmp/grades.xlsx 4 --json  # delete column 4 (D)
```

---

### Sorting & Filtering

#### `xlsx-sort <file> <column> [--sheet <name>] [--desc] [--numeric]`
Sort data by column, preserving the header row. Column can be letter (`B`) or name.

```bash
cli-anything-onlyoffice xlsx-sort /tmp/grades.xlsx B --sheet Grades --desc --numeric --json
```

#### `xlsx-filter <file> <column> <op> <value> [--sheet <name>]`
Filter rows by condition. Returns matching rows.

Operators: `eq`, `ne`, `gt`, `lt`, `ge`, `le`, `contains`, `startswith`, `endswith`

```bash
cli-anything-onlyoffice xlsx-filter /tmp/grades.xlsx B gt 80 --sheet Grades --json
cli-anything-onlyoffice xlsx-filter /tmp/grades.xlsx A contains "li" --json
```
```json
{"success": true, "rows": [["Alice", 85, 90]], "count": 1, "column": "B", "op": "gt", "value": "80"}
```

---

### Formulas & Calculations

#### `xlsx-formula <file> <cell> <formula> [--sheet <name>]`
Write a formula to a cell.

```bash
cli-anything-onlyoffice xlsx-formula /tmp/grades.xlsx D2 "=AVERAGE(B2:C2)" --json
cli-anything-onlyoffice xlsx-formula /tmp/grades.xlsx E2 "=IF(D2>=85,\"A\",\"B\")" --json
```

#### `xlsx-calc <file> <column> <operation> [--sheet <name>] [--include-formulas]`
Column statistics. Operations: `sum`, `avg`, `min`, `max`, `all`

```bash
cli-anything-onlyoffice xlsx-calc /tmp/grades.xlsx B avg --sheet Grades --json
cli-anything-onlyoffice xlsx-calc /tmp/grades.xlsx B all --json
```
```json
{"success": true, "column": "B", "count": 3, "sum": 255, "average": 85.0, "min": 78, "max": 92}
```

#### `xlsx-formula-audit <file> [--sheet <name>]`
Audit formula complexity and risk for production safety.

```bash
cli-anything-onlyoffice xlsx-formula-audit /tmp/data.xlsx --json
```

---

### Sheet Management

#### `xlsx-sheet-list <file>`
List all sheets with row/column counts.

```bash
cli-anything-onlyoffice xlsx-sheet-list /tmp/grades.xlsx --json
```
```json
{"success": true, "sheets": [{"name": "Grades", "rows": 4, "cols": 4}], "count": 1}
```

#### `xlsx-sheet-add <file> <name> [--position <n>]`
Add a new sheet.

```bash
cli-anything-onlyoffice xlsx-sheet-add /tmp/grades.xlsx "Charts" --json
cli-anything-onlyoffice xlsx-sheet-add /tmp/grades.xlsx "Summary" --position 0 --json
```

#### `xlsx-sheet-delete <file> <name>`
Delete a sheet by name.

```bash
cli-anything-onlyoffice xlsx-sheet-delete /tmp/grades.xlsx "OldSheet" --json
```

#### `xlsx-sheet-rename <file> <old_name> <new_name>`
Rename a sheet.

```bash
cli-anything-onlyoffice xlsx-sheet-rename /tmp/grades.xlsx "Sheet1" "Grades" --json
```

---

### Cell Formatting

#### `xlsx-merge-cells <file> <range> [--sheet <name>]`
Merge a cell range.

```bash
cli-anything-onlyoffice xlsx-merge-cells /tmp/grades.xlsx A1:D1 --json
```

#### `xlsx-unmerge-cells <file> <range> [--sheet <name>]`
Unmerge a previously merged range.

```bash
cli-anything-onlyoffice xlsx-unmerge-cells /tmp/grades.xlsx A1:D1 --json
```

#### `xlsx-format-cells <file> <range> [options] [--sheet <name>]`
Apply rich formatting to a cell range.

Options: `--bold`, `--italic`, `--wrap`, `--font-name <name>`, `--font-size <pts>`, `--color <RRGGBB>`, `--bg-color <RRGGBB>`, `--number-format <fmt>`, `--align <left|center|right>`

```bash
# Bold white text on blue header
cli-anything-onlyoffice xlsx-format-cells /tmp/grades.xlsx A1:D1 \
  --bold --color FFFFFF --bg-color 4472C4 --json

# Currency format
cli-anything-onlyoffice xlsx-format-cells /tmp/budget.xlsx B2:B100 \
  --number-format '"$"#,##0.00' --json
```

---

### CSV Import/Export

#### `xlsx-csv-import <xlsx_file> <csv_file> [--sheet <name>] [--delimiter <char>]`
Import a CSV file into a sheet (replaces sheet content).

```bash
cli-anything-onlyoffice xlsx-csv-import /tmp/data.xlsx /tmp/raw.csv --sheet Imported --json
cli-anything-onlyoffice xlsx-csv-import /tmp/data.xlsx /tmp/european.csv --delimiter ";" --json
```

#### `xlsx-csv-export <xlsx_file> <csv_file> [--sheet <name>] [--delimiter <char>]`
Export a sheet to CSV.

```bash
cli-anything-onlyoffice xlsx-csv-export /tmp/grades.xlsx /tmp/grades.csv --sheet Grades --json
```

---

### Statistical Analysis

All statistical commands return APA-formatted results with effect sizes and confidence intervals where applicable.

#### `xlsx-freq <file> <column> [--sheet <name>] [--valid <values_csv>]`
Frequency table with percentages.

```bash
cli-anything-onlyoffice xlsx-freq /tmp/survey.xlsx C --sheet Sheet0 \
  --valid "Strongly Agree,Agree,Neutral,Disagree,Strongly Disagree" --json
```
```json
{
  "success": true,
  "frequencies": {"Agree": 45, "Strongly Agree": 20, "Neutral": 15},
  "percentages": {"Agree": 54.9, "Strongly Agree": 24.4, "Neutral": 18.3},
  "n": 82
}
```

#### `xlsx-corr <file> <x_col> <y_col> [--sheet <name>] [--method pearson|spearman]`
Correlation test with APA output.

```bash
cli-anything-onlyoffice xlsx-corr /tmp/data.xlsx B C --sheet Sheet0 --method pearson --json
```
```json
{
  "success": true, "r": 0.742, "p_value": 0.001,
  "significant": true, "apa": "r(45) = .742, p < .001"
}
```

#### `xlsx-ttest <file> <value_col> <group_col> <group_a> <group_b> [options]`
Independent samples t-test (Welch default). Includes Cohen's d.

```bash
cli-anything-onlyoffice xlsx-ttest /tmp/data.xlsx B A Male Female \
  --sheet Sheet0 --json
```
```json
{
  "success": true, "t": 2.34, "p_value": 0.023, "df": 78,
  "cohens_d": 0.52, "significant": true,
  "apa": "t(78) = 2.34, p = .023, d = 0.52"
}
```

#### `xlsx-mannwhitney <file> <value_col> <group_col> <group_a> <group_b> [--sheet <name>]`
Non-parametric Mann-Whitney U test.

```bash
cli-anything-onlyoffice xlsx-mannwhitney /tmp/data.xlsx B A GroupX GroupY --json
```

#### `xlsx-chi2 <file> <row_col> <col_col> [options]`
Chi-square test of independence with Cramér's V effect size.

```bash
cli-anything-onlyoffice xlsx-chi2 /tmp/survey.xlsx C D --sheet Sheet0 \
  --row-valid "Yes,No" --col-valid "Male,Female" --json
```
```json
{
  "success": true, "chi2": 5.84, "p_value": 0.016, "df": 1,
  "cramers_v": 0.27, "apa": "χ²(1) = 5.84, p = .016, V = .27"
}
```

#### `xlsx-text-extract <file> <column> [options]`
Extract open-text responses for qualitative coding.

```bash
cli-anything-onlyoffice xlsx-text-extract /tmp/survey.xlsx E --sheet Sheet0 \
  --limit 50 --min-length 20 --json
```

#### `xlsx-text-keywords <file> <column> [options]`
Generate keyword frequency summary from text responses.

```bash
cli-anything-onlyoffice xlsx-text-keywords /tmp/survey.xlsx E --top 15 --json
```

#### `xlsx-research-pack <file> [options]`
Bundled research analysis pack. Runs freq tables, t-tests, chi-square, and correlations in one shot.

```bash
cli-anything-onlyoffice xlsx-research-pack /tmp/survey.xlsx \
  --sheet Sheet0 --profile hlth3112 --json
```

---

## Mode 3 — Charts (.xlsx)

**4 commands** — embedded charts rendered directly in the workbook.

Chart types: `bar`, `column`, `bar_horizontal`, `line`, `pie`, `scatter`

#### `chart-create <file> <type> <data_range> <categories_range> <title> [options]`
Create a chart from explicit cell ranges.

Options: `--sheet <name>`, `--output-sheet <name>`, `--x-label <text>`, `--y-label <text>`, `--labels`, `--no-legend`, `--legend-pos right|top|bottom|left`, `--colors <RRGGBB,RRGGBB>`

```bash
# Bar chart
cli-anything-onlyoffice chart-create /tmp/grades.xlsx bar B2:D4 A2:A4 "Assignment Comparison" \
  --output-sheet Charts --x-label "Student" --y-label "Score" --labels --json

# Line chart with custom colors
cli-anything-onlyoffice chart-create /tmp/grades.xlsx line B2:D10 A2:A10 "Score Trend" \
  --colors FF0000,00BB00,0000FF --json
```

#### `chart-comparison <file> <type> <title> [options]`
Smart comparison chart — auto-detects series from structured data layout.

Options: `--sheet <name>`, `--start-row <n>`, `--start-col <n>`, `--cats <n>`, `--series <n>`, `--cat-col <n>`, `--value-cols <n,n,n>`, `--output <cell>`, `--labels`, `--no-legend`

```bash
cli-anything-onlyoffice chart-comparison /tmp/grades.xlsx bar "Assignment Trends" \
  --start-row 2 --cat-col 1 --value-cols 2,3,4 --output A10 --labels --json
```

#### `chart-grade-dist <file> <grade_col> <title> [--sheet <name>] [--output <cell>]`
Auto-generate a pie chart from grade distribution in a column.

```bash
cli-anything-onlyoffice chart-grade-dist /tmp/grades.xlsx B "Grade Distribution" \
  --output F2 --json
```
```json
{"success": true, "chart_type": "pie", "distribution": {"A": 2, "B": 2, "C": 1}, "total_grades": 5}
```

#### `chart-progress <file> <student_col> <grade_col> <title> [options]`
Horizontal bar chart of individual grades.

Options: `--sheet <name>`, `--output <cell>`, `--labels`, `--no-labels`

```bash
cli-anything-onlyoffice chart-progress /tmp/grades.xlsx A B "Student Grades" \
  --output D2 --labels --json
```

---

## Mode 4 — Presentations (.pptx)

**10 commands** — full slide lifecycle.

#### `pptx-create <file> <title> [subtitle]`
Create a new presentation with a title slide. Slide size is **16:9 widescreen (13.333" × 7.5")** — the modern standard for PowerPoint and OnlyOffice.

```bash
cli-anything-onlyoffice pptx-create /tmp/lecture.pptx "Biology 101" "Introduction to Cell Structure" --json
```

#### `pptx-add-slide <file> <title> [content] [layout]`
Add a slide. Layouts: `title_only`, `content`, `blank`, `two_content`, `comparison`

```bash
cli-anything-onlyoffice pptx-add-slide /tmp/lecture.pptx "Agenda" "Topics we'll cover today" content --json
```

#### `pptx-add-bullets <file> <title> <bullets>`
Add a bullet-point slide. Separate bullets with `\n`.

```bash
cli-anything-onlyoffice pptx-add-bullets /tmp/lecture.pptx "Learning Objectives" \
  "Understand cell theory\nIdentify organelles\nExplain cellular functions" --json
```

#### `pptx-add-table <file> <title> <headers_csv> <data_csv> [--coerce-rows]`
Add a table slide. Rows separated by `;`.

```bash
cli-anything-onlyoffice pptx-add-table /tmp/lecture.pptx "Cell Types" \
  "Type,Nucleus,Size,Examples" \
  "Prokaryotic,No,1-5µm,Bacteria;Eukaryotic,Yes,10-100µm,Animals" --json
```

#### `pptx-add-image <file> <title> <image_path>`
Add an image slide.

```bash
cli-anything-onlyoffice pptx-add-image /tmp/lecture.pptx "Cell Diagram" /tmp/cell.png --json
```

#### `pptx-read <file>`
Read all slides — titles, content, notes, layouts.

```bash
cli-anything-onlyoffice pptx-read /tmp/lecture.pptx --json
```
```json
{
  "success": true,
  "slides": [
    {"index": 0, "title": "Biology 101", "content": "Introduction...", "notes": ""}
  ],
  "slide_count": 1
}
```

#### `pptx-slide-count <file>`
Get slide count and all slide titles.

```bash
cli-anything-onlyoffice pptx-slide-count /tmp/lecture.pptx --json
```
```json
{"success": true, "count": 5, "titles": ["Biology 101", "Agenda", "Cell Types", ...]}
```

#### `pptx-delete-slide <file> <index>`
Delete a slide by 0-based index.

```bash
cli-anything-onlyoffice pptx-delete-slide /tmp/lecture.pptx 2 --json
```

#### `pptx-speaker-notes <file> <slide_index> [notes_text]`
Read or set speaker notes. Omit `notes_text` to read.

```bash
# Read notes
cli-anything-onlyoffice pptx-speaker-notes /tmp/lecture.pptx 0 --json

# Set notes
cli-anything-onlyoffice pptx-speaker-notes /tmp/lecture.pptx 0 "Remember to introduce yourself" --json
```

#### `pptx-update-text <file> <slide_index> [--title <t>] [--body <b>]`
Update title and/or body text of an existing slide.

```bash
cli-anything-onlyoffice pptx-update-text /tmp/lecture.pptx 1 \
  --title "Updated Agenda" --body "New content here" --json
```

---

## Mode 5 — RDF Knowledge Graphs

**10 commands** — full CRUD, SPARQL 1.1, multi-format I/O, optional SHACL validation.

Requires: `rdflib>=7.0.0` (included in core). Optional: `pyshacl` for `rdf-validate`.

**Supported formats:** `turtle` (.ttl), `xml` (.rdf), `n3`, `nt`, `json-ld`, `trig`

**Built-in prefixes (auto-bound on create):** `rdf`, `rdfs`, `owl`, `xsd`, `foaf`, `dcterms`, `skos`

---

#### `rdf-create <file> [options]`
Create an empty RDF graph with optional base URI and custom prefixes.

Options: `--base <uri>`, `--format turtle|xml|n3|json-ld`, `--prefix <p>=<uri>`

```bash
cli-anything-onlyoffice rdf-create /tmp/knowledge.ttl \
  --base "http://example.org/" \
  --prefix ex="http://example.org/" \
  --format turtle --json
```
```json
{"success": true, "file": "/tmp/knowledge.ttl", "format": "turtle", "triples": 0, "prefixes": ["rdf", "rdfs", "owl", "xsd", "foaf", "dcterms", "skos", "ex"]}
```

#### `rdf-read <file> [--limit <n>]`
Parse an RDF file and return triples. Default limit: 100.

```bash
cli-anything-onlyoffice rdf-read /tmp/knowledge.ttl --limit 50 --json
```
```json
{
  "success": true,
  "triples": [
    {"subject": "http://example.org/Alice", "predicate": "http://xmlns.com/foaf/0.1/name", "object": "Alice"}
  ],
  "triple_count": 1,
  "namespaces": {"foaf": "http://xmlns.com/foaf/0.1/"}
}
```

#### `rdf-add <file> <subject> <predicate> <object> [options]`
Add a single triple. Object types: `uri` (default), `literal`, `bnode`

Options: `--type uri|literal|bnode`, `--lang <language_tag>`, `--datatype <xsd_uri>`, `--format <f>`

```bash
# URI object
cli-anything-onlyoffice rdf-add /tmp/knowledge.ttl \
  "http://example.org/Alice" \
  "http://www.w3.org/1999/02/22-rdf-syntax-ns#type" \
  "http://xmlns.com/foaf/0.1/Person" --json

# Literal object
cli-anything-onlyoffice rdf-add /tmp/knowledge.ttl \
  "http://example.org/Alice" \
  "http://xmlns.com/foaf/0.1/name" \
  "Alice Smith" --type literal --lang en --json

# Typed literal (date)
cli-anything-onlyoffice rdf-add /tmp/knowledge.ttl \
  "http://example.org/Alice" \
  "http://schema.org/birthDate" \
  "1990-01-01" --type literal \
  --datatype "http://www.w3.org/2001/XMLSchema#date" --json
```

#### `rdf-remove <file> [options]`
Remove triples matching a pattern. `None` / omitting acts as wildcard.

Options: `--subject <uri>`, `--predicate <uri>`, `--object <value>`, `--format <f>`

```bash
# Remove all triples about Alice
cli-anything-onlyoffice rdf-remove /tmp/knowledge.ttl \
  --subject "http://example.org/Alice" --json

# Remove specific triple
cli-anything-onlyoffice rdf-remove /tmp/knowledge.ttl \
  --subject "http://example.org/Alice" \
  --predicate "http://xmlns.com/foaf/0.1/name" \
  --object "Alice Smith" --json
```

#### `rdf-query <file> <sparql_query> [--limit <n>]`
Execute a SPARQL 1.1 query. Default limit: 100.

```bash
# SELECT query
cli-anything-onlyoffice rdf-query /tmp/knowledge.ttl \
  "SELECT ?s ?name WHERE { ?s <http://xmlns.com/foaf/0.1/name> ?name } LIMIT 10" --json

# ASK query
cli-anything-onlyoffice rdf-query /tmp/knowledge.ttl \
  "ASK { <http://example.org/Alice> a <http://xmlns.com/foaf/0.1/Person> }" --json
```
```json
{
  "success": true,
  "query_type": "SELECT",
  "results": [{"s": "http://example.org/Alice", "name": "Alice Smith"}],
  "count": 1
}
```

#### `rdf-export <file> <output_file> [--format <format>]`
Convert and export RDF to a different serialisation format.

```bash
# Turtle → JSON-LD
cli-anything-onlyoffice rdf-export /tmp/knowledge.ttl /tmp/knowledge.jsonld \
  --format json-ld --json

# Turtle → N-Triples
cli-anything-onlyoffice rdf-export /tmp/knowledge.ttl /tmp/knowledge.nt \
  --format nt --json
```

#### `rdf-merge <file_a> <file_b> [--output <file>] [--format <f>]`
Merge two RDF graphs into one. If `--output` is omitted, merges into `file_a`.

```bash
cli-anything-onlyoffice rdf-merge /tmp/graph1.ttl /tmp/graph2.ttl \
  --output /tmp/merged.ttl --format turtle --json
```
```json
{"success": true, "triples_a": 10, "triples_b": 15, "triples_merged": 25}
```

#### `rdf-stats <file>`
Graph statistics: triple count, unique subjects/predicates/objects, top predicates, RDF types.

```bash
cli-anything-onlyoffice rdf-stats /tmp/knowledge.ttl --json
```
```json
{
  "success": true,
  "triples": 42,
  "unique_subjects": 8,
  "unique_predicates": 12,
  "rdf_types": {"foaf:Person": 5, "foaf:Organization": 3},
  "top_predicates": [["foaf:name", 8], ["dcterms:title", 5]]
}
```

#### `rdf-namespace <file> [<prefix> <uri>] [--format <f>]`
List all namespace prefixes, or bind a new prefix.

```bash
# List all prefixes
cli-anything-onlyoffice rdf-namespace /tmp/knowledge.ttl --json

# Add a prefix
cli-anything-onlyoffice rdf-namespace /tmp/knowledge.ttl schema "http://schema.org/" --json
```

#### `rdf-validate <data_file> <shapes_file>`
Validate an RDF graph against a SHACL shapes graph. Requires `pyshacl`.

```bash
cli-anything-onlyoffice rdf-validate /tmp/data.ttl /tmp/shapes.ttl --json
```
```json
{
  "success": true,
  "conforms": false,
  "violations": [
    {"severity": "Violation", "focus": "http://example.org/Alice", "message": "Missing required foaf:mbox"}
  ]
}
```

---

## General Commands

#### `list`
List recent .docx/.xlsx/.pptx files from `~/Documents` and `~/Downloads`.

```bash
cli-anything-onlyoffice list --json
```

#### `open <file> [gui|web]`
Open a file in OnlyOffice Desktop Editors GUI or web viewer.

```bash
cli-anything-onlyoffice open /tmp/report.xlsx gui --json
```

#### `watch <file> [gui|web]`
Watch a file for changes and keep the GUI open for real-time viewing.

```bash
# Terminal 1: watch
cli-anything-onlyoffice watch /tmp/essay.docx gui

# Terminal 2: agent writes content, GUI reflects changes live
cli-anything-onlyoffice doc-append /tmp/essay.docx "New paragraph..." --json
```

#### `info <file>`
File metadata: type, size, sheet/slide/paragraph counts.

```bash
cli-anything-onlyoffice info /tmp/grades.xlsx --json
```

#### `status`
Check installation and all capability flags.

```bash
cli-anything-onlyoffice status --json
```
```json
{
  "success": true,
  "version": "4.0.0",
  "python": "/path/to/.venv/bin/python3",
  "python_docx": true,
  "openpyxl": true,
  "python_pptx": true,
  "rdflib": true,
  "rdflib_version": "7.6.0",
  "pyshacl": true,
  "capabilities": {
    "docx_create": true, "xlsx_charts": true,
    "rdf_create": true, "rdf_validate": true
  }
}
```

The `python` field shows which interpreter is running. If it doesn't point inside your `.venv`, you are using the wrong Python and imports will fail.

#### `help`
Machine-readable command reference (JSON mode recommended for agents).

```bash
cli-anything-onlyoffice help --json
```

---

### Backup Management

All writes auto-create a backup in `~/.cli-anything/backups/`.

#### `backup-list <file> [--limit <n>]`
List backups for a file.

```bash
cli-anything-onlyoffice backup-list /tmp/grades.xlsx --limit 10 --json
```

#### `backup-prune [options]`
Prune old backups by count or age.

```bash
cli-anything-onlyoffice backup-prune --file /tmp/grades.xlsx --keep 10 --json
cli-anything-onlyoffice backup-prune --days 30 --json  # prune all backups older than 30 days
```

#### `backup-restore <file> [options]`
Restore from backup.

```bash
cli-anything-onlyoffice backup-restore /tmp/grades.xlsx --latest --json
cli-anything-onlyoffice backup-restore /tmp/grades.xlsx --latest --dry-run --json  # preview only
```

---

## Command Reference Summary

| Category | Count | Commands |
|----------|-------|----------|
| Documents (.docx) | 25 | doc-create, doc-read, doc-append, doc-replace, doc-search, doc-insert, doc-delete, doc-format, doc-set-style, doc-list-styles, doc-highlight, doc-comment, doc-layout, doc-formatting-info, doc-add-table, doc-read-tables, doc-add-image, doc-add-hyperlink, doc-add-page-break, doc-add-list, doc-add-reference, doc-build-references, doc-set-metadata, doc-get-metadata, doc-word-count |
| Spreadsheets (.xlsx) | 32 | xlsx-create, xlsx-write, xlsx-read, xlsx-append, xlsx-search, xlsx-cell-read, xlsx-cell-write, xlsx-range-read, xlsx-delete-rows, xlsx-delete-cols, xlsx-sort, xlsx-filter, xlsx-calc, xlsx-formula, xlsx-formula-audit, xlsx-freq, xlsx-corr, xlsx-ttest, xlsx-mannwhitney, xlsx-chi2, xlsx-research-pack, xlsx-text-extract, xlsx-text-keywords, xlsx-sheet-list, xlsx-sheet-add, xlsx-sheet-delete, xlsx-sheet-rename, xlsx-merge-cells, xlsx-unmerge-cells, xlsx-format-cells, xlsx-csv-import, xlsx-csv-export |
| Charts (.xlsx) | 4 | chart-create, chart-comparison, chart-grade-dist, chart-progress |
| Presentations (.pptx) | 10 | pptx-create, pptx-add-slide, pptx-add-bullets, pptx-add-table, pptx-add-image, pptx-read, pptx-slide-count, pptx-delete-slide, pptx-speaker-notes, pptx-update-text |
| RDF Knowledge Graphs | 10 | rdf-create, rdf-read, rdf-add, rdf-remove, rdf-query, rdf-export, rdf-merge, rdf-stats, rdf-namespace, rdf-validate |
| General | 9 | list, open, watch, info, backup-list, backup-prune, backup-restore, status, help |
| **Total** | **90** | |

---

## Agent Workflow Examples

### Complete Grade Tracker

```bash
# 1. Create spreadsheet
cli-anything-onlyoffice xlsx-write /tmp/grades.xlsx \
  "Student,A1,A2,A3,Total" \
  "Alice,85,90,88,=B2+C2+D2;Bob,78,82,85,=B3+C3+D3;Charlie,92,88,95,=B4+C4+D4" \
  --sheet Grades --json

# 2. Style the header
cli-anything-onlyoffice xlsx-format-cells /tmp/grades.xlsx A1:E1 \
  --bold --color FFFFFF --bg-color 4472C4 --json

# 3. Calculate averages
cli-anything-onlyoffice xlsx-calc /tmp/grades.xlsx B all --sheet Grades --json

# 4. Visualize
cli-anything-onlyoffice chart-progress /tmp/grades.xlsx A E "Student Totals" --labels --json
cli-anything-onlyoffice chart-grade-dist /tmp/grades.xlsx E "Total Distribution" --json
```

### Research Report with APA Stats

```bash
# Frequency analysis
cli-anything-onlyoffice xlsx-freq /tmp/survey.xlsx C --sheet Sheet0 \
  --valid "SA,A,N,D,SD" --json

# Correlation
cli-anything-onlyoffice xlsx-corr /tmp/survey.xlsx B C --sheet Sheet0 --json

# T-test by gender
cli-anything-onlyoffice xlsx-ttest /tmp/survey.xlsx B A Male Female --json

# Full pack
cli-anything-onlyoffice xlsx-research-pack /tmp/survey.xlsx --sheet Sheet0 --json
```

### Lecture Presentation

```bash
cli-anything-onlyoffice pptx-create /tmp/lecture.pptx "Biology 101" "Spring 2026" --json
cli-anything-onlyoffice pptx-add-bullets /tmp/lecture.pptx "Objectives" \
  "Cell theory\nDNA structure\nMitosis vs Meiosis" --json
cli-anything-onlyoffice pptx-add-table /tmp/lecture.pptx "Cell Comparison" \
  "Type,Nucleus,Size" "Prokaryotic,No,1-5µm;Eukaryotic,Yes,10-100µm" --json
cli-anything-onlyoffice pptx-speaker-notes /tmp/lecture.pptx 0 "Introduce yourself first" --json
cli-anything-onlyoffice pptx-slide-count /tmp/lecture.pptx --json
```

### RDF Knowledge Graph Pipeline

```bash
# Create graph
cli-anything-onlyoffice rdf-create /tmp/knowledge.ttl \
  --base "http://example.org/" --prefix ex="http://example.org/" --json

# Add entities
cli-anything-onlyoffice rdf-add /tmp/knowledge.ttl \
  "http://example.org/Alice" \
  "http://www.w3.org/1999/02/22-rdf-syntax-ns#type" \
  "http://xmlns.com/foaf/0.1/Person" --json

cli-anything-onlyoffice rdf-add /tmp/knowledge.ttl \
  "http://example.org/Alice" "http://xmlns.com/foaf/0.1/name" "Alice" \
  --type literal --lang en --json

# Query
cli-anything-onlyoffice rdf-query /tmp/knowledge.ttl \
  "SELECT ?s ?name WHERE { ?s a <http://xmlns.com/foaf/0.1/Person> ; <http://xmlns.com/foaf/0.1/name> ?name }" --json

# Export to JSON-LD
cli-anything-onlyoffice rdf-export /tmp/knowledge.ttl /tmp/knowledge.jsonld --format json-ld --json

# Stats
cli-anything-onlyoffice rdf-stats /tmp/knowledge.ttl --json
```

### Essay with References

```bash
cli-anything-onlyoffice doc-create /tmp/essay.docx "Research Essay" "" --json
cli-anything-onlyoffice doc-insert /tmp/essay.docx "Introduction" 0 --style "Heading 1" --json
cli-anything-onlyoffice doc-append /tmp/essay.docx "Health outcomes improve when..." --json
cli-anything-onlyoffice doc-set-metadata /tmp/essay.docx --author "SLOANE Agent" --title "Health Research 2026" --json

# Add reference to sidecar
cli-anything-onlyoffice doc-add-reference /tmp/essay.docx \
  '{"author":"Smith, J.", "year":"2024", "title":"Health Outcomes Study", "source":"Journal of Health", "type":"journal", "doi":"10.1234/jh.2024"}' --json

# Build references section
cli-anything-onlyoffice doc-build-references /tmp/essay.docx --json
cli-anything-onlyoffice doc-word-count /tmp/essay.docx --json
```

---

## SLOANE OS Integration

This CLI is called by SLOANE subject agents via:

```python
result = cli_anything_run(tool="onlyoffice", args=[
    "xlsx-write", "/tmp/report.xlsx",
    "Month,Revenue", "Jan,5000;Feb,6200",
    "--sheet", "Data", "--json"
])
```

### Agent Tips

1. **Always use `--json`** — machine-readable, structured, parseable.
2. **Check `success` first** — every response has `{"success": true/false}`.
3. **Run `status --json`** on first run to confirm all capabilities are available. Check the `python` field to verify the venv interpreter is active.
4. **Run `help --json`** to get the full command reference programmatically.
5. **Backups are automatic** — every write is snapshotted. Use `backup-restore --latest` on error.
6. **Atomic writes** — no partial file corruption, safe to run concurrently from multiple threads or processes.
7. **RDF for knowledge** — use the RDF mode to build structured knowledge graphs that can be queried with SPARQL, exported to any format, and validated against SHACL shapes.
8. **Always invoke via the venv binary** — never `cd .venv && python3`. Use the full path: `.venv/bin/cli-anything-onlyoffice` or `.venv/bin/python3 -m cli_anything.onlyoffice.core.cli`. Running system `python3` will fail with `ModuleNotFoundError`.

---

## File Tree

```
agent-harness/
├── setup.py                          # Package config (v4.0.2)
├── README.md                         # This file
├── cli_anything/
│   └── onlyoffice/
│       ├── core/
│       │   ├── __init__.py
│       │   └── cli.py                # CLI router + dispatcher (~2,815 lines)
│       ├── utils/
│       │   ├── __init__.py
│       │   └── docserver.py          # Backend engine (~5,034 lines)
│       ├── skills/
│       │   ├── __init__.py
│       │   └── SKILL.md              # SLOANE skill manifest
│       └── tests/
│           ├── __init__.py
│           ├── test_concurrency_stress.py
│           ├── test_formula_safety.py
│           ├── test_inferential_stats.py
│           ├── test_production_readiness.py
│           └── test_research_pack.py
```

---

## Version History

| Version | Changes |
|---------|---------|
| **4.0.2** | Comprehensive bug-fix audit across all four modes: RDF full rewrite — 13 bugs fixed (ASK/CONSTRUCT/DESCRIBE query support, `rdf-remove` literal/bnode type flag, file-not-found guard, double-iteration fix, locking + atomic saves on all write methods, lang+datatype mutual exclusion, self-merge guard, `rdf-validate` structured violations output); xlsx — `xlsx-filter` now validates operator before executing, `xlsx-read` returns error on unknown sheet name instead of silently reading all sheets; docx — `doc-layout` landscape correctly swaps page dimensions, `doc-search` NameError fixed on table-only documents; pptx — `pptx-add-bullets` leading-empty-line enumerate-index bug fixed (orphan empty first paragraph) |
| **4.0.1** | Bug fixes: two-layer file locking (threading.Lock + fcntl.flock) fixes concurrent write loss under thread load; docx defaults corrected to A4/1" margins/Calibri 11pt/double spacing; xlsx auto-fits column widths and sets A4 paper size; pptx defaults to 16:9 (13.333"×7.5"); status exposes active Python interpreter path |
| **4.0.0** | Added RDF mode (10 commands), 42 new CRUD/sheet/cell commands across all modes, atomic saves, file locking, auto-backups, full JSON output, SHACL validation support |
| 3.0.0 | Chart creation (bar, line, pie, scatter), statistical tests (t-test, chi-square, correlation), research analysis pack |
| 2.0.0 | Presentation support, formula safety auditing |
| 1.0.0 | Documents and spreadsheets |

---

**Author:** SLOANE OS  
**License:** MIT  
**Python:** ≥ 3.8
