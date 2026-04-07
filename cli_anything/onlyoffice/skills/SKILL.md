---
name: onlyoffice
version: 3.0.0
author: SLOANE OS
description: Full programmatic control over Documents (.docx), Spreadsheets (.xlsx), and Presentations (.pptx)
tags: [productivity, documents, office, onlyoffice, charts, spreadsheets]
---

# CLI-Anything OnlyOffice v3.0

A comprehensive command-line interface for creating, reading, and editing Office documents programmatically. Supports full document manipulation, spreadsheet operations with charts, and presentation building.

## Installation

```bash
pip install -e /home/benbi/cli-anything/onlyoffice/agent-harness
```

Verify:
```bash
which cli-anything-onlyoffice
cli-anything-onlyoffice status --json
```

## Quick Reference

| Document Type | Library | Commands |
|--------------|---------|----------|
| Documents (.docx) | python-docx | 8 commands |
| Spreadsheets (.xlsx) | openpyxl | 11 commands (+ charts) |
| Presentations (.pptx) | python-pptx | 6 commands |

---

## DOCUMENTS (.docx)

### open <file> [mode]
Open a file in OnlyOffice GUI or web viewer.

```bash
# Open in desktop GUI
cli-anything-onlyoffice open /tmp/grades.xlsx gui --json

# Open in web viewer (requires Document Server setup)
cli-anything-onlyoffice open /tmp/grades.xlsx web --json
```

**JSON Output:**
```json
{
  "success": true,
  "file": "/tmp/grades.xlsx",
  "mode": "gui",
  "message": "Opened in OnlyOffice Desktop Editors"
}
```

### watch <file> [mode]
Watch a file for changes and auto-open in GUI for real-time viewing.

```bash
# Watch file and auto-open GUI (Ctrl+C to stop)
cli-anything-onlyoffice watch /tmp/essay.docx gui --json
```

**Use Case:** Perfect for real-time feedback as the agent works on documents. The GUI will show changes as they happen!

### doc-create <file> <title> <content> [--open]
Create a new .docx document. Add `--open` to auto-launch GUI.

```bash
cli-anything-onlyoffice doc-create essay.docx "My Essay" "Introduction..." --open --json
```

### xlsx-write <file> <headers> <data> [--open]
Write data to spreadsheet. Add `--open` to auto-launch GUI.

```bash
cli-anything-onlyoffice xlsx-write grades.xlsx "Student,Grade" "Alice,85" --open --json
```

**JSON Output:**
```json
{
  "success": true,
  "file": "essay.docx",
  "title": "My Essay",
  "size": 8192
}
```

### doc-read <file>
Read and extract all content from a .docx document.

```bash
cli-anything-onlyoffice doc-read essay.docx --json
```

**JSON Output:**
```json
{
  "success": true,
  "file": "essay.docx",
  "paragraphs": ["Introduction...", "Body paragraph..."],
  "paragraph_count": 5,
  "full_text": "Introduction...\n\nBody paragraph..."
}
```

### doc-append <file> <content>
Append text to a .docx document.

```bash
cli-anything-onlyoffice doc-append essay.docx "New paragraph about findings..." --json
```

### doc-replace <file> <search> <replace>
Find and replace text in a .docx document.

```bash
cli-anything-onlyoffice doc-replace essay.docx "draft" "final" --json
```

---

## SPREADSHEETS (.xlsx)

### xlsx-create <file> [sheet]
Create a new .xlsx spreadsheet.

```bash
cli-anything-onlyoffice xlsx-create grades.xlsx "Grades" --json
```

### xlsx-write <file> <headers> <data>
Write headers and data rows to a spreadsheet. Values starting with '=' are formulas.

```bash
cli-anything-onlyoffice xlsx-write grades.xlsx "Student,Assignment1,Assignment2,Total" \
  "Alice,85,90,=B2+C2;Bob,78,82,=B3+C3" --json
```

**JSON Output:**
```json
{
  "success": true,
  "file": "grades.xlsx",
  "rows_written": 2,
  "columns": 4
}
```

### xlsx-read <file> [sheet]
Read all data from a spreadsheet.

```bash
cli-anything-onlyoffice xlsx-read grades.xlsx --json
```

### xlsx-append <file> <row-data>
Append a row to a spreadsheet.

```bash
cli-anything-onlyoffice xlsx-append grades.xlsx "Charlie,92,88" --json
```

### xlsx-search <file> <text>
Search for text in a spreadsheet.

```bash
cli-anything-onlyoffice xlsx-search grades.xlsx "Alice" --json
```

### xlsx-calc <file> <column> <operation>
Calculate column statistics (sum, avg, min, max).

```bash
cli-anything-onlyoffice xlsx-calc grades.xlsx B sum --json
```

**JSON Output:**
```json
{
  "success": true,
  "column": "B",
  "count": 5,
  "sum": 425,
  "average": 85.0,
  "min": 70,
  "max": 95
}
```

### xlsx-formula <file> <cell> <formula>
Add formula to a specific cell.

```bash
cli-anything-onlyoffice xlsx-formula grades.xlsx D2 "=AVERAGE(B2:C2)" --json
```

---

## CHARTS (.xlsx) - NEW IN v3.0

### chart-create <file> <type> <data_range> <cat_range> <title> [options]
Create a chart in a spreadsheet.

**Chart Types:** bar, column, bar_horizontal, line, pie, scatter

```bash
# Basic bar chart
cli-anything-onlyoffice chart-create grades.xlsx bar B2:D6 A2:A6 "Assignment Comparison" --json

# With options
cli-anything-onlyoffice chart-create grades.xlsx bar B2:D6 A2:A6 "Sales Data" \
  --sheet Sheet1 --output-sheet Charts --x-label "Students" --y-label "Scores" \
  --labels --legend-pos bottom --colors FF0000,00FF00,0000FF --json
```

**Options:**
- `--sheet <name>`: Source sheet name
- `--output-sheet <name>`: Sheet to place chart (creates if not exists)
- `--x-label <text>`: X-axis label
- `--y-label <text>`: Y-axis label
- `--labels`: Show data labels on chart
- `--no-legend`: Hide legend
- `--legend-pos <pos>`: Legend position (right, top, bottom, left)
- `--colors <hex,hex>`: Custom colors for series

### chart-comparison <file> <type> <title> [options]
Create a comparison chart from structured data.

```bash
cli-anything-onlyoffice chart-comparison grades.xlsx line "Assignment Trends" \
  --start-row 2 --cat-col 1 --value-cols 2,3,4 --output A10 --labels --json
```

**Options:**
- `--sheet <name>`: Source sheet
- `--start-row <n>`: Starting row (default: 1)
- `--start-col <n>`: Starting column (default: 1)
- `--cats <n>`: Number of category rows
- `--series <n>`: Number of data series
- `--cat-col <n>`: Category column (1-indexed)
- `--value-cols <n,n,n>`: Data columns (1-indexed)
- `--output <cell>`: Output cell position
- `--labels`: Show data labels
- `--no-legend`: Hide legend

### chart-grade-dist <file> <grade_col> <title> [options]
Create a pie chart showing grade distribution. Automatically counts grade frequencies.

```bash
cli-anything-onlyoffice chart-grade-dist grades.xlsx B "Grade Distribution" --json
```

**Options:**
- `--sheet <name>`: Source sheet
- `--output <cell>`: Output cell (default: F2)

**JSON Output:**
```json
{
  "success": true,
  "file": "grades.xlsx",
  "chart_type": "pie",
  "title": "Grade Distribution",
  "total_grades": 5,
  "distribution": {"A": 2, "B": 2, "C": 1},
  "output_cell": "F2"
}
```

### chart-progress <file> <student_col> <grade_col> <title> [options]
Create a horizontal bar chart showing individual student grades.

```bash
cli-anything-onlyoffice chart-progress grades.xlsx A B "Student Grades" --json
```

**Options:**
- `--sheet <name>`: Source sheet
- `--output <cell>`: Output cell (default: D2)
- `--labels`: Show data labels (default: true)
- `--no-labels`: Hide data labels

---

## PRESENTATIONS (.pptx)

### pptx-create <file> <title> [subtitle]
Create a new .pptx presentation with title slide.

```bash
cli-anything-onlyoffice pptx-create lecture.pptx "Biology 101" "Introduction to Cell Structure" --json
```

### pptx-add-slide <file> <title> [content] [layout]
Add a slide. Layouts: title_only, content, blank, two_content.

```bash
cli-anything-onlyoffice pptx-add-slide lecture.pptx "Key Concepts" "Main points here" content --json
```

### pptx-add-bullets <file> <title> <bullets>
Add a bullet-point slide. Bullets separated by \n.

```bash
cli-anything-onlyoffice pptx-add-bullets lecture.pptx "Key Concepts" "Cell theory\nDNA structure\nMitochondria" --json
```

### pptx-add-table <file> <title> <headers> <data>
Add a table slide. Headers comma-separated, rows semicolon-separated.

```bash
cli-anything-onlyoffice pptx-add-table lecture.pptx "Cell Types" "Type,Size,Features" \
  "Prokaryotic,1-5um,No nucleus;Eukaryotic,10-100um,Has nucleus" --json
```

### pptx-add-image <file> <title> <image_path>
Add an image slide.

```bash
cli-anything-onlyoffice pptx-add-image lecture.pptx "Cell Diagram" ~/images/cell.png --json
```

### pptx-read <file>
Read all slides and content from a presentation.

```bash
cli-anything-onlyoffice pptx-read lecture.pptx --json
```

---

## GENERAL COMMANDS

### list
List recent documents, spreadsheets, and presentations.

```bash
cli-anything-onlyoffice list --json
```

### info <file>
Get file information.

```bash
cli-anything-onlyoffice info grades.xlsx --json
```

### status
Check installation and capabilities.

```bash
cli-anything-onlyoffice status --json
```

**JSON Output:**
```json
{
  "success": true,
  "document_server": {"healthy": true},
  "python_docx": true,
  "openpyxl": true,
  "python_pptx": true,
  "capabilities": {
    "docx_create": true,
    "xlsx_create": true,
    "xlsx_charts": true,
    "pptx_create": true
  }
}
```

### help
Show usage information.

```bash
cli-anything-onlyoffice help --json
```

---

## GUI Integration - Real-Time Document Viewing

The CLI integrates with OnlyOffice Desktop Editors GUI for real-time document viewing:

### Auto-Open After Creation
Add `--open` flag to automatically launch the GUI after creating/modifying a file:

```bash
# Document opens in GUI after creation
cli-anything-onlyoffice doc-create essay.docx "Title" "Content" --open

# Spreadsheet opens in GUI after write
cli-anything-onlyoffice xlsx-write data.xlsx "A,B,C" "1,2,3" --open
```

### Watch Mode for Real-Time Feedback
Use `watch` command to monitor a file and keep GUI open for live updates:

```bash
# Terminal 1: Start watching
cli-anything-onlyoffice watch /tmp/essay.docx gui

# Terminal 2: Agent writes content
cli-anything-onlyoffice doc-append /tmp/essay.docx "New paragraph..."
# GUI updates automatically!
```

### SLOANE UI Integration
The SLOANE web interface includes a **Document Viewer** panel (📄 button in header) that:
- Embeds OnlyOffice Document Server in an iframe
- Auto-refreshes when files change
- Shows documents created by agents in real-time

```python
# Agent can trigger document viewer
cli_anything_run(tool="onlyoffice", args=["open", "/tmp/essay.docx", "gui"])
```

---

## ACADEMIC WORKFLOW EXAMPLES

### Create Grade Tracker with Charts

```bash
# 1. Create grade spreadsheet
cli-anything-onlyoffice xlsx-write grades.xlsx \
  "Student,Assignment1,Assignment2,Assignment3,Total" \
  "Alice,85,90,88,=B2+C2+D2;Bob,78,82,85,=B3+C3+D3;Charlie,92,88,95,=B4+C4+D4" --json

# 2. Add student progress chart
cli-anything-onlyoffice chart-progress grades.xlsx A B "Student Performance" --output D2 --labels --json

# 3. Add assignment comparison chart
cli-anything-onlyoffice chart-create grades.xlsx bar B2:D4 A2:A4 "Assignment Comparison" \
  --output-sheet Charts --x-label "Student" --y-label "Score" --labels --json

# 4. Add grade distribution pie chart
cli-anything-onlyoffice chart-grade-dist grades.xlsx E "Total Grade Distribution" --output H2 --json
```

### Create Lecture Presentation

```bash
# 1. Create presentation
cli-anything-onlyoffice pptx-create lecture.pptx "Biology 101" "Cell Biology Introduction" --json

# 2. Add bullet points slide
cli-anything-onlyoffice pptx-add-bullets lecture.pptx "Learning Objectives" \
  "Understand cell theory\nIdentify organelles\nExplain cellular functions" --json

# 3. Add comparison table
cli-anything-onlyoffice pptx-add-table lecture.pptx "Cell Types" \
  "Type,Nucleus,Size,Examples" \
  "Prokaryotic,No,1-5um,Bacteria;Eukaryotic,Yes,10-100um,Animals,Plants" --json
```

### Create Essay Document

```bash
# 1. Create document
cli-anything-onlyoffice doc-create essay.docx "Research Essay" "Introduction paragraph..." --json

# 2. Add content
cli-anything-onlyoffice doc-append essay.docx "Body paragraph with analysis..." --json
cli-anything-onlyoffice doc-append essay.docx "Conclusion summarizing findings..." --json

# 3. Find and replace
cli-anything-onlyoffice doc-replace essay.docx "draft" "final" --json

# 4. Read and verify
cli-anything-onlyoffice doc-read essay.docx --json
```

---

## Agent Usage Patterns

### Pattern 1: Create Academic Report with Charts

```python
# Create spreadsheet with data
cli_anything_run(tool="onlyoffice", args=[
    "xlsx-write", "report.xlsx",
    "Month,Revenue,Expenses,Profit",
    "Jan,5000,3000,=B2-C2;Feb,6000,3500,=B3-C3;Mar,5500,3200,=B4-C4"
])

# Add visualizations
cli_anything_run(tool="onlyoffice", args=[
    "chart-create", "report.xlsx", "line", "B2:D4", "A2:A4", "Revenue Trend",
    "--labels", "--x-label", "Month", "--y-label", "Amount"
])

cli_anything_run(tool="onlyoffice", args=[
    "chart-grade-dist", "report.xlsx", "D", "Profit Distribution", "--output", "F2"
])
```

### Pattern 2: Student Grade Analysis

```python
# Load grades
grades = cli_anything_run(tool="onlyoffice", args=["xlsx-read", "grades.xlsx", "--json"])

# Create analysis charts
cli_anything_run(tool="onlyoffice", args=[
    "chart-progress", "grades.xlsx", "A", "E", "Student Grades", "--output", "G2"
])

cli_anything_run(tool="onlyoffice", args=[
    "chart-comparison", "grades.xlsx", "bar", "Assignment Comparison",
    "--start-row", "2", "--cat-col", "1", "--value-cols", "2,3,4", "--output", "I10"
])
```

---

## Supported Formats

| Category | Formats | Library |
|----------|---------|---------|
| Documents | .docx, .doc, .odt, .txt, .rtf | python-docx |
| Spreadsheets | .xlsx, .xls, .ods, .csv | openpyxl |
| Presentations | .pptx, .ppt, .odp | python-pptx |

## Chart Types Supported

| Type | Use Case | Example |
|------|----------|---------|
| bar/column | Comparing categories | Assignment scores |
| bar_horizontal | Long labels | Student names |
| line | Trends over time | Grade progression |
| pie | Distribution | Grade breakdown |
| scatter | Correlation | Study time vs grades |

## Integration with SLOANE

The CLI is automatically indexed for SLOANE subject agents after installation. Agents can:

- ✅ Create and edit documents programmatically
- ✅ Create spreadsheets with formulas
- ✅ **Generate charts (bar, line, pie, scatter)**
- ✅ Build presentations with slides, bullets, tables, images
- ✅ Search and extract content from files
- ✅ Perform calculations on spreadsheet data

## Requirements

- Python 3.10+
- python-docx (for .docx)
- openpyxl (for .xlsx with charts)
- python-pptx (for .pptx)

---

**Last Updated:** 2026-04-01
**Version:** 3.0.0
**New Features:** Chart creation (bar, line, pie, scatter), comparison charts, grade distribution, student progress charts