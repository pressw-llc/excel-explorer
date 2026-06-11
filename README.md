# Excel Explorer

CLI tool for exploring Excel workbooks. Trace formula dependencies, understand model structure, read data with pagination. Designed for use by Claude Code agents.

## Install

```bash
# As a standalone CLI tool (recommended)
uv tool install excel-explorer

# Or with pip
pip install excel-explorer

# Or run without installing
uvx --from excel-explorer xlx overview "file.xlsx"
```

This installs the `xlx` command.

### Development setup

```bash
git clone https://github.com/pressw-llc/excel-explorer.git
cd excel-explorer
uv sync --group dev
uv run pytest
```

## Usage

Every command takes a workbook path as its first argument. Paths with spaces should be quoted.

### Orientation

Get a high-level view of a workbook before diving into details.

```bash
# Workbook overview: sheets, dimensions, formula counts, named ranges
xlx overview "Financials.xlsx"

# Describe a single sheet: headers, merged cells, data regions
xlx describe "Financials.xlsx" "Balance Sheet"

# List all named ranges and their targets
xlx named-ranges "Financials.xlsx"
```

### Reading

Read cell values and structured data from sheets.

```bash
# Read a range (values only)
xlx read "Financials.xlsx" "Income Statement" A1:D20

# Read a range showing formulas alongside computed values
xlx read --formulas "Financials.xlsx" "Income Statement" A1:D20

# Read a full row with auto-detected headers
xlx read-row "Financials.xlsx" "Income Statement" 5

# Read a full column
xlx read-col "Financials.xlsx" "Income Statement" B

# Auto-detect contiguous table regions on a sheet
xlx tables "Financials.xlsx" "Income Statement"
```

### Dependencies

Understand how cells relate to each other and trace formula chains.

```bash
# Trace a cell's dependency tree down to leaf inputs
xlx trace "Financials.xlsx" "Income Statement!B10"
xlx trace "Financials.xlsx" B10 --sheet "Income Statement"

# Find all cells that depend on a given cell (reverse trace)
xlx dependents "Financials.xlsx" "Assumptions!C5"

# Show unique formula patterns on a sheet, grouped by structure
xlx formula-map "Financials.xlsx" "Income Statement"

# Find hardcoded cells that other cells depend on (model inputs/assumptions)
xlx find-inputs "Financials.xlsx"
xlx find-inputs "Financials.xlsx" --sheet "Assumptions"

# Map which sheets reference which other sheets
xlx sheet-flow "Financials.xlsx"
```

### Financial Analysis

Higher-level commands built for financial models.

```bash
# Find and summarize assumption sheets and named ranges
xlx summarize-assumptions "Financials.xlsx"

# Read a time-series row and compute period-over-period changes
xlx compare-periods "Financials.xlsx" "Income Statement" 10

# Detect formula pattern breaks (hardcoded values in formula rows, etc.)
xlx find-anomalies "Financials.xlsx" "Income Statement"

# Check that Assets = Liabilities + Equity across all period columns
xlx validate-balance "Financials.xlsx" "Balance Sheet"
```

### Search

Search across the entire workbook.

```bash
# Search all cells for a value or text pattern
xlx search "Financials.xlsx" "revenue"

# Also search inside formula strings
xlx search --formulas "Financials.xlsx" "VLOOKUP"

# Report cells with notable formatting: bold, fill colors, borders
xlx find-formatting "Financials.xlsx" "Income Statement"
```

## Pagination

Large workbooks can return a lot of data. Use these flags to page through results without overwhelming the context window.

| Flag | Default | Description |
|---|---|---|
| `--limit N` | 50 | Maximum number of rows (or items) to return |
| `--offset N` | 0 | Skip the first N rows before returning results |
| `--max-cols N` | 20 | Maximum number of columns to show |
| `--col-offset N` | 0 | Skip the first N columns before returning results |
| `--depth N` | 5 | Maximum recursion depth for `trace` and `dependents` |

Example — page through a large sheet in chunks of 100 rows (use a bounded
row range; column-only ranges like `A:Z` are returned column-major and don't
paginate by row):

```bash
xlx read "Financials.xlsx" "Transactions" A1:Z1000 --limit 100 --offset 0
xlx read "Financials.xlsx" "Transactions" A1:Z1000 --limit 100 --offset 100
xlx read "Financials.xlsx" "Transactions" A1:Z1000 --limit 100 --offset 200
```

## Output Format

Every command outputs a metadata header followed by `---` and then the body. The header uses a YAML-style `key: value` format describing what was returned so agents can make pagination decisions without re-reading. For example, `xlx read`:

```
file: financial_model.xlsx
sheet: P&L
range: A1:C5
showing: rows 1-5 of 5
truncated: false
---
  A1: Profit & Loss Statement
  A3: Line Item
  A4: Gross Revenue
  B4: ='Revenue Build'!B8
  C4: ='Revenue Build'!C8
  ...
```

The exact header keys vary by command (`range`, `row`, `column`, `query`, `matches`, ...). The pagination-relevant ones:

- `showing: rows X-Y of Z` — current window and total rows available
- `truncated: true` — results were cut off; advance `--offset` or increase `--limit`
- `shown` / `offset` / `limit` — pagination position on search-style commands

Commands that return summaries rather than rows (`overview`, `named-ranges`, `sheet-flow`, `validate-balance`, `trace`, `dependents`) don't paginate and omit these fields.
