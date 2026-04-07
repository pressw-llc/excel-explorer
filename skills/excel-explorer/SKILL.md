---
name: excel-explorer
description: "Use this skill when exploring, analyzing, or understanding Excel workbooks (.xlsx). This includes tasks where the user wants to: understand a workbook's structure (sheets, named ranges, data regions); read cell values or formulas from specific ranges, rows, or columns; trace how a formula is calculated by walking its dependency tree to leaf inputs; find what cells would be affected by changing an input; understand a financial model's logic (formula patterns, assumptions, cross-sheet references); detect anomalies like hardcoded overrides in formula rows; validate that a balance sheet balances; compare time-series data period-over-period; search across sheets for values or formula functions; or explore formatting patterns. Trigger when the user references an .xlsx file and wants to understand or analyze it — not modify it. Do NOT trigger when the user wants to create, write, or edit Excel files, generate charts/visualizations, or convert between formats."
---

# Excel Explorer

Explore Excel workbooks using the `xlx` CLI tool — trace formula dependencies, understand model structure, and read data with pagination. Designed for AI agents exploring financial models.

## Prerequisites

Excel Explorer must be installed. Verify with:

```bash
xlx --help
```

If not installed, install from the repository:

```bash
# Option A: install globally
uv tool install git+https://github.com/pressw-llc/excel-explorer.git

# Option B: run directly without installing
uvx --from git+https://github.com/pressw-llc/excel-explorer.git xlx overview "file.xlsx"
```

## Exploration Workflow

When given an unfamiliar workbook, follow this sequence. Each step informs the next.

### Step 1: Orient — What Am I Looking At?

```bash
xlx overview "workbook.xlsx"
```

This returns every sheet with row/column counts, formula counts, and all named ranges. Use this to understand the workbook's architecture before diving in.

### Step 2: Understand Sheet Relationships

```bash
xlx sheet-flow "workbook.xlsx"
```

Shows which sheets feed data to which other sheets. Identifies source sheets (raw data), sink sheets (final outputs), and isolated sheets. This reveals the model's data flow.

### Step 3: Examine Specific Sheets

```bash
xlx describe "workbook.xlsx" "Sheet Name" --limit 20
```

Shows headers, dimensions, merged cells, and a data preview. For sheets with many formulas, also run:

```bash
xlx formula-map "workbook.xlsx" "Sheet Name" --limit 10
```

This groups formulas by pattern — revealing the sheet's logic without reading every cell.

### Step 4: Read Data

```bash
# Read a range
xlx read "workbook.xlsx" "Sheet" "A1:D20"

# Read with formulas visible
xlx read "workbook.xlsx" "Sheet" "A1:D20" --formulas

# Read a single row with column headers
xlx read-row "workbook.xlsx" "Income Statement" 5

# Read a column (e.g., line item labels)
xlx read-col "workbook.xlsx" "Sheet" "A" --limit 30
```

### Step 5: Trace Dependencies (The Key Feature)

To understand how a number is calculated:

```bash
xlx trace "workbook.xlsx" "Summary!B15" --depth 5
```

This recursively walks the formula tree, showing every input that feeds the cell. Leaf nodes marked `[INPUT]` are the model's hardcoded assumptions. Shared dependencies are flagged.

To see what would break if an input changed:

```bash
xlx dependents "workbook.xlsx" "Assumptions!B2"
```

### Step 6: Find Model Inputs

```bash
xlx find-inputs "workbook.xlsx"
xlx find-inputs "workbook.xlsx" --sheet Assumptions
```

Returns all hardcoded cells that other cells depend on, sorted by impact (most dependents first). These are the model's drivers.

## Pagination

All commands support pagination to avoid overwhelming context windows:

| Flag | Default | Purpose |
|------|---------|---------|
| `--limit N` | 50 | Max rows returned |
| `--offset N` | 0 | Skip first N rows |
| `--max-cols N` | 20 | Max columns returned |
| `--col-offset N` | 0 | Skip first N columns |
| `--depth N` | 5 | Max recursion for trace/dependents |

Every command returns a metadata header indicating whether results were truncated:

```
file: workbook.xlsx
sheet: Summary
showing: rows 1-50 of 200
truncated: true
---
```

When `truncated: true`, increase `--offset` to page through remaining data.

## Financial Analysis Commands

For financial workbooks, these specialized commands are available:

```bash
# Summarize all assumptions and named ranges
xlx summarize-assumptions "workbook.xlsx"

# Compare a row period-over-period (auto-detects date headers)
xlx compare-periods "workbook.xlsx" "Income Statement" 7

# Find formula anomalies (hardcoded overrides, pattern breaks)
xlx find-anomalies "workbook.xlsx" "Income Statement"

# Validate balance sheet equation (A = L + E)
xlx validate-balance "workbook.xlsx" "Balance Sheet"
```

## Search

```bash
# Search cell values across all sheets
xlx search "workbook.xlsx" "Revenue"

# Also search inside formula strings
xlx search "workbook.xlsx" "SUMPRODUCT" --formulas

# Find cells with semantic formatting (bold, colors, borders)
xlx find-formatting "workbook.xlsx" "Sheet Name"
```

## Command Reference

For detailed help on any command: `xlx <command> --help`

For the complete command reference, consult **`references/commands.md`**.

## Tips for Effective Exploration

- **Start broad, then narrow.** Always run `overview` and `sheet-flow` before reading specific cells.
- **Use formula-map before trace.** Understanding the dominant patterns on a sheet makes individual traces more meaningful.
- **Named ranges are important.** Financial models use them for key data columns and assumptions. Check `named-ranges` early.
- **Pagination defaults are conservative.** For large workbooks, start with `--limit 10` to sample, then increase.
- **Formatting carries meaning.** In financial models, blue fill = input assumption, bold = subtotal. Use `find-formatting` to surface these patterns.
- **Trace is the killer feature.** When asked "how is this number calculated?", `trace` gives the complete answer.
