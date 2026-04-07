# Excel Explorer Command Reference

Complete reference for all 17 `xlx` commands.

## Orientation Commands

### overview

Show workbook overview: sheets, dimensions, formula counts, named ranges.

```bash
xlx overview FILE
```

The first command to run on any workbook. Returns every sheet with its row/column count and number of formulas, plus all named ranges.

### describe

Describe a sheet: headers, dimensions, merged cells, and data preview.

```bash
xlx describe FILE SHEET [--limit N] [--offset N] [--max-cols N] [--col-offset N]
```

Shows row/column counts, merged cell ranges, the header row, and a paginated preview. Formulas shown as `[formula]` prefixed values.

### named-ranges

List all named ranges with their targets.

```bash
xlx named-ranges FILE
```

Named ranges are workbook-level aliases pointing to cell ranges (e.g., "Sales" -> 'Data'!$G$3:$G$500000). In financial models they typically represent key data columns, assumptions, or lookup tables.

## Reading Commands

### read

Read cell values from a range on a sheet.

```bash
xlx read FILE SHEET RANGE [--formulas] [--limit N] [--offset N] [--max-cols N] [--col-offset N]
```

Returns cell values. With `--formulas`, shows both the computed value and the underlying formula side by side.

### read-row

Read a full row with auto-detected column headers.

```bash
xlx read-row FILE SHEET ROW [--formulas] [--max-cols N] [--col-offset N]
```

Labels each cell with its column header from row 1. Natural for reading a P&L line item across time periods.

### read-col

Read a full column.

```bash
xlx read-col FILE SHEET COL [--formulas] [--limit N] [--offset N]
```

All values down a single column. Good for reading line item labels (column A) or a single period's data.

### tables

Auto-detect contiguous table regions on a sheet.

```bash
xlx tables FILE SHEET [--limit N] [--offset N]
```

Scans for blocks of non-empty rows separated by blank rows. Returns each table's range, dimensions, and detected header row. Useful for sheets with multiple logical sections.

## Dependency Commands

### trace

Trace the full dependency tree from a cell down to its leaf inputs.

```bash
xlx trace FILE CELL_REF [--sheet TEXT] [--depth N]
```

CELL_REF format: `"Sheet!A1"` or just `A1` with `--sheet`. Recursively walks formula references across sheets, building an indented tree. Leaf nodes (hardcoded values) marked `[INPUT]`. Shared dependencies flagged. This is the key command for understanding how a number is calculated.

Example output:
```
SUMMARY!B15 = =Revenue-COGS-OpEx-Tax
├── REVENUE BUILD!B20 = =Units*Price
│   ├── ASSUMPTIONS!B5 = 10000  [INPUT]
│   └── ASSUMPTIONS!B6 = 45.00  [INPUT]
└── TAX!B5 = =(Revenue-COGS-OpEx)*TaxRate
    └── ASSUMPTIONS!B15 = 0.21  [INPUT]

leaf_inputs: 3 unique
```

### dependents

Find all cells that depend on a given cell (reverse trace).

```bash
xlx dependents FILE CELL_REF [--sheet TEXT] [--depth N]
```

The opposite of trace — shows what would be affected by changing this cell. Walks successors across sheets up to `--depth` levels.

### formula-map

Show unique formula patterns on a sheet, grouped by structure.

```bash
xlx formula-map FILE SHEET [--limit N] [--offset N]
```

Normalizes formulas by replacing specific cell references with placeholders, then groups identical patterns. Reveals the model's logic without reading every cell.

### find-inputs

Find hardcoded cells that other cells depend on (model inputs/assumptions).

```bash
xlx find-inputs FILE [--sheet TEXT] [--limit N] [--offset N]
```

Scans the dependency graph for cells with no formula that have successors. Results sorted by impact (most dependents first). These are the model's drivers.

### sheet-flow

Map which sheets reference which other sheets.

```bash
xlx sheet-flow FILE
```

Parses cross-sheet formula references to build a directed graph. Identifies source sheets (data origins), sink sheets (final outputs), and isolated sheets (no cross-sheet refs).

## Financial Analysis Commands

### summarize-assumptions

Summarize model assumptions: named ranges and input sheets.

```bash
xlx summarize-assumptions FILE [--limit N] [--offset N]
```

Collects all named ranges and scans for sheets named "Assumptions", "Inputs", "Parameters", or "Drivers", extracting parameter/value pairs.

### compare-periods

Compare a time-series row period-over-period.

```bash
xlx compare-periods FILE SHEET ROW [--max-cols N] [--col-offset N]
```

Reads a row, auto-detects date headers, computes absolute change and growth rate between consecutive periods. Shows summary stats (min, max, average, total growth).

### find-anomalies

Detect formula anomalies — pattern breaks and hardcoded overrides.

```bash
xlx find-anomalies FILE SHEET [--limit N] [--offset N]
```

Scans each row for its dominant formula pattern, flags cells that deviate: different formula structure, or a hardcoded value in the middle of a formula row.

### validate-balance

Validate that a balance sheet balances (Assets = Liabilities + Equity).

```bash
xlx validate-balance FILE SHEET
```

Finds rows by label matching (case-insensitive), checks the equation across every period column. Reports PASS/FAIL per period and overall BALANCED/IMBALANCED.

## Search Commands

### search

Search for a value or text across all sheets.

```bash
xlx search FILE QUERY [--formulas] [--limit N] [--offset N]
```

Case-insensitive search across every cell. With `--formulas`, also matches inside formula strings (e.g., find all cells using VLOOKUP).

### find-formatting

Report cells with notable formatting (bold, fill colors, borders).

```bash
xlx find-formatting FILE SHEET [--limit N] [--offset N]
```

In financial models, formatting carries meaning — blue fill = input assumption, bold = subtotal, borders = section boundaries.

## Global Pagination Flags

| Flag | Default | Applies to |
|------|---------|-----------|
| `--limit N` | 50 | All read/search commands |
| `--offset N` | 0 | All read/search commands |
| `--max-cols N` | 20 | read, read-row, describe, compare-periods |
| `--col-offset N` | 0 | read, read-row, describe, compare-periods |
| `--depth N` | 5 | trace, dependents |
| `--formulas` | off | read, read-row, read-col, search |
| `--sheet TEXT` | - | trace, dependents, find-inputs |
