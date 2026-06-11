"""Excel Explorer CLI — Click entry point for all 18 commands."""

import sys
import click

import excel_explorer.explorer as explorer
import excel_explorer.dependencies as dependencies
import excel_explorer.analysis as analysis
import excel_explorer.search as search_mod


@click.group()
def cli():
    """Excel Explorer — explore workbook structure, formulas, and dependencies."""
    pass


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _parse_cell_ref(cell_ref: str, sheet_option: str | None) -> tuple[str, str]:
    """Parse a cell reference that may be in 'Sheet!Cell' or just 'Cell' format.

    If cell_ref contains '!', split on it and use the left part as the sheet
    name (stripping surrounding quotes). Otherwise, use sheet_option.

    Returns (sheet, cell).
    """
    if "!" in cell_ref:
        sheet_part, cell_part = cell_ref.split("!", 1)
        sheet_name = sheet_part.strip("'\"")
        return sheet_name, cell_part
    if sheet_option:
        return sheet_option, cell_ref
    raise click.UsageError(
        "Provide cell as 'Sheet!Cell' or pass --sheet and just the cell address."
    )


# ---------------------------------------------------------------------------
# Tier 1 — Orientation
# ---------------------------------------------------------------------------

@cli.command()
@click.argument("file", type=click.Path(exists=True))
def overview(file):
    """Show workbook overview: sheets, dimensions, formula counts, named ranges.

    The first command to run on any workbook. Returns every sheet with its row/column
    count and number of formulas, plus all named ranges.

    \b
    Examples:
      xlx overview "Model.xlsx"
      xlx overview "2025 - Peer Group Analysis.xlsx"
    """
    try:
        click.echo(explorer.overview(file))
    except Exception as exc:
        click.echo(f"Error: {exc}", err=True)
        sys.exit(1)


@cli.command()
@click.argument("file", type=click.Path(exists=True))
@click.argument("sheet")
@click.option("--limit", default=50, show_default=True, help="Max rows to preview.")
@click.option("--offset", default=0, show_default=True, help="Row offset for pagination.")
@click.option("--max-cols", default=20, show_default=True, help="Max columns to show.")
@click.option("--col-offset", default=0, show_default=True, help="Column offset.")
def describe(file, sheet, limit, offset, max_cols, col_offset):
    """Describe a sheet: headers, dimensions, merged cells, and data preview.

    Shows row/column counts, merged cell ranges, the header row, and a paginated
    preview of the sheet's data. Formulas are shown as [formula] prefixed values.

    \b
    Examples:
      xlx describe "Model.xlsx" "Income Statement"
      xlx describe "Model.xlsx" "Summary" --limit 10 --max-cols 5
    """
    try:
        click.echo(explorer.describe(file, sheet, limit=limit, offset=offset,
                                     max_cols=max_cols, col_offset=col_offset))
    except Exception as exc:
        click.echo(f"Error: {exc}", err=True)
        sys.exit(1)


@cli.command("named-ranges")
@click.argument("file", type=click.Path(exists=True))
def named_ranges(file):
    """List all named ranges with their targets.

    Named ranges are workbook-level aliases that point to cell ranges
    (e.g., "Sales" -> 'Data'!$G$3:$G$500000). In financial models they
    typically represent key data columns, assumptions, or lookup tables.

    \b
    Example:
      xlx named-ranges "2025 - Peer Group Analysis.xlsx"
    """
    try:
        click.echo(explorer.named_ranges(file))
    except Exception as exc:
        click.echo(f"Error: {exc}", err=True)
        sys.exit(1)


# ---------------------------------------------------------------------------
# Tier 2 — Reading
# ---------------------------------------------------------------------------

@cli.command()
@click.argument("file", type=click.Path(exists=True))
@click.argument("sheet")
@click.argument("range_str", metavar="RANGE")
@click.option("--formulas", is_flag=True, default=False, help="Show formulas alongside computed values.")
@click.option("--limit", default=50, show_default=True, help="Max rows to return.")
@click.option("--offset", default=0, show_default=True, help="Row offset for pagination.")
@click.option("--max-cols", default=20, show_default=True, help="Max columns to show.")
@click.option("--col-offset", default=0, show_default=True, help="Column offset.")
def read(file, sheet, range_str, formulas, limit, offset, max_cols, col_offset):
    """Read cell values from a range on a sheet.

    Returns the values for each cell in the range. With --formulas, shows both
    the computed value and the underlying formula side by side.

    \b
    Examples:
      xlx read "Model.xlsx" "Income Statement" "A1:D20"
      xlx read "Model.xlsx" "Summary" "B4:C4" --formulas
      xlx read "Model.xlsx" "Data" "A1:Z100" --limit 20 --offset 40
    """
    try:
        click.echo(explorer.read_range(file, sheet, range_str, formulas=formulas,
                                       limit=limit, offset=offset,
                                       max_cols=max_cols, col_offset=col_offset))
    except Exception as exc:
        click.echo(f"Error: {exc}", err=True)
        sys.exit(1)


@cli.command("read-row")
@click.argument("file", type=click.Path(exists=True))
@click.argument("sheet")
@click.argument("row", type=int)
@click.option("--max-cols", default=20, show_default=True, help="Max columns to show.")
@click.option("--col-offset", default=0, show_default=True, help="Column offset.")
@click.option("--formulas", is_flag=True, default=False, help="Show raw formula strings.")
def read_row(file, sheet, row, max_cols, col_offset, formulas):
    """Read a full row with auto-detected column headers.

    Reads the specified row and labels each cell with its column header
    (from row 1). Natural for reading a P&L line item across time periods.

    \b
    Examples:
      xlx read-row "Model.xlsx" "Income Statement" 5
      xlx read-row "Model.xlsx" "Summary" 10 --formulas --max-cols 12
    """
    try:
        click.echo(explorer.read_row(file, sheet, row, max_cols=max_cols,
                                     col_offset=col_offset, formulas=formulas))
    except Exception as exc:
        click.echo(f"Error: {exc}", err=True)
        sys.exit(1)


@cli.command("read-col")
@click.argument("file", type=click.Path(exists=True))
@click.argument("sheet")
@click.argument("col")
@click.option("--limit", default=50, show_default=True, help="Max rows to return.")
@click.option("--offset", default=0, show_default=True, help="Row offset for pagination.")
@click.option("--formulas", is_flag=True, default=False, help="Show raw formula strings.")
def read_col(file, sheet, col, limit, offset, formulas):
    """Read a full column.

    Reads all values down a single column. Good for reading line item labels
    (column A) or a single period's data.

    \b
    Examples:
      xlx read-col "Model.xlsx" "Income Statement" A --limit 30
      xlx read-col "Model.xlsx" "Summary" B --formulas
    """
    try:
        click.echo(explorer.read_col(file, sheet, col, limit=limit, offset=offset,
                                     formulas=formulas))
    except Exception as exc:
        click.echo(f"Error: {exc}", err=True)
        sys.exit(1)


@cli.command()
@click.argument("file", type=click.Path(exists=True))
@click.argument("sheet")
@click.option("--limit", default=50, show_default=True, help="Max tables to return.")
@click.option("--offset", default=0, show_default=True, help="Pagination offset.")
def tables(file, sheet, limit, offset):
    """Auto-detect contiguous table regions on a sheet.

    Scans for blocks of non-empty rows separated by blank rows. Returns each
    table's range, dimensions, and detected header row. Useful for understanding
    sheets with multiple logical sections.

    \b
    Example:
      xlx tables "Model.xlsx" "Income Statement"
    """
    try:
        click.echo(explorer.detect_tables(file, sheet, limit=limit, offset=offset))
    except Exception as exc:
        click.echo(f"Error: {exc}", err=True)
        sys.exit(1)


# ---------------------------------------------------------------------------
# Tier 3 — Dependencies
# ---------------------------------------------------------------------------

@cli.command()
@click.argument("file", type=click.Path(exists=True))
@click.argument("cell_ref")
@click.option("--sheet", default=None, help="Sheet name (if not embedded in CELL_REF).")
@click.option("--depth", default=5, show_default=True, help="Max recursion depth.")
def trace(file, cell_ref, sheet, depth):
    """Trace the full dependency tree from a cell down to its leaf inputs.

    Recursively walks formula references across sheets, building an indented
    tree. Leaf nodes (hardcoded values that other cells depend on) are marked
    [INPUT]. Shared dependencies are flagged. This is the key command for
    understanding how a number in a financial model is calculated.

    CELL_REF can be 'Sheet!A1' or just 'A1' with --sheet.

    \b
    Examples:
      xlx trace "Model.xlsx" "Summary!B15"
      xlx trace "Model.xlsx" B4 --sheet Summary --depth 3
    """
    try:
        sheet_name, cell = _parse_cell_ref(cell_ref, sheet)
        click.echo(dependencies.trace_cell(file, sheet_name, cell, depth=depth))
    except click.UsageError:
        raise
    except Exception as exc:
        click.echo(f"Error: {exc}", err=True)
        sys.exit(1)


@cli.command()
@click.argument("file", type=click.Path(exists=True))
@click.argument("cell_ref")
@click.option("--sheet", default=None, help="Sheet name (if not embedded in CELL_REF).")
@click.option("--depth", default=5, show_default=True, help="Max recursion depth.")
def dependents(file, cell_ref, sheet, depth):
    """Find all cells that depend on a given cell (reverse trace).

    The opposite of trace — shows what would be affected if you changed this
    cell. Walks successors across sheets up to --depth levels.

    CELL_REF can be 'Sheet!A1' or just 'A1' with --sheet.

    \b
    Examples:
      xlx dependents "Model.xlsx" "Assumptions!B2"
      xlx dependents "Model.xlsx" B5 --sheet Summary --depth 2
    """
    try:
        sheet_name, cell = _parse_cell_ref(cell_ref, sheet)
        click.echo(dependencies.find_dependents(file, sheet_name, cell, depth=depth))
    except click.UsageError:
        raise
    except Exception as exc:
        click.echo(f"Error: {exc}", err=True)
        sys.exit(1)


@cli.command("formula-map")
@click.argument("file", type=click.Path(exists=True))
@click.argument("sheet")
@click.option("--limit", default=50, show_default=True, help="Max patterns to return.")
@click.option("--offset", default=0, show_default=True, help="Pagination offset.")
def formula_map(file, sheet, limit, offset):
    """Show unique formula patterns on a sheet, grouped by structure.

    Normalizes formulas by replacing specific cell references with placeholders,
    then groups identical patterns. Reveals the model's logic without reading
    every cell — e.g., seeing that 200 cells all use =SUMPRODUCT(Sales,...).

    \b
    Examples:
      xlx formula-map "Model.xlsx" "Income Statement"
      xlx formula-map "Analysis.xlsx" "Revenue Detail" --limit 5
    """
    try:
        click.echo(dependencies.formula_map(file, sheet, limit=limit, offset=offset))
    except Exception as exc:
        click.echo(f"Error: {exc}", err=True)
        sys.exit(1)


@cli.command("find-inputs")
@click.argument("file", type=click.Path(exists=True))
@click.option("--sheet", default=None, help="Limit to a specific sheet.")
@click.option("--limit", default=50, show_default=True, help="Max inputs to return.")
@click.option("--offset", default=0, show_default=True, help="Pagination offset.")
def find_inputs(file, sheet, limit, offset):
    """Find hardcoded cells that other cells depend on (model inputs/assumptions).

    Scans the dependency graph for cells with no formula that have successors.
    These are the model's drivers — changing them ripples through the model.
    Results are sorted by impact (most dependents first).

    \b
    Examples:
      xlx find-inputs "Model.xlsx"
      xlx find-inputs "Model.xlsx" --sheet Assumptions
    """
    try:
        click.echo(dependencies.find_inputs(file, sheet=sheet, limit=limit, offset=offset))
    except Exception as exc:
        click.echo(f"Error: {exc}", err=True)
        sys.exit(1)


@cli.command("sheet-flow")
@click.argument("file", type=click.Path(exists=True))
def sheet_flow(file):
    """Map which sheets reference which other sheets.

    Parses cross-sheet formula references to build a directed graph of
    sheet dependencies. Identifies source sheets (data origins), sink sheets
    (final outputs), and isolated sheets (no cross-sheet refs).

    \b
    Example:
      xlx sheet-flow "Model.xlsx"
    """
    try:
        click.echo(dependencies.sheet_flow(file))
    except Exception as exc:
        click.echo(f"Error: {exc}", err=True)
        sys.exit(1)


# ---------------------------------------------------------------------------
# Tier 4 — Analysis
# ---------------------------------------------------------------------------

@cli.command("summarize-assumptions")
@click.argument("file", type=click.Path(exists=True))
@click.option("--limit", default=50, show_default=True, help="Max rows per assumption sheet.")
@click.option("--offset", default=0, show_default=True, help="Pagination offset.")
def summarize_assumptions(file, limit, offset):
    """Summarize model assumptions: named ranges and input sheets.

    Collects all named ranges and scans for sheets named "Assumptions",
    "Inputs", "Parameters", or "Drivers", extracting their parameter/value
    pairs. The first thing a PE/VC analyst does with a model.

    \b
    Example:
      xlx summarize-assumptions "Model.xlsx"
    """
    try:
        click.echo(analysis.summarize_assumptions(file, limit=limit, offset=offset))
    except Exception as exc:
        click.echo(f"Error: {exc}", err=True)
        sys.exit(1)


@cli.command("compare-periods")
@click.argument("file", type=click.Path(exists=True))
@click.argument("sheet")
@click.argument("row", type=int)
@click.option("--max-cols", default=20, show_default=True, help="Max period columns.")
@click.option("--col-offset", default=0, show_default=True, help="Column offset.")
def compare_periods(file, sheet, row, max_cols, col_offset):
    """Compare a time-series row period-over-period.

    Reads a row of data, auto-detects date headers, and computes absolute
    change and growth rate between each consecutive period. Also shows
    summary stats (min, max, average, total growth).

    \b
    Examples:
      xlx compare-periods "Model.xlsx" "Income Statement" 5
      xlx compare-periods "Model.xlsx" "Summary" 10 --max-cols 12
    """
    try:
        click.echo(analysis.compare_periods(file, sheet, row,
                                             max_cols=max_cols, col_offset=col_offset))
    except Exception as exc:
        click.echo(f"Error: {exc}", err=True)
        sys.exit(1)


@cli.command("find-anomalies")
@click.argument("file", type=click.Path(exists=True))
@click.argument("sheet")
@click.option("--limit", default=50, show_default=True, help="Max rows to scan.")
@click.option("--offset", default=0, show_default=True, help="Row offset.")
def find_anomalies(file, sheet, limit, offset):
    """Detect formula anomalies — pattern breaks and hardcoded overrides.

    Scans each row for its dominant formula pattern, then flags cells that
    deviate: different formula structure, or a hardcoded value in the middle
    of a formula row. These are often manual overrides, errors, or special
    cases worth investigating.

    \b
    Example:
      xlx find-anomalies "Model.xlsx" "Income Statement"
    """
    try:
        click.echo(analysis.find_anomalies(file, sheet, limit=limit, offset=offset))
    except Exception as exc:
        click.echo(f"Error: {exc}", err=True)
        sys.exit(1)


@cli.command("validate-balance")
@click.argument("file", type=click.Path(exists=True))
@click.argument("sheet")
def validate_balance(file, sheet):
    """Validate that a balance sheet balances (Assets = Liabilities + Equity).

    Finds rows labeled "Total Assets", "Total Liabilities", and "Total Equity"
    (case-insensitive), then checks the equation across every period column.
    Reports PASS/FAIL per period and overall BALANCED/IMBALANCED.

    \b
    Example:
      xlx validate-balance "Model.xlsx" "Balance Sheet"
    """
    try:
        click.echo(analysis.validate_balance(file, sheet))
    except Exception as exc:
        click.echo(f"Error: {exc}", err=True)
        sys.exit(1)


# ---------------------------------------------------------------------------
# Tier 5 — Search
# ---------------------------------------------------------------------------

@cli.command()
@click.argument("file", type=click.Path(exists=True))
@click.argument("query")
@click.option("--formulas", is_flag=True, default=False, help="Also search inside formula strings.")
@click.option("--limit", default=50, show_default=True, help="Max results.")
@click.option("--offset", default=0, show_default=True, help="Pagination offset.")
def search(file, query, formulas, limit, offset):
    """Search for a value or text across all sheets.

    Case-insensitive search across every cell in the workbook. By default only
    searches cell values; with --formulas also matches inside formula strings
    (e.g., find all cells using VLOOKUP).

    \b
    Examples:
      xlx search "Model.xlsx" "Revenue"
      xlx search "Model.xlsx" "SUMPRODUCT" --formulas
      xlx search "Model.xlsx" "1000" --limit 10
    """
    try:
        click.echo(search_mod.search(file, query, formulas=formulas, limit=limit, offset=offset))
    except Exception as exc:
        click.echo(f"Error: {exc}", err=True)
        sys.exit(1)


@cli.command("find-formatting")
@click.argument("file", type=click.Path(exists=True))
@click.argument("sheet")
@click.option("--limit", default=50, show_default=True, help="Max cells to return.")
@click.option("--offset", default=0, show_default=True, help="Pagination offset.")
def find_formatting(file, sheet, limit, offset):
    """Report cells with notable formatting (bold, fill colors, borders).

    In financial models, formatting carries meaning — blue fill typically means
    an input/assumption, bold means a subtotal, borders delineate sections.
    This command surfaces those semantic signals.

    \b
    Example:
      xlx find-formatting "Model.xlsx" "Income Statement"
    """
    try:
        click.echo(search_mod.find_formatting(file, sheet, limit=limit, offset=offset))
    except Exception as exc:
        click.echo(f"Error: {exc}", err=True)
        sys.exit(1)
