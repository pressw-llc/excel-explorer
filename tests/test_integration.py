"""Integration tests against bundled synthetic workbooks.

Tests every command against realistic (but fully fabricated) Excel workbooks
that exercise different patterns: values-only exports, formula-heavy models,
SUMPRODUCT aggregations, and transaction-level data.
"""
import pytest
from click.testing import CliRunner
from excel_explorer.cli import cli


@pytest.fixture
def runner():
    return CliRunner()


# ---------------------------------------------------------------------------
# Financial Statements (values only, no formulas — like a QBO export)
# ---------------------------------------------------------------------------

class TestFinancialStatements:
    def test_overview(self, runner, financial_statements):
        result = runner.invoke(cli, ["overview", financial_statements])
        assert result.exit_code == 0
        assert "Balance Sheet" in result.output
        assert "Income Statement" in result.output
        assert "sheets: 2" in result.output

    def test_describe_balance_sheet(self, runner, financial_statements):
        result = runner.invoke(cli, ["describe", financial_statements, "Balance Sheet", "--limit", "10"])
        assert result.exit_code == 0
        assert "rows:" in result.output

    def test_read_range(self, runner, financial_statements):
        result = runner.invoke(cli, ["read", financial_statements, "Income Statement", "A6:C12"])
        assert result.exit_code == 0
        assert "Gross Sales" in result.output

    def test_read_row(self, runner, financial_statements):
        result = runner.invoke(cli, ["read-row", financial_statements, "Income Statement", "7"])
        assert result.exit_code == 0
        assert "Sales-Food" in result.output

    def test_read_col(self, runner, financial_statements):
        result = runner.invoke(cli, ["read-col", financial_statements, "Balance Sheet", "A", "--limit", "20"])
        assert result.exit_code == 0
        assert "Assets" in result.output

    def test_tables(self, runner, financial_statements):
        result = runner.invoke(cli, ["tables", financial_statements, "Balance Sheet"])
        assert result.exit_code == 0
        assert "tables_found:" in result.output

    def test_validate_balance(self, runner, financial_statements):
        result = runner.invoke(cli, ["validate-balance", financial_statements, "Balance Sheet"])
        assert result.exit_code == 0

    def test_compare_periods(self, runner, financial_statements):
        result = runner.invoke(cli, ["compare-periods", financial_statements, "Income Statement", "7"])
        assert result.exit_code == 0
        assert "change" in result.output.lower() or "growth" in result.output.lower()

    def test_search(self, runner, financial_statements):
        result = runner.invoke(cli, ["search", financial_statements, "Sales"])
        assert result.exit_code == 0
        assert "Sales" in result.output

    def test_find_formatting(self, runner, financial_statements):
        result = runner.invoke(cli, ["find-formatting", financial_statements, "Balance Sheet"])
        assert result.exit_code == 0
        assert "bold" in result.output

    def test_named_ranges_empty(self, runner, financial_statements):
        result = runner.invoke(cli, ["named-ranges", financial_statements])
        assert result.exit_code == 0
        assert "named_ranges: 0" in result.output or "No named ranges" in result.output

    def test_sheet_flow_no_formulas(self, runner, financial_statements):
        result = runner.invoke(cli, ["sheet-flow", financial_statements])
        assert result.exit_code == 0
        assert "isolated" in result.output


# ---------------------------------------------------------------------------
# Financial Model (formulas, cross-sheet refs, named ranges)
# ---------------------------------------------------------------------------

class TestFinancialModel:
    def test_overview(self, runner, financial_model):
        result = runner.invoke(cli, ["overview", financial_model])
        assert result.exit_code == 0
        assert "Assumptions" in result.output
        assert "Revenue Build" in result.output
        assert "P&L" in result.output

    def test_named_ranges(self, runner, financial_model):
        result = runner.invoke(cli, ["named-ranges", financial_model])
        assert result.exit_code == 0
        assert "GrowthRate" in result.output or "TaxRate" in result.output or "AvgCheck" in result.output

    def test_read_formulas(self, runner, financial_model):
        result = runner.invoke(cli, ["read", financial_model, "P&L", "A4:B13", "--formulas"])
        assert result.exit_code == 0
        assert "formula:" in result.output or "=" in result.output

    def test_sheet_flow(self, runner, financial_model):
        result = runner.invoke(cli, ["sheet-flow", financial_model])
        assert result.exit_code == 0
        # Assumptions and Revenue Build should be referenced by P&L
        assert "->" in result.output

    def test_formula_map(self, runner, financial_model):
        result = runner.invoke(cli, ["formula-map", financial_model, "P&L"])
        assert result.exit_code == 0
        assert "pattern" in result.output.lower()

    def test_describe_assumptions(self, runner, financial_model):
        result = runner.invoke(cli, ["describe", financial_model, "Assumptions"])
        assert result.exit_code == 0
        assert "Parameter" in result.output

    def test_summarize_assumptions(self, runner, financial_model):
        result = runner.invoke(cli, ["summarize-assumptions", financial_model])
        assert result.exit_code == 0
        assert "Assumptions" in result.output or "named_ranges" in result.output

    def test_find_anomalies(self, runner, financial_model):
        result = runner.invoke(cli, ["find-anomalies", financial_model, "P&L"])
        assert result.exit_code == 0
        # The fixture has one anomaly cell (H21 is a formula while rest of row is hardcoded)

    def test_find_formatting(self, runner, financial_model):
        result = runner.invoke(cli, ["find-formatting", financial_model, "Assumptions"])
        assert result.exit_code == 0
        # Assumptions has blue fill on input cells
        assert "fill:" in result.output

    def test_trace(self, runner, financial_model):
        result = runner.invoke(cli, ["trace", financial_model, "P&L!B13"])
        assert result.exit_code == 0

    def test_dependents(self, runner, financial_model):
        result = runner.invoke(cli, ["dependents", financial_model, "Assumptions!B4"])
        assert result.exit_code == 0

    def test_find_inputs(self, runner, financial_model):
        result = runner.invoke(cli, ["find-inputs", financial_model])
        assert result.exit_code == 0


# ---------------------------------------------------------------------------
# Sales Analysis (SUMPRODUCT, named ranges, data sheet)
# ---------------------------------------------------------------------------

class TestSalesAnalysis:
    def test_overview(self, runner, sales_analysis):
        result = runner.invoke(cli, ["overview", sales_analysis])
        assert result.exit_code == 0
        assert "Sales Data" in result.output
        assert "Store Analysis" in result.output

    def test_named_ranges(self, runner, sales_analysis):
        result = runner.invoke(cli, ["named-ranges", sales_analysis])
        assert result.exit_code == 0
        assert "SaleRevenue" in result.output
        assert "Region" in result.output

    def test_describe_data_sheet(self, runner, sales_analysis):
        result = runner.invoke(cli, ["describe", sales_analysis, "Sales Data", "--limit", "5"])
        assert result.exit_code == 0
        assert "rows:" in result.output

    def test_formula_map(self, runner, sales_analysis):
        result = runner.invoke(cli, ["formula-map", sales_analysis, "Store Analysis"])
        assert result.exit_code == 0
        assert "SUMPRODUCT" in result.output

    def test_search_store(self, runner, sales_analysis):
        result = runner.invoke(cli, ["search", sales_analysis, "Maple"])
        assert result.exit_code == 0
        assert "Maple" in result.output

    def test_search_formulas(self, runner, sales_analysis):
        result = runner.invoke(cli, ["search", sales_analysis, "SUMPRODUCT", "--formulas"])
        assert result.exit_code == 0
        assert "SUMPRODUCT" in result.output

    def test_sheet_flow(self, runner, sales_analysis):
        result = runner.invoke(cli, ["sheet-flow", sales_analysis])
        assert result.exit_code == 0


# ---------------------------------------------------------------------------
# OpEx Detail (transactions, vendor parsing, VLOOKUP, summaries)
# ---------------------------------------------------------------------------

class TestOpExDetail:
    def test_overview(self, runner, opex_detail):
        result = runner.invoke(cli, ["overview", opex_detail])
        assert result.exit_code == 0
        assert "Transactions" in result.output
        assert "Summary" in result.output

    def test_describe_transactions(self, runner, opex_detail):
        result = runner.invoke(cli, ["describe", opex_detail, "Transactions", "--limit", "5"])
        assert result.exit_code == 0
        assert "rows:" in result.output

    def test_formula_map_transactions(self, runner, opex_detail):
        result = runner.invoke(cli, ["formula-map", opex_detail, "Transactions"])
        assert result.exit_code == 0
        assert "TRIM" in result.output

    def test_read_summary(self, runner, opex_detail):
        result = runner.invoke(cli, ["read", opex_detail, "Summary", "A1:G15"])
        assert result.exit_code == 0

    def test_search_vendor(self, runner, opex_detail):
        result = runner.invoke(cli, ["search", opex_detail, "Sysco", "--limit", "5"])
        assert result.exit_code == 0
        assert "Sysco" in result.output

    def test_tables(self, runner, opex_detail):
        result = runner.invoke(cli, ["tables", opex_detail, "Summary"])
        assert result.exit_code == 0
