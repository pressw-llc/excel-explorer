from click.testing import CliRunner
from excel_explorer.cli import cli


def test_cli_help():
    runner = CliRunner()
    result = runner.invoke(cli, ["--help"])
    assert result.exit_code == 0
    assert "Excel Explorer" in result.output or "explore" in result.output.lower()


def test_cli_overview(tmp_workbook):
    runner = CliRunner()
    result = runner.invoke(cli, ["overview", str(tmp_workbook)])
    assert result.exit_code == 0
    assert "Summary" in result.output


def test_cli_describe(tmp_workbook):
    runner = CliRunner()
    result = runner.invoke(cli, ["describe", str(tmp_workbook), "Summary"])
    assert result.exit_code == 0
    assert "rows:" in result.output


def test_cli_read(tmp_workbook):
    runner = CliRunner()
    result = runner.invoke(cli, ["read", str(tmp_workbook), "Summary", "A1:C3"])
    assert result.exit_code == 0
    assert "Revenue" in result.output


def test_cli_read_with_formulas(tmp_workbook):
    runner = CliRunner()
    result = runner.invoke(cli, ["read", str(tmp_workbook), "Summary", "B4:C4", "--formulas"])
    assert result.exit_code == 0
    assert "=B2-B3" in result.output


def test_cli_read_with_pagination(tmp_workbook):
    runner = CliRunner()
    result = runner.invoke(cli, ["read", str(tmp_workbook), "Summary", "A1:C6", "--limit", "2"])
    assert result.exit_code == 0
    assert "truncated: true" in result.output


def test_cli_read_row(tmp_workbook):
    runner = CliRunner()
    result = runner.invoke(cli, ["read-row", str(tmp_workbook), "Summary", "2"])
    assert result.exit_code == 0
    assert "Revenue" in result.output


def test_cli_read_col(tmp_workbook):
    runner = CliRunner()
    result = runner.invoke(cli, ["read-col", str(tmp_workbook), "Summary", "A"])
    assert result.exit_code == 0
    assert "Revenue" in result.output


def test_cli_tables(tmp_workbook):
    runner = CliRunner()
    result = runner.invoke(cli, ["tables", str(tmp_workbook), "Summary"])
    assert result.exit_code == 0
    assert "tables_found:" in result.output


def test_cli_named_ranges(tmp_workbook):
    runner = CliRunner()
    result = runner.invoke(cli, ["named-ranges", str(tmp_workbook)])
    assert result.exit_code == 0


def test_cli_search(tmp_workbook):
    runner = CliRunner()
    result = runner.invoke(cli, ["search", str(tmp_workbook), "Revenue"])
    assert result.exit_code == 0
    assert "Revenue" in result.output


def test_cli_sheet_flow(tmp_workbook):
    runner = CliRunner()
    result = runner.invoke(cli, ["sheet-flow", str(tmp_workbook)])
    assert result.exit_code == 0


def test_cli_formula_map(tmp_workbook):
    runner = CliRunner()
    result = runner.invoke(cli, ["formula-map", str(tmp_workbook), "Summary"])
    assert result.exit_code == 0


def test_cli_file_not_found():
    runner = CliRunner()
    result = runner.invoke(cli, ["overview", "/nonexistent/file.xlsx"])
    assert result.exit_code != 0
