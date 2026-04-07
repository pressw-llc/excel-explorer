"""Integration tests against real workbooks.

Skipped unless XLX_TEST_WORKBOOKS env var points at a directory with .xlsx files.
These tests verify that every command runs without error on real-world workbooks.
"""
import pytest
from pathlib import Path
from click.testing import CliRunner
from excel_explorer.cli import cli


@pytest.fixture
def workbooks(real_workbook_dir):
    """Collect all .xlsx files in the test workbook directory (recursively)."""
    files = list(Path(real_workbook_dir).rglob("*.xlsx"))
    if not files:
        pytest.skip("No .xlsx files found in XLX_TEST_WORKBOOKS directory")
    return files


@pytest.fixture
def first_workbook(workbooks):
    """Return the first workbook found (for single-file tests)."""
    return str(workbooks[0])


@pytest.fixture
def runner():
    return CliRunner()


class TestOrientationCommands:
    def test_overview(self, runner, first_workbook):
        result = runner.invoke(cli, ["overview", first_workbook])
        assert result.exit_code == 0
        assert "sheets:" in result.output

    def test_overview_all_workbooks(self, runner, workbooks):
        """Every workbook should produce a valid overview."""
        for wb in workbooks:
            result = runner.invoke(cli, ["overview", str(wb)])
            assert result.exit_code == 0, f"overview failed for {wb.name}: {result.output}"

    def test_describe_first_sheet(self, runner, first_workbook):
        """Describe the first sheet of the first workbook."""
        overview = runner.invoke(cli, ["overview", first_workbook])
        # Extract first sheet name from overview output
        for line in overview.output.splitlines():
            if line.strip().startswith("- name:"):
                sheet_name = line.split("- name:")[1].strip()
                break
        else:
            pytest.skip("Could not parse sheet name from overview")
        result = runner.invoke(cli, ["describe", first_workbook, sheet_name, "--limit", "10"])
        assert result.exit_code == 0
        assert "rows:" in result.output

    def test_named_ranges(self, runner, first_workbook):
        result = runner.invoke(cli, ["named-ranges", first_workbook])
        assert result.exit_code == 0


class TestReadCommands:
    def test_read_range(self, runner, first_workbook):
        """Read a small range from the first sheet."""
        overview = runner.invoke(cli, ["overview", first_workbook])
        for line in overview.output.splitlines():
            if line.strip().startswith("- name:"):
                sheet_name = line.split("- name:")[1].strip()
                break
        else:
            pytest.skip("Could not parse sheet name")
        result = runner.invoke(cli, ["read", first_workbook, sheet_name, "A1:C10"])
        assert result.exit_code == 0

    def test_read_row(self, runner, first_workbook):
        overview = runner.invoke(cli, ["overview", first_workbook])
        for line in overview.output.splitlines():
            if line.strip().startswith("- name:"):
                sheet_name = line.split("- name:")[1].strip()
                break
        else:
            pytest.skip("Could not parse sheet name")
        result = runner.invoke(cli, ["read-row", first_workbook, sheet_name, "1"])
        assert result.exit_code == 0

    def test_read_col(self, runner, first_workbook):
        overview = runner.invoke(cli, ["overview", first_workbook])
        for line in overview.output.splitlines():
            if line.strip().startswith("- name:"):
                sheet_name = line.split("- name:")[1].strip()
                break
        else:
            pytest.skip("Could not parse sheet name")
        result = runner.invoke(cli, ["read-col", first_workbook, sheet_name, "A", "--limit", "10"])
        assert result.exit_code == 0

    def test_tables(self, runner, first_workbook):
        overview = runner.invoke(cli, ["overview", first_workbook])
        for line in overview.output.splitlines():
            if line.strip().startswith("- name:"):
                sheet_name = line.split("- name:")[1].strip()
                break
        else:
            pytest.skip("Could not parse sheet name")
        result = runner.invoke(cli, ["tables", first_workbook, sheet_name])
        assert result.exit_code == 0


class TestDependencyCommands:
    def test_sheet_flow(self, runner, first_workbook):
        result = runner.invoke(cli, ["sheet-flow", first_workbook])
        assert result.exit_code == 0
        assert "sheets:" in result.output

    def test_formula_map(self, runner, first_workbook):
        """Find a sheet with formulas and run formula-map on it."""
        overview = runner.invoke(cli, ["overview", first_workbook])
        for line in overview.output.splitlines():
            if line.strip().startswith("- name:"):
                sheet_name = line.split("- name:")[1].strip()
            if line.strip().startswith("formulas:"):
                count = int(line.split("formulas:")[1].strip())
                if count > 0:
                    result = runner.invoke(cli, ["formula-map", first_workbook, sheet_name, "--limit", "5"])
                    assert result.exit_code == 0
                    return
        pytest.skip("No sheets with formulas found")


class TestSearchCommands:
    def test_search(self, runner, first_workbook):
        result = runner.invoke(cli, ["search", first_workbook, "a", "--limit", "5"])
        assert result.exit_code == 0

    def test_search_formulas(self, runner, first_workbook):
        result = runner.invoke(cli, ["search", first_workbook, "SUM", "--formulas", "--limit", "5"])
        assert result.exit_code == 0

    def test_find_formatting(self, runner, first_workbook):
        overview = runner.invoke(cli, ["overview", first_workbook])
        for line in overview.output.splitlines():
            if line.strip().startswith("- name:"):
                sheet_name = line.split("- name:")[1].strip()
                break
        else:
            pytest.skip("Could not parse sheet name")
        result = runner.invoke(cli, ["find-formatting", first_workbook, sheet_name, "--limit", "5"])
        assert result.exit_code == 0
