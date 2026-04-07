import pytest
from excel_explorer.workbook import load_workbook, get_cell_value, get_named_ranges, get_sheet_names


def test_load_workbook(tmp_workbook):
    wb = load_workbook(str(tmp_workbook))
    assert wb is not None


def test_load_workbook_not_found():
    with pytest.raises(FileNotFoundError):
        load_workbook("/nonexistent/file.xlsx")


def test_get_sheet_names(tmp_workbook):
    wb = load_workbook(str(tmp_workbook))
    names = get_sheet_names(wb)
    assert "Summary" in names
    assert "Assumptions" in names
    assert "Projections" in names


def test_get_cell_value(tmp_workbook):
    wb = load_workbook(str(tmp_workbook))
    val = get_cell_value(wb, "Summary", "A2")
    assert val == "Revenue"


def test_get_cell_formula(tmp_workbook):
    wb = load_workbook(str(tmp_workbook))
    ws = wb["Summary"]
    assert ws["B4"].value == "=B2-B3"


def test_get_named_ranges_empty(tmp_workbook):
    wb = load_workbook(str(tmp_workbook))
    ranges = get_named_ranges(wb)
    assert isinstance(ranges, dict)
