from excel_explorer.explorer import overview, describe, named_ranges, read_range, read_row, read_col, detect_tables


def test_overview(tmp_workbook):
    result = overview(str(tmp_workbook))
    assert "Summary" in result
    assert "Assumptions" in result
    assert "Projections" in result
    assert "sheets: 3" in result


def test_overview_shows_formula_count(tmp_workbook):
    result = overview(str(tmp_workbook))
    assert "formulas:" in result


def test_describe_sheet(tmp_workbook):
    result = describe(str(tmp_workbook), "Summary")
    assert "Summary" in result
    assert "rows:" in result
    assert "cols:" in result


def test_describe_shows_headers(tmp_workbook):
    result = describe(str(tmp_workbook), "Summary")
    assert "Metric" in result


def test_named_ranges_empty(tmp_workbook):
    result = named_ranges(str(tmp_workbook))
    assert "named_ranges: 0" in result or "no named ranges" in result.lower()


def test_read_range(tmp_workbook):
    result = read_range(str(tmp_workbook), "Summary", "A1:C3")
    assert "Revenue" in result
    assert "1000" in result


def test_read_range_with_formulas(tmp_workbook):
    result = read_range(str(tmp_workbook), "Summary", "B4:C4", formulas=True)
    assert "=B2-B3" in result


def test_read_range_pagination(tmp_workbook):
    result = read_range(str(tmp_workbook), "Summary", "A1:C6", limit=2, offset=0)
    assert "rows 1-2 of 6" in result
    assert "truncated: true" in result


def test_read_row(tmp_workbook):
    result = read_row(str(tmp_workbook), "Summary", 2)
    assert "Revenue" in result
    assert "1000" in result


def test_read_col(tmp_workbook):
    result = read_col(str(tmp_workbook), "Summary", "A")
    assert "Metric" in result
    assert "Revenue" in result
    assert "COGS" in result


def test_detect_tables(tmp_workbook):
    result = detect_tables(str(tmp_workbook), "Summary")
    assert "tables_found:" in result
    assert "headers:" in result
