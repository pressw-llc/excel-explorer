from excel_explorer.search import search, find_formatting


def test_search_by_value(tmp_workbook):
    result = search(str(tmp_workbook), "Revenue")
    assert "Revenue" in result
    assert "Summary" in result


def test_search_by_number(tmp_workbook):
    result = search(str(tmp_workbook), "1000")
    assert "1000" in result


def test_search_in_formulas(tmp_workbook):
    result = search(str(tmp_workbook), "SUM", formulas=True)
    assert "SUM" in result


def test_search_no_results(tmp_workbook):
    result = search(str(tmp_workbook), "NONEXISTENT_VALUE_XYZ")
    assert "0" in result


def test_search_pagination(tmp_workbook):
    result = search(str(tmp_workbook), "Revenue", limit=1)
    assert "truncated" in result


def test_find_formatting(tmp_workbook):
    result = find_formatting(str(tmp_workbook), "Summary")
    assert "file:" in result
