from excel_explorer.formatters import format_output, format_metadata


def test_format_metadata_basic():
    meta = {"file": "test.xlsx", "sheet": "Summary"}
    result = format_metadata(meta)
    assert "file: test.xlsx" in result
    assert "sheet: Summary" in result


def test_format_metadata_truncation():
    meta = {
        "file": "test.xlsx",
        "sheet": "Summary",
        "showing": "rows 1-50 of 100",
        "truncated": True,
    }
    result = format_metadata(meta)
    assert "truncated: true" in result


def test_format_output_with_metadata():
    meta = {"file": "test.xlsx"}
    body = "some content\nmore content"
    result = format_output(meta, body)
    assert result.startswith("file: test.xlsx")
    assert "---" in result
    assert "some content" in result


def test_format_output_empty_body():
    meta = {"file": "test.xlsx"}
    result = format_output(meta, "")
    assert "file: test.xlsx" in result
    assert "---" in result
