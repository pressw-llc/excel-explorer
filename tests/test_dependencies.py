import pytest
from excel_explorer.dependencies import (
    build_dag,
    trace_cell,
    find_dependents,
    formula_map,
    find_inputs,
    sheet_flow,
)


def test_build_dag(tmp_workbook):
    dag = build_dag(str(tmp_workbook))
    assert dag is not None
    assert hasattr(dag, "dsp")


def test_trace_cell_simple(tmp_workbook):
    result = trace_cell(str(tmp_workbook), "Summary", "B4")
    # B4 = B2 - B3, both are hardcoded values
    assert "B2" in result
    assert "B3" in result
    assert "[INPUT]" in result


def test_trace_cell_nested(tmp_workbook):
    result = trace_cell(str(tmp_workbook), "Summary", "B5")
    # B5 = B4/B2, B4 = B2-B3, so should see B2, B3 as leaves
    assert "B4" in result
    assert "B2" in result


def test_trace_cell_cross_sheet(tmp_workbook):
    result = trace_cell(str(tmp_workbook), "Projections", "B3")
    # B3 = B2*(1+Assumptions!B2), B2 = Summary!B2
    assert "Assumptions" in result or "assumptions" in result.lower()


def test_trace_cell_depth_limit(tmp_workbook):
    result = trace_cell(str(tmp_workbook), "Summary", "B5", depth=1)
    assert "depth: 1" in result


def test_find_dependents(tmp_workbook):
    result = find_dependents(str(tmp_workbook), "Summary", "B2")
    # B2 is used by B4 (=B2-B3), B5 (=B4/B2), B6 (=SUM(B2:C2)), and Projections!B2
    assert "B4" in result or "B5" in result or "B6" in result


def test_formula_map(tmp_workbook):
    result = formula_map(str(tmp_workbook), "Summary")
    assert "pattern" in result.lower() or "=" in result


def test_find_inputs(tmp_workbook):
    result = find_inputs(str(tmp_workbook))
    assert "INPUT" in result or "input" in result.lower()


def test_find_inputs_with_sheet_filter(tmp_workbook):
    result = find_inputs(str(tmp_workbook), sheet="Assumptions")
    assert "Assumptions" in result


def test_sheet_flow(tmp_workbook):
    result = sheet_flow(str(tmp_workbook))
    assert "Summary" in result
    assert "Projections" in result
    assert "->" in result
