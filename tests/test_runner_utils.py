"""Tests for pure utility functions in excelbench.harness.runner."""

from __future__ import annotations

from datetime import date, datetime
from typing import Any

import pytest

from excelbench.harness.runner import (
    _border_from_expected,
    _cell_format_from_expected,
    _cell_to_coord,
    _cell_value_from_expected,
    _cell_value_from_raw,
    _cells_in_range,
    _coord_to_cell,
    _deep_compare,
    _extract_column,
    _extract_formula_sheet_names,
    _find_by_key,
    _find_range,
    _find_rule,
    _find_validation,
    _first_non_top_left_cell,
    _normalize_formula,
    _normalize_number_format,
    _normalize_range,
    _normalize_sheet_quotes,
    _project_rule,
    _split_range,
    compare_results,
)
from excelbench.models import (
    BorderStyle,
    CellType,
)

JSONDict = dict[str, Any]


# ─────────────────────────────────────────────────
# _cell_to_coord
# ─────────────────────────────────────────────────


@pytest.mark.parametrize(
    "cell, expected",
    [
        ("A1", (1, 1)),
        ("B3", (3, 2)),
        ("Z1", (1, 26)),
        ("AA1", (1, 27)),
        ("AZ1", (1, 52)),
        ("BA1", (1, 53)),
        ("a1", (1, 1)),  # lowercase
        ("C100", (100, 3)),
    ],
)
def test_cell_to_coord(cell: str, expected: tuple[int, int]) -> None:
    assert _cell_to_coord(cell) == expected


def test_cell_to_coord_invalid_returns_default() -> None:
    assert _cell_to_coord("123") == (1, 1)
    assert _cell_to_coord("") == (1, 1)


# ─────────────────────────────────────────────────
# _coord_to_cell
# ─────────────────────────────────────────────────


@pytest.mark.parametrize(
    "row, col, expected",
    [
        (1, 1, "A1"),
        (3, 2, "B3"),
        (1, 26, "Z1"),
        (1, 27, "AA1"),
        (1, 52, "AZ1"),
        (1, 53, "BA1"),
        (100, 3, "C100"),
    ],
)
def test_coord_to_cell(row: int, col: int, expected: str) -> None:
    assert _coord_to_cell(row, col) == expected


@pytest.mark.parametrize("cell", ["A1", "B3", "Z1", "AA1", "AZ1", "BA1", "C100"])
def test_coord_roundtrip(cell: str) -> None:
    row, col = _cell_to_coord(cell)
    assert _coord_to_cell(row, col) == cell


# ─────────────────────────────────────────────────
# _extract_column
# ─────────────────────────────────────────────────


@pytest.mark.parametrize(
    "cell, expected",
    [
        ("A1", "A"),
        ("AB123", "AB"),
        ("z5", "Z"),
    ],
)
def test_extract_column(cell: str, expected: str) -> None:
    assert _extract_column(cell) == expected


def test_extract_column_no_letter_returns_default() -> None:
    assert _extract_column("123") == "B"


# ─────────────────────────────────────────────────
# _split_range
# ─────────────────────────────────────────────────


@pytest.mark.parametrize(
    "range_str, expected",
    [
        ("A1:B2", ("A1", "B2")),
        ("$A$1:$B$2", ("A1", "B2")),
        ("A1", ("A1", "A1")),
        ("$C$3", ("C3", "C3")),
    ],
)
def test_split_range(range_str: str, expected: tuple[str, str]) -> None:
    assert _split_range(range_str) == expected


# ─────────────────────────────────────────────────
# _cells_in_range
# ─────────────────────────────────────────────────


def test_cells_in_range_single_cell() -> None:
    assert _cells_in_range("A1", "A1") == ["A1"]


def test_cells_in_range_row() -> None:
    assert _cells_in_range("A1", "C1") == ["A1", "B1", "C1"]


def test_cells_in_range_column() -> None:
    assert _cells_in_range("A1", "A3") == ["A1", "A2", "A3"]


def test_cells_in_range_block() -> None:
    cells = _cells_in_range("A1", "B2")
    assert cells == ["A1", "B1", "A2", "B2"]


# ─────────────────────────────────────────────────
# _first_non_top_left_cell
# ─────────────────────────────────────────────────


def test_first_non_top_left_cell_multi() -> None:
    assert _first_non_top_left_cell("A1", "B2") == "B1"


def test_first_non_top_left_cell_single() -> None:
    assert _first_non_top_left_cell("A1", "A1") is None


# ─────────────────────────────────────────────────
# _normalize_range
# ─────────────────────────────────────────────────


@pytest.mark.parametrize(
    "input_str, expected",
    [
        ("$A$1:$B$2", "A1:B2"),
        ("a1:b2", "A1:B2"),
        ("C3", "C3"),
    ],
)
def test_normalize_range(input_str: str, expected: str) -> None:
    assert _normalize_range(input_str) == expected


# ─────────────────────────────────────────────────
# _normalize_formula
# ─────────────────────────────────────────────────


@pytest.mark.parametrize(
    "value, expected",
    [
        ("=SUM(A1:A3)", "SUM(A1:A3)"),
        ("SUM(A1:A3)", "SUM(A1:A3)"),
        ('"hello"', "hello"),
        ('="hello"', "hello"),
        (42, 42),
        (None, None),
        ("  =A1  ", "A1"),
    ],
)
def test_normalize_formula(value: Any, expected: Any) -> None:
    assert _normalize_formula(value) == expected


# ─────────────────────────────────────────────────
# _normalize_sheet_quotes
# ─────────────────────────────────────────────────


@pytest.mark.parametrize(
    "formula, expected",
    [
        ("=References!B2", "='References'!B2"),
        ("='Already Quoted'!B2", "='Already Quoted'!B2"),
        ("=SUM(A1:A3)", "=SUM(A1:A3)"),  # no sheet ref, unchanged
        ("=Sheet1!$A$1", "='Sheet1'!$A$1"),
    ],
)
def test_normalize_sheet_quotes(formula: str, expected: str) -> None:
    assert _normalize_sheet_quotes(formula) == expected


# ─────────────────────────────────────────────────
# _normalize_number_format
# ─────────────────────────────────────────────────


@pytest.mark.parametrize(
    "fmt, expected",
    [
        (r"yyyy\-mm\-dd", "yyyy-mm-dd"),
        ('"$"#,##0.00', "$#,##0.00"),
        ('"USD" 0.00', '"USD" 0.00'),  # multi-char quoted preserved
        ("0.00%", "0.00%"),
        (r"h\:mm\:ss", "h:mm:ss"),
        (r"mm\/dd\/yyyy", "mm/dd/yyyy"),
    ],
)
def test_normalize_number_format(fmt: str, expected: str) -> None:
    assert _normalize_number_format(fmt) == expected


# ─────────────────────────────────────────────────
# _extract_formula_sheet_names
# ─────────────────────────────────────────────────


def test_extract_formula_sheet_names_quoted() -> None:
    assert _extract_formula_sheet_names("='Sheet A'!B1") == ["Sheet A"]


def test_extract_formula_sheet_names_unquoted() -> None:
    names = _extract_formula_sheet_names("=Sheet1!A1+Sheet2!B1")
    assert "Sheet1" in names
    assert "Sheet2" in names


def test_extract_formula_sheet_names_empty() -> None:
    assert _extract_formula_sheet_names("") == []
    assert _extract_formula_sheet_names("=SUM(A1:A3)") == []


# ─────────────────────────────────────────────────
# _find_range
# ─────────────────────────────────────────────────


def test_find_range_matches_normalized() -> None:
    ranges = ["$A$1:$B$2", "C3:D4"]
    assert _find_range(ranges, "A1:B2") == "$A$1:$B$2"


def test_find_range_no_match() -> None:
    assert _find_range(["A1:B2"], "C3:D4") is None


# ─────────────────────────────────────────────────
# _find_by_key
# ─────────────────────────────────────────────────


def test_find_by_key_found() -> None:
    items: list[JSONDict] = [{"id": 1, "name": "a"}, {"id": 2, "name": "b"}]
    assert _find_by_key(items, "id", 2) == {"id": 2, "name": "b"}


def test_find_by_key_not_found() -> None:
    items: list[JSONDict] = [{"id": 1}]
    assert _find_by_key(items, "id", 99) is None


# ─────────────────────────────────────────────────
# _find_rule
# ─────────────────────────────────────────────────


def test_find_rule_by_range_and_type() -> None:
    rules: list[JSONDict] = [
        {"range": "B2:B6", "rule_type": "cellIs"},
        {"range": "C1:C3", "rule_type": "containsText"},
    ]
    expected: JSONDict = {"range": "B2:B6", "rule_type": "cellIs"}
    assert _find_rule(rules, expected) == rules[0]


def test_find_rule_no_match() -> None:
    rules: list[JSONDict] = [{"range": "A1:A5", "rule_type": "dataBar"}]
    expected: JSONDict = {"range": "A1:A5", "rule_type": "cellIs"}
    assert _find_rule(rules, expected) is None


def test_find_rule_formula_mismatch() -> None:
    rules: list[JSONDict] = [{"range": "A1:A5", "formula": "=A1>10"}]
    expected: JSONDict = {"range": "A1:A5", "formula": "=B1>10"}
    assert _find_rule(rules, expected) is None


# ─────────────────────────────────────────────────
# _find_validation
# ─────────────────────────────────────────────────


def test_find_validation_match() -> None:
    validations: list[JSONDict] = [
        {"range": "A1:A10", "validation_type": "list", "formula1": '"a,b,c"'},
    ]
    expected: JSONDict = {"range": "$A$1:$A$10", "validation_type": "list"}
    assert _find_validation(validations, expected) == validations[0]


def test_find_validation_formula_mismatch() -> None:
    validations: list[JSONDict] = [
        {"range": "A1:A10", "validation_type": "list", "formula1": '"a,b,c"'},
    ]
    expected: JSONDict = {"range": "A1:A10", "formula1": '"x,y,z"'}
    assert _find_validation(validations, expected) is None


# ─────────────────────────────────────────────────
# _project_rule
# ─────────────────────────────────────────────────


def test_project_rule_keeps_expected_keys_only() -> None:
    actual: JSONDict = {"range": "A1:B2", "rule_type": "cellIs", "priority": 5}
    expected: JSONDict = {"range": "A1:B2", "rule_type": "cellIs"}
    result = _project_rule(actual, expected)
    assert result == {"range": "A1:B2", "rule_type": "cellIs"}
    assert "priority" not in result


def test_project_rule_path_fallback() -> None:
    actual: JSONDict = {"range": "A1:B2"}
    expected: JSONDict = {"range": "A1:B2", "path": "/img.png"}
    result = _project_rule(actual, expected)
    assert result["path"] == "/img.png"


# ─────────────────────────────────────────────────
# _deep_compare (extended)
# ─────────────────────────────────────────────────


def test_deep_compare_tuple_vs_list() -> None:
    assert _deep_compare((1, 2), [1, 2])


def test_deep_compare_tuple_vs_non_sequence() -> None:
    assert not _deep_compare((1, 2), 3)


def test_deep_compare_empty_collections() -> None:
    assert _deep_compare({}, {})
    assert _deep_compare([], [])


def test_deep_compare_hash_color_case() -> None:
    assert _deep_compare("#ff0000", "#FF0000")


def test_deep_compare_non_hash_string_case_sensitive() -> None:
    assert not _deep_compare("hello", "HELLO")


def test_deep_compare_dict_vs_non_dict() -> None:
    assert not _deep_compare({"a": 1}, [1])


def test_deep_compare_list_vs_non_list() -> None:
    assert not _deep_compare([1, 2], "nope")


def test_deep_compare_nested_list_in_dict() -> None:
    expected: JSONDict = {"items": [{"v": 1}, {"v": 2}]}
    actual: JSONDict = {"items": [{"v": 2}, {"v": 1}]}
    assert _deep_compare(expected, actual)


# ─────────────────────────────────────────────────
# compare_results (extended)
# ─────────────────────────────────────────────────


def test_compare_results_error_in_actual() -> None:
    assert not compare_results({}, {"error": "something broke"})


def test_compare_results_exact_match() -> None:
    assert compare_results({"val": "hello"}, {"val": "hello"})


# ─────────────────────────────────────────────────
# _cell_value_from_expected
# ─────────────────────────────────────────────────


def test_cell_value_from_expected_string() -> None:
    cv = _cell_value_from_expected({"type": "string", "value": "hello"})
    assert cv.type == CellType.STRING
    assert cv.value == "hello"


def test_cell_value_from_expected_blank() -> None:
    cv = _cell_value_from_expected({"type": "blank"})
    assert cv.type == CellType.BLANK
    assert cv.value is None


def test_cell_value_from_expected_boolean() -> None:
    cv = _cell_value_from_expected({"type": "boolean", "value": True})
    assert cv.type == CellType.BOOLEAN
    assert cv.value is True


def test_cell_value_from_expected_number() -> None:
    cv = _cell_value_from_expected({"type": "number", "value": 42.5})
    assert cv.type == CellType.NUMBER
    assert cv.value == 42.5


def test_cell_value_from_expected_date_string() -> None:
    cv = _cell_value_from_expected({"type": "date", "value": "2026-01-15"})
    assert cv.type == CellType.DATE
    assert cv.value == date(2026, 1, 15)


def test_cell_value_from_expected_datetime_string() -> None:
    cv = _cell_value_from_expected({"type": "datetime", "value": "2026-01-15T10:30:00"})
    assert cv.type == CellType.DATETIME
    assert cv.value == datetime(2026, 1, 15, 10, 30, 0)


def test_cell_value_from_expected_error() -> None:
    cv = _cell_value_from_expected({"type": "error", "value": "#DIV/0!"})
    assert cv.type == CellType.ERROR
    assert cv.value == "#DIV/0!"


def test_cell_value_from_expected_formula() -> None:
    cv = _cell_value_from_expected(
        {"type": "formula", "value": 10, "formula": "=SUM(A1:A3)"}
    )
    assert cv.type == CellType.FORMULA
    assert cv.value == 10
    assert cv.formula == "=SUM(A1:A3)"


def test_cell_value_from_expected_default_type() -> None:
    cv = _cell_value_from_expected({"value": "test"})
    assert cv.type == CellType.STRING


# ─────────────────────────────────────────────────
# _cell_value_from_raw
# ─────────────────────────────────────────────────


def test_cell_value_from_raw_none() -> None:
    cv = _cell_value_from_raw(None)
    assert cv.type == CellType.BLANK


def test_cell_value_from_raw_bool() -> None:
    cv = _cell_value_from_raw(True)
    assert cv.type == CellType.BOOLEAN
    assert cv.value is True


def test_cell_value_from_raw_int() -> None:
    cv = _cell_value_from_raw(42)
    assert cv.type == CellType.NUMBER
    assert cv.value == 42


def test_cell_value_from_raw_float() -> None:
    cv = _cell_value_from_raw(3.14)
    assert cv.type == CellType.NUMBER
    assert cv.value == 3.14


def test_cell_value_from_raw_string() -> None:
    cv = _cell_value_from_raw("hello")
    assert cv.type == CellType.STRING
    assert cv.value == "hello"


# ─────────────────────────────────────────────────
# _cell_format_from_expected
# ─────────────────────────────────────────────────


def test_cell_format_from_expected_full() -> None:
    data: JSONDict = {
        "bold": True,
        "italic": False,
        "underline": "single",
        "strikethrough": True,
        "font_name": "Arial",
        "font_size": 12.0,
        "font_color": "#FF0000",
        "bg_color": "#00FF00",
        "number_format": "0.00",
        "h_align": "center",
        "v_align": "top",
        "wrap": True,
        "rotation": 45,
        "indent": 2,
    }
    fmt = _cell_format_from_expected(data)
    assert fmt.bold is True
    assert fmt.italic is False
    assert fmt.underline == "single"
    assert fmt.strikethrough is True
    assert fmt.font_name == "Arial"
    assert fmt.font_size == 12.0
    assert fmt.font_color == "#FF0000"
    assert fmt.bg_color == "#00FF00"
    assert fmt.number_format == "0.00"
    assert fmt.h_align == "center"
    assert fmt.v_align == "top"
    assert fmt.wrap is True
    assert fmt.rotation == 45
    assert fmt.indent == 2


def test_cell_format_from_expected_empty() -> None:
    fmt = _cell_format_from_expected({})
    assert fmt.bold is None
    assert fmt.font_name is None


# ─────────────────────────────────────────────────
# _border_from_expected
# ─────────────────────────────────────────────────


def test_border_from_expected_uniform() -> None:
    data: JSONDict = {"border_style": "thin", "border_color": "#FF0000"}
    border = _border_from_expected(data)
    assert border.top is not None
    assert border.top.style == BorderStyle.THIN
    assert border.top.color == "#FF0000"
    assert border.bottom is not None
    assert border.bottom.style == BorderStyle.THIN
    assert border.left is not None
    assert border.right is not None


def test_border_from_expected_per_edge() -> None:
    data: JSONDict = {"border_top": "thick", "border_bottom": "thin"}
    border = _border_from_expected(data)
    assert border.top is not None
    assert border.top.style == BorderStyle.THICK
    assert border.bottom is not None
    assert border.bottom.style == BorderStyle.THIN
    assert border.left is None
    assert border.right is None


def test_border_from_expected_color_implies_thin() -> None:
    data: JSONDict = {"border_color": "#0000FF"}
    border = _border_from_expected(data)
    assert border.top is not None
    assert border.top.style == BorderStyle.THIN
    assert border.top.color == "#0000FF"


def test_border_from_expected_no_border() -> None:
    data: JSONDict = {}
    border = _border_from_expected(data)
    assert border.top is None
    assert border.bottom is None
    assert border.left is None
    assert border.right is None


def test_border_from_expected_default_color() -> None:
    data: JSONDict = {"border_style": "medium"}
    border = _border_from_expected(data)
    assert border.top is not None
    assert border.top.color == "#000000"


def test_border_from_expected_diagonal() -> None:
    data: JSONDict = {"border_diagonal_up": "thin", "border_diagonal_down": "double"}
    border = _border_from_expected(data)
    assert border.diagonal_up is not None
    assert border.diagonal_up.style == BorderStyle.THIN
    assert border.diagonal_down is not None
    assert border.diagonal_down.style == BorderStyle.DOUBLE


def test_border_from_expected_edge_color_no_style() -> None:
    data: JSONDict = {"border_top_color": "#FF0000"}
    border = _border_from_expected(data)
    assert border.top is not None
    assert border.top.style == BorderStyle.THIN
    assert border.top.color == "#FF0000"
