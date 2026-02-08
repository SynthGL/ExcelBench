"""Tests for runner data transformation and normalization functions."""

from __future__ import annotations

from datetime import date, datetime
from typing import Any

from excelbench.harness.runner import (
    _border_from_expected,
    _cell_format_from_expected,
    _cell_value_from_expected,
    _cell_value_from_raw,
    _extract_formula_sheet_names,
    _find_rule,
    _normalize_formula,
    _normalize_number_format,
    _normalize_sheet_quotes,
    _project_rule,
)
from excelbench.models import CellType

JSONDict = dict[str, Any]


# ═════════════════════════════════════════════════
# _normalize_formula
# ═════════════════════════════════════════════════


class TestNormalizeFormula:
    def test_non_string(self) -> None:
        assert _normalize_formula(42) == 42
        assert _normalize_formula(None) is None

    def test_strip_equals(self) -> None:
        assert _normalize_formula("=1+1") == "1+1"

    def test_strip_quotes(self) -> None:
        assert _normalize_formula('"hello"') == "hello"

    def test_strip_both(self) -> None:
        assert _normalize_formula('="SUM(A1)"') == "SUM(A1)"

    def test_plain_text(self) -> None:
        assert _normalize_formula("SUM(A1:A5)") == "SUM(A1:A5)"

    def test_whitespace(self) -> None:
        assert _normalize_formula("  =1+1  ") == "1+1"


# ═════════════════════════════════════════════════
# _normalize_number_format
# ═════════════════════════════════════════════════


class TestNormalizeNumberFormat:
    def test_backslash_escapes(self) -> None:
        assert _normalize_number_format("yyyy\\-mm\\-dd") == "yyyy-mm-dd"

    def test_single_char_quotes(self) -> None:
        assert _normalize_number_format('"$"#,##0.00') == "$#,##0.00"

    def test_no_change(self) -> None:
        assert _normalize_number_format("#,##0.00") == "#,##0.00"

    def test_backslash_space(self) -> None:
        assert _normalize_number_format("0\\ %") == "0 %"

    def test_mixed(self) -> None:
        assert _normalize_number_format('"$"#,##0\\-00') == "$#,##0-00"


# ═════════════════════════════════════════════════
# _normalize_sheet_quotes
# ═════════════════════════════════════════════════


class TestNormalizeSheetQuotes:
    def test_unquoted(self) -> None:
        assert _normalize_sheet_quotes("=References!B2") == "='References'!B2"

    def test_already_quoted(self) -> None:
        result = _normalize_sheet_quotes("='References'!B2")
        assert result == "='References'!B2"

    def test_no_sheet_ref(self) -> None:
        assert _normalize_sheet_quotes("=SUM(A1:A5)") == "=SUM(A1:A5)"

    def test_dollar_signs(self) -> None:
        assert _normalize_sheet_quotes("=Data!$A$1") == "='Data'!$A$1"


# ═════════════════════════════════════════════════
# _extract_formula_sheet_names
# ═════════════════════════════════════════════════


class TestExtractFormulaSheetNames:
    def test_empty(self) -> None:
        assert _extract_formula_sheet_names("") == []

    def test_quoted(self) -> None:
        result = _extract_formula_sheet_names("='My Sheet'!A1")
        assert "My Sheet" in result

    def test_unquoted(self) -> None:
        result = _extract_formula_sheet_names("=Data!A1")
        assert "Data" in result

    def test_multiple(self) -> None:
        result = _extract_formula_sheet_names("='Sheet1'!A1+Data!B2")
        assert "Sheet1" in result
        assert "Data" in result

    def test_no_duplicates(self) -> None:
        result = _extract_formula_sheet_names("='Data'!A1+Data!B2")
        assert result.count("Data") == 1


# ═════════════════════════════════════════════════
# _find_rule
# ═════════════════════════════════════════════════


class TestFindRule:
    def test_match_by_range(self) -> None:
        rules: list[JSONDict] = [
            {"range": "A1:A5", "rule_type": "cellIs"},
            {"range": "B1:B5", "rule_type": "colorScale"},
        ]
        result = _find_rule(rules, {"range": "B1:B5"})
        assert result is not None
        assert result["rule_type"] == "colorScale"

    def test_match_by_rule_type(self) -> None:
        rules: list[JSONDict] = [
            {"range": "A1:A5", "rule_type": "cellIs"},
            {"range": "A1:A5", "rule_type": "colorScale"},
        ]
        result = _find_rule(rules, {"range": "A1:A5", "rule_type": "colorScale"})
        assert result is not None
        assert result["rule_type"] == "colorScale"

    def test_match_by_formula(self) -> None:
        rules: list[JSONDict] = [
            {"range": "A1:A5", "formula": "=10"},
            {"range": "A1:A5", "formula": "=20"},
        ]
        result = _find_rule(rules, {"range": "A1:A5", "formula": "=10"})
        assert result is not None
        assert result["formula"] == "=10"

    def test_no_match(self) -> None:
        rules: list[JSONDict] = [{"range": "A1:A5", "rule_type": "cellIs"}]
        assert _find_rule(rules, {"range": "Z1:Z5"}) is None

    def test_empty_rules(self) -> None:
        assert _find_rule([], {"range": "A1:A5"}) is None


# ═════════════════════════════════════════════════
# _project_rule
# ═════════════════════════════════════════════════


class TestProjectRule:
    def test_projects_keys(self) -> None:
        actual: JSONDict = {"a": 1, "b": 2, "c": 3}
        expected: JSONDict = {"a": 1, "b": 2}
        result = _project_rule(actual, expected)
        assert result == {"a": 1, "b": 2}
        assert "c" not in result

    def test_missing_key(self) -> None:
        actual: JSONDict = {"a": 1}
        expected: JSONDict = {"a": 1, "b": 2}
        result = _project_rule(actual, expected)
        assert result == {"a": 1, "b": None}

    def test_path_fallback(self) -> None:
        actual: JSONDict = {"cell": "A1"}
        expected: JSONDict = {"cell": "A1", "path": "test.png"}
        result = _project_rule(actual, expected)
        assert result["path"] == "test.png"


# ═════════════════════════════════════════════════
# _cell_value_from_expected
# ═════════════════════════════════════════════════


class TestCellValueFromExpected:
    def test_blank(self) -> None:
        cv = _cell_value_from_expected({"type": "blank"})
        assert cv.type == CellType.BLANK

    def test_string(self) -> None:
        cv = _cell_value_from_expected({"type": "string", "value": "Hello"})
        assert cv.type == CellType.STRING
        assert cv.value == "Hello"

    def test_number(self) -> None:
        cv = _cell_value_from_expected({"type": "number", "value": 42})
        assert cv.type == CellType.NUMBER
        assert cv.value == 42

    def test_boolean(self) -> None:
        cv = _cell_value_from_expected({"type": "boolean", "value": True})
        assert cv.type == CellType.BOOLEAN
        assert cv.value is True

    def test_date_string(self) -> None:
        cv = _cell_value_from_expected({"type": "date", "value": "2026-01-15"})
        assert cv.type == CellType.DATE
        assert isinstance(cv.value, date)

    def test_datetime_string(self) -> None:
        cv = _cell_value_from_expected(
            {"type": "datetime", "value": "2026-01-15T10:30:00"}
        )
        assert cv.type == CellType.DATETIME
        assert isinstance(cv.value, datetime)

    def test_error(self) -> None:
        cv = _cell_value_from_expected({"type": "error", "value": "#DIV/0!"})
        assert cv.type == CellType.ERROR
        assert cv.value == "#DIV/0!"

    def test_formula(self) -> None:
        cv = _cell_value_from_expected(
            {"type": "formula", "formula": "=SUM(A1:A5)", "value": "15"}
        )
        assert cv.type == CellType.FORMULA
        assert cv.formula == "=SUM(A1:A5)"

    def test_default_string(self) -> None:
        cv = _cell_value_from_expected({"value": "fallback"})
        assert cv.type == CellType.STRING
        assert cv.value == "fallback"


# ═════════════════════════════════════════════════
# _cell_value_from_raw
# ═════════════════════════════════════════════════


class TestCellValueFromRaw:
    def test_none(self) -> None:
        assert _cell_value_from_raw(None).type == CellType.BLANK

    def test_bool(self) -> None:
        cv = _cell_value_from_raw(True)
        assert cv.type == CellType.BOOLEAN
        assert cv.value is True

    def test_int(self) -> None:
        cv = _cell_value_from_raw(42)
        assert cv.type == CellType.NUMBER
        assert cv.value == 42

    def test_float(self) -> None:
        cv = _cell_value_from_raw(3.14)
        assert cv.type == CellType.NUMBER

    def test_string(self) -> None:
        cv = _cell_value_from_raw("hello")
        assert cv.type == CellType.STRING
        assert cv.value == "hello"


# ═════════════════════════════════════════════════
# _cell_format_from_expected
# ═════════════════════════════════════════════════


class TestCellFormatFromExpected:
    def test_empty(self) -> None:
        fmt = _cell_format_from_expected({})
        assert fmt.bold is None
        assert fmt.font_name is None

    def test_all_fields(self) -> None:
        fmt = _cell_format_from_expected(
            {
                "bold": True,
                "italic": True,
                "underline": True,
                "strikethrough": True,
                "font_name": "Arial",
                "font_size": 14,
                "font_color": "#FF0000",
                "bg_color": "#00FF00",
                "number_format": "#,##0",
                "h_align": "center",
                "v_align": "middle",
                "wrap": True,
                "rotation": 45,
                "indent": 2,
            }
        )
        assert fmt.bold is True
        assert fmt.italic is True
        assert fmt.font_name == "Arial"
        assert fmt.font_size == 14
        assert fmt.h_align == "center"
        assert fmt.rotation == 45

    def test_partial(self) -> None:
        fmt = _cell_format_from_expected({"bold": True, "font_size": 12})
        assert fmt.bold is True
        assert fmt.font_size == 12
        assert fmt.italic is None


# ═════════════════════════════════════════════════
# _border_from_expected
# ═════════════════════════════════════════════════


class TestBorderFromExpected:
    def test_uniform_style(self) -> None:
        border = _border_from_expected({"border_style": "thin"})
        assert border.top is not None
        assert border.top.style.value == "thin"
        assert border.bottom is not None
        assert border.left is not None
        assert border.right is not None

    def test_default_color(self) -> None:
        border = _border_from_expected({"border_style": "thin"})
        assert border.top is not None
        assert border.top.color == "#000000"

    def test_custom_color(self) -> None:
        border = _border_from_expected(
            {"border_style": "thin", "border_color": "#FF0000"}
        )
        assert border.top is not None
        assert border.top.color == "#FF0000"

    def test_color_without_style_defaults_thin(self) -> None:
        border = _border_from_expected({"border_color": "#0000FF"})
        assert border.top is not None
        assert border.top.style.value == "thin"
        assert border.top.color == "#0000FF"

    def test_per_edge(self) -> None:
        border = _border_from_expected(
            {"border_top": "thick", "border_bottom": "thin"}
        )
        assert border.top is not None
        assert border.top.style.value == "thick"
        assert border.bottom is not None
        assert border.bottom.style.value == "thin"
        assert border.left is None
        assert border.right is None

    def test_diagonal_up(self) -> None:
        border = _border_from_expected({"border_diagonal_up": "thin"})
        assert border.diagonal_up is not None
        assert border.diagonal_up.style.value == "thin"

    def test_diagonal_down(self) -> None:
        border = _border_from_expected({"border_diagonal_down": "medium"})
        assert border.diagonal_down is not None
        assert border.diagonal_down.style.value == "medium"

    def test_empty(self) -> None:
        border = _border_from_expected({})
        assert border.top is None
        assert border.bottom is None
        assert border.left is None
        assert border.right is None

    def test_edge_color_without_edge_style(self) -> None:
        border = _border_from_expected({"border_top_color": "#FF0000"})
        assert border.top is not None
        assert border.top.style.value == "thin"
        assert border.top.color == "#FF0000"
