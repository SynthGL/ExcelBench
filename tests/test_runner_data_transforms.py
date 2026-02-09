"""Tests for runner data transformation and normalization functions."""

from __future__ import annotations

from datetime import date, datetime
from typing import Any

from excelbench.harness.runner import (
    _border_from_expected,
    _cell_format_from_expected,
    _cell_value_from_expected,
    _cell_value_from_raw,
    _collect_sheet_names,
    _extract_formula_sheet_names,
    _find_rule,
    _find_validation,
    _normalize_formula,
    _normalize_number_format,
    _normalize_sheet_quotes,
    _project_rule,
    _strip_cf_priority,
    get_write_verifier,
    get_write_verifier_for_adapter,
    get_write_verifier_for_feature,
)
from excelbench.models import CellType
from excelbench.models import TestCase as BenchCase
from excelbench.models import TestFile as BenchFile

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
        cv = _cell_value_from_expected({"type": "datetime", "value": "2026-01-15T10:30:00"})
        assert cv.type == CellType.DATETIME
        assert isinstance(cv.value, datetime)

    def test_error(self) -> None:
        cv = _cell_value_from_expected({"type": "error", "value": "#DIV/0!"})
        assert cv.type == CellType.ERROR
        assert cv.value == "#DIV/0!"

    def test_formula(self) -> None:
        cv = _cell_value_from_expected({"type": "formula", "formula": "=SUM(A1:A5)", "value": "15"})
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
        border = _border_from_expected({"border_style": "thin", "border_color": "#FF0000"})
        assert border.top is not None
        assert border.top.color == "#FF0000"

    def test_color_without_style_defaults_thin(self) -> None:
        border = _border_from_expected({"border_color": "#0000FF"})
        assert border.top is not None
        assert border.top.style.value == "thin"
        assert border.top.color == "#0000FF"

    def test_per_edge(self) -> None:
        border = _border_from_expected({"border_top": "thick", "border_bottom": "thin"})
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


# ═════════════════════════════════════════════════
# _strip_cf_priority
# ═════════════════════════════════════════════════


class TestStripCfPriority:
    def test_no_cf_rule(self) -> None:
        d: JSONDict = {"range": "A1:A5"}
        assert _strip_cf_priority(d) is d  # returns same object

    def test_cf_rule_not_dict(self) -> None:
        d: JSONDict = {"cf_rule": "not_a_dict"}
        assert _strip_cf_priority(d) is d

    def test_strips_priority(self) -> None:
        d: JSONDict = {"cf_rule": {"range": "A1:A5", "priority": 1, "rule_type": "cellIs"}}
        result = _strip_cf_priority(d)
        assert "priority" not in result["cf_rule"]
        assert result["cf_rule"]["range"] == "A1:A5"
        assert result["cf_rule"]["rule_type"] == "cellIs"

    def test_preserves_original(self) -> None:
        d: JSONDict = {"cf_rule": {"priority": 1, "x": 2}}
        _strip_cf_priority(d)
        assert "priority" in d["cf_rule"]  # original not mutated

    def test_no_priority_in_rule(self) -> None:
        d: JSONDict = {"cf_rule": {"rule_type": "cellIs"}}
        result = _strip_cf_priority(d)
        assert result["cf_rule"] == {"rule_type": "cellIs"}


# ═════════════════════════════════════════════════
# _find_validation
# ═════════════════════════════════════════════════


class TestFindValidation:
    def test_match_by_range(self) -> None:
        validations: list[JSONDict] = [
            {"range": "A1:A5", "validation_type": "list"},
            {"range": "B1:B5", "validation_type": "whole"},
        ]
        result = _find_validation(validations, {"range": "B1:B5"})
        assert result is not None
        assert result["validation_type"] == "whole"

    def test_match_by_validation_type(self) -> None:
        validations: list[JSONDict] = [
            {"range": "A1:A5", "validation_type": "list"},
            {"range": "A1:A5", "validation_type": "whole"},
        ]
        result = _find_validation(validations, {"range": "A1:A5", "validation_type": "whole"})
        assert result is not None
        assert result["validation_type"] == "whole"

    def test_match_by_formula1(self) -> None:
        validations: list[JSONDict] = [
            {"range": "A1:A5", "formula1": "=10"},
            {"range": "A1:A5", "formula1": "=20"},
        ]
        result = _find_validation(validations, {"range": "A1:A5", "formula1": "=10"})
        assert result is not None
        assert result["formula1"] == "=10"

    def test_no_match(self) -> None:
        validations: list[JSONDict] = [{"range": "A1:A5"}]
        assert _find_validation(validations, {"range": "Z1:Z5"}) is None

    def test_empty(self) -> None:
        assert _find_validation([], {"range": "A1:A5"}) is None

    def test_range_normalization(self) -> None:
        validations: list[JSONDict] = [{"range": "$A$1:$A$5"}]
        result = _find_validation(validations, {"range": "A1:A5"})
        assert result is not None


# ═════════════════════════════════════════════════
# _collect_sheet_names
# ═════════════════════════════════════════════════


def _tc(
    tc_id: str,
    expected: JSONDict,
    *,
    sheet: str | None = None,
) -> BenchCase:
    """Shortcut to build a TestCase for _collect_sheet_names tests."""
    return BenchCase(id=tc_id, label=tc_id, row=1, expected=expected, sheet=sheet)


class TestCollectSheetNames:
    def test_empty_test_cases(self) -> None:
        tf = BenchFile(path="a.xlsx", feature="cell_values", tier=1, test_cases=[])
        result = _collect_sheet_names(tf)
        assert result == ["cell_values"]

    def test_explicit_sheet_names(self) -> None:
        tf = BenchFile(
            path="a.xlsx",
            feature="multiple_sheets",
            tier=1,
            test_cases=[_tc("t1", {"sheet_names": ["Sheet1", "Sheet2"]})],
        )
        result = _collect_sheet_names(tf)
        assert result == ["Sheet1", "Sheet2"]

    def test_formula_sheet_extraction(self) -> None:
        tf = BenchFile(
            path="a.xlsx",
            feature="formulas",
            tier=1,
            test_cases=[_tc("t1", {"formula": "='Data'!A1"})],
        )
        result = _collect_sheet_names(tf)
        assert "formulas" in result
        assert "Data" in result

    def test_cf_formula_sheet_extraction(self) -> None:
        tf = BenchFile(
            path="a.xlsx",
            feature="conditional_formatting",
            tier=1,
            test_cases=[_tc("t1", {"cf_rule": {"formula": "='Ref'!A1>0"}})],
        )
        result = _collect_sheet_names(tf)
        assert "Ref" in result

    def test_dv_formula_sheet_extraction(self) -> None:
        tf = BenchFile(
            path="a.xlsx",
            feature="data_validation",
            tier=1,
            test_cases=[_tc("t1", {"validation": {"formula1": "='Lists'!A1:A5"}})],
        )
        result = _collect_sheet_names(tf)
        assert "Lists" in result

    def test_feature_prepended(self) -> None:
        tf = BenchFile(
            path="a.xlsx",
            feature="borders",
            tier=1,
            test_cases=[_tc("t1", {"border_style": "thin"})],
        )
        result = _collect_sheet_names(tf)
        assert result[0] == "borders"

    def test_explicit_sheet_on_tc(self) -> None:
        tf = BenchFile(
            path="a.xlsx",
            feature="cell_values",
            tier=1,
            test_cases=[_tc("t1", {"type": "string"}, sheet="Custom")],
        )
        result = _collect_sheet_names(tf)
        assert "Custom" in result
        assert "cell_values" in result


# ═════════════════════════════════════════════════
# get_write_verifier / get_write_verifier_for_feature
# ═════════════════════════════════════════════════


class TestGetWriteVerifier:
    def test_openpyxl_env(self, monkeypatch: Any) -> None:
        monkeypatch.setenv("EXCELBENCH_WRITE_ORACLE", "openpyxl")
        v = get_write_verifier()
        assert v.name == "openpyxl"

    def test_default_returns_adapter(self, monkeypatch: Any) -> None:
        monkeypatch.setenv("EXCELBENCH_WRITE_ORACLE", "auto")
        v = get_write_verifier()
        assert v.name  # returns some adapter

    def test_excel_env(self, monkeypatch: Any) -> None:
        monkeypatch.setenv("EXCELBENCH_WRITE_ORACLE", "excel")
        v = get_write_verifier()
        # Returns excel_oracle if xlwings available, else openpyxl
        assert v.name in {"openpyxl", "excel_oracle"}


class TestGetWriteVerifierForFeature:
    def test_openpyxl_env_override(self, monkeypatch: Any) -> None:
        monkeypatch.setenv("EXCELBENCH_WRITE_ORACLE", "openpyxl")
        v = get_write_verifier_for_feature("images")
        assert v.name == "openpyxl"

    def test_simple_feature_on_darwin(self, monkeypatch: Any) -> None:
        monkeypatch.setenv("EXCELBENCH_WRITE_ORACLE", "auto")
        v = get_write_verifier_for_feature("cell_values")
        assert v.name  # returns valid adapter


class TestGetWriteVerifierForAdapter:
    def test_xls_adapter_uses_xlrd(self) -> None:
        from unittest.mock import MagicMock

        adapter = MagicMock()
        adapter.output_extension = ".xls"
        v = get_write_verifier_for_adapter(adapter, "cell_values")
        assert v.name == "xlrd"

    def test_xlsx_adapter_uses_feature_verifier(self, monkeypatch: Any) -> None:
        from unittest.mock import MagicMock

        monkeypatch.setenv("EXCELBENCH_WRITE_ORACLE", "openpyxl")
        adapter = MagicMock()
        adapter.output_extension = ".xlsx"
        v = get_write_verifier_for_adapter(adapter, "cell_values")
        assert v.name == "openpyxl"
