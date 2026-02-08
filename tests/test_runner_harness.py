"""Tests for runner harness functions: read_*_actual, _write_*_case, scoring, verifier selection."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any
from unittest.mock import MagicMock, patch

import pytest

from excelbench.harness.adapters.openpyxl_adapter import OpenpyxlAdapter
from excelbench.harness.runner import (
    _collect_sheet_names,
    _read_cell_scalar,
    _strip_cf_priority,
    _write_alignment_case,
    _write_background_color_case,
    _write_border_case,
    _write_cell_value_case,
    _write_comment_case,
    _write_conditional_format_case,
    _write_data_validation_case,
    _write_dimensions_case,
    _write_formula_case,
    _write_freeze_panes_case,
    _write_hyperlink_case,
    _write_merged_cells_case,
    _write_multi_sheet_case,
    _write_number_format_case,
    _write_text_format_case,
    calculate_score,
    get_write_verifier,
    get_write_verifier_for_adapter,
    get_write_verifier_for_feature,
    read_background_color_actual,
    read_border_actual,
    read_cell_value_actual,
    read_comment_actual,
    read_conditional_format_actual,
    read_data_validation_actual,
    read_dimensions_actual,
    read_formula_actual,
    read_freeze_panes_actual,
    read_hyperlink_actual,
    read_image_actual,
    read_merged_cells_actual,
    read_number_format_actual,
    read_sheet_names_actual,
    read_text_format_actual,
)
from excelbench.harness.runner import (
    test_read_case as _test_read_case,
)
from excelbench.models import (
    CellType,
    Importance,
    OperationType,
    TestCase,
    TestFile,
    TestResult,
)

JSONDict = dict[str, Any]

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "excel"

pytestmark = pytest.mark.skipif(
    not (FIXTURES_DIR / "manifest.json").exists(),
    reason="Canonical fixtures not found",
)


@pytest.fixture
def adapter() -> OpenpyxlAdapter:
    return OpenpyxlAdapter()


# ═════════════════════════════════════════════════
# read_cell_value_actual
# ═════════════════════════════════════════════════


class TestReadCellValueActual:
    def test_string(self, adapter: OpenpyxlAdapter) -> None:
        wb = adapter.open_workbook(FIXTURES_DIR / "tier1/01_cell_values.xlsx")
        result = read_cell_value_actual(adapter, wb, "cell_values", "B2", {})
        assert result["type"] == "string"
        assert result["value"] == "Hello World"
        adapter.close_workbook(wb)

    def test_number(self, adapter: OpenpyxlAdapter) -> None:
        wb = adapter.open_workbook(FIXTURES_DIR / "tier1/01_cell_values.xlsx")
        result = read_cell_value_actual(adapter, wb, "cell_values", "B7", {})
        assert result["type"] == "number"
        assert result["value"] == 42
        adapter.close_workbook(wb)

    def test_blank(self, adapter: OpenpyxlAdapter) -> None:
        wb = adapter.open_workbook(FIXTURES_DIR / "tier1/01_cell_values.xlsx")
        result = read_cell_value_actual(adapter, wb, "cell_values", "B4", {})
        assert result["type"] == "blank"
        assert "value" not in result
        adapter.close_workbook(wb)

    def test_date(self, adapter: OpenpyxlAdapter) -> None:
        wb = adapter.open_workbook(FIXTURES_DIR / "tier1/01_cell_values.xlsx")
        result = read_cell_value_actual(adapter, wb, "cell_values", "B12", {})
        assert result["type"] == "date"
        assert result["value"] == "2026-02-04"
        adapter.close_workbook(wb)

    def test_datetime(self, adapter: OpenpyxlAdapter) -> None:
        wb = adapter.open_workbook(FIXTURES_DIR / "tier1/01_cell_values.xlsx")
        result = read_cell_value_actual(adapter, wb, "cell_values", "B13", {})
        assert result["type"] == "datetime"
        assert result["value"] == "2026-02-04T10:30:45"
        adapter.close_workbook(wb)


# ═════════════════════════════════════════════════
# read_text_format_actual
# ═════════════════════════════════════════════════


class TestReadTextFormatActual:
    def test_bold(self, adapter: OpenpyxlAdapter) -> None:
        wb = adapter.open_workbook(FIXTURES_DIR / "tier1/03_text_formatting.xlsx")
        result = read_text_format_actual(adapter, wb, "text_formatting", "B2")
        assert result["bold"] is True
        adapter.close_workbook(wb)

    def test_italic(self, adapter: OpenpyxlAdapter) -> None:
        wb = adapter.open_workbook(FIXTURES_DIR / "tier1/03_text_formatting.xlsx")
        result = read_text_format_actual(adapter, wb, "text_formatting", "B3")
        assert result["italic"] is True
        adapter.close_workbook(wb)

    def test_font_color(self, adapter: OpenpyxlAdapter) -> None:
        wb = adapter.open_workbook(FIXTURES_DIR / "tier1/03_text_formatting.xlsx")
        result = read_text_format_actual(adapter, wb, "text_formatting", "B15")
        assert result["font_color"] == "#FF0000"
        adapter.close_workbook(wb)

    def test_font_name(self, adapter: OpenpyxlAdapter) -> None:
        wb = adapter.open_workbook(FIXTURES_DIR / "tier1/03_text_formatting.xlsx")
        result = read_text_format_actual(adapter, wb, "text_formatting", "B12")
        assert result["font_name"] == "Arial"
        adapter.close_workbook(wb)


# ═════════════════════════════════════════════════
# read_background_color_actual
# ═════════════════════════════════════════════════


class TestReadBackgroundColorActual:
    def test_has_bg_color(self, adapter: OpenpyxlAdapter) -> None:
        wb = adapter.open_workbook(FIXTURES_DIR / "tier1/04_background_colors.xlsx")
        result = read_background_color_actual(adapter, wb, "background_colors", "B2")
        assert "bg_color" in result
        assert result["bg_color"].startswith("#")
        adapter.close_workbook(wb)

    def test_no_bg_color(self, adapter: OpenpyxlAdapter) -> None:
        wb = adapter.open_workbook(FIXTURES_DIR / "tier1/04_background_colors.xlsx")
        # Use a cell far from data area that has no fill applied
        result = read_background_color_actual(adapter, wb, "background_colors", "Z1")
        assert result == {} or result.get("bg_color") is None
        adapter.close_workbook(wb)


# ═════════════════════════════════════════════════
# read_number_format_actual
# ═════════════════════════════════════════════════


class TestReadNumberFormatActual:
    def test_general(self, adapter: OpenpyxlAdapter) -> None:
        wb = adapter.open_workbook(FIXTURES_DIR / "tier1/01_cell_values.xlsx")
        result = read_number_format_actual(adapter, wb, "cell_values", "B7")
        # General format or empty result
        assert result.get("number_format", "General") == "General"
        adapter.close_workbook(wb)


# ═════════════════════════════════════════════════
# read_border_actual
# ═════════════════════════════════════════════════


class TestReadBorderActual:
    def test_thin_border(self, adapter: OpenpyxlAdapter) -> None:
        wb = adapter.open_workbook(FIXTURES_DIR / "tier1/07_borders.xlsx")
        result = read_border_actual(adapter, wb, "borders", "B2")
        # Should have some border info
        has_border = any(k.startswith("border_") for k in result)
        assert has_border
        adapter.close_workbook(wb)


# ═════════════════════════════════════════════════
# read_dimensions_actual
# ═════════════════════════════════════════════════


class TestReadDimensionsActual:
    def test_row_height(self, adapter: OpenpyxlAdapter) -> None:
        wb = adapter.open_workbook(FIXTURES_DIR / "tier1/08_dimensions.xlsx")
        tc = TestCase(
            id="dim_1",
            label="row height",
            row=2,
            expected={"row_height": 30.0},
            importance=Importance.BASIC,
        )
        result = read_dimensions_actual(adapter, wb, "dimensions", "B2", tc)
        assert "row_height" in result
        adapter.close_workbook(wb)

    def test_column_width(self, adapter: OpenpyxlAdapter) -> None:
        wb = adapter.open_workbook(FIXTURES_DIR / "tier1/08_dimensions.xlsx")
        tc = TestCase(
            id="dim_2",
            label="",
            row=2,
            expected={"column_width": 20.0},
            importance=Importance.BASIC,
        )
        result = read_dimensions_actual(adapter, wb, "dimensions", "B2", tc)
        assert "column_width" in result
        adapter.close_workbook(wb)


# ═════════════════════════════════════════════════
# read_formula_actual
# ═════════════════════════════════════════════════


class TestReadFormulaActual:
    def test_simple_formula(self, adapter: OpenpyxlAdapter) -> None:
        wb = adapter.open_workbook(FIXTURES_DIR / "tier1/02_formulas.xlsx")
        result = read_formula_actual(adapter, wb, "formulas", "B2")
        assert result.get("type") == "formula" or "error" in result
        adapter.close_workbook(wb)


# ═════════════════════════════════════════════════
# Tier 2 read_*_actual functions
# ═════════════════════════════════════════════════


class TestReadMergedCellsActual:
    def test_returns_merged_range(self, adapter: OpenpyxlAdapter) -> None:
        wb = adapter.open_workbook(FIXTURES_DIR / "tier2/10_merged_cells.xlsx")
        tc = TestCase(
            id="merge_1",
            label="",
            row=2,
            expected={"merged_range": "A2:C2"},
            importance=Importance.BASIC,
        )
        result = read_merged_cells_actual(adapter, wb, "merged_cells", tc)
        assert "merged_range" in result
        adapter.close_workbook(wb)


class TestReadConditionalFormatActual:
    def test_returns_cf_rule(self, adapter: OpenpyxlAdapter) -> None:
        wb = adapter.open_workbook(FIXTURES_DIR / "tier2/11_conditional_formatting.xlsx")
        expected: JSONDict = {"cf_rule": {"range": "B2:B6", "rule_type": "cellIs"}}
        result = read_conditional_format_actual(adapter, wb, "conditional_formatting", expected)
        assert "cf_rule" in result
        adapter.close_workbook(wb)


class TestReadDataValidationActual:
    def test_returns_validation(self, adapter: OpenpyxlAdapter) -> None:
        wb = adapter.open_workbook(FIXTURES_DIR / "tier2/12_data_validation.xlsx")
        # Fixture has per-cell validations; B2 has validation_type=list
        expected: JSONDict = {"validation": {"range": "B2", "validation_type": "list"}}
        result = read_data_validation_actual(adapter, wb, "data_validation", expected)
        assert "validation" in result
        adapter.close_workbook(wb)


class TestReadHyperlinkActual:
    def test_returns_hyperlink(self, adapter: OpenpyxlAdapter) -> None:
        wb = adapter.open_workbook(FIXTURES_DIR / "tier2/13_hyperlinks.xlsx")
        expected: JSONDict = {"hyperlink": {"cell": "B2", "target": "https://example.com"}}
        result = read_hyperlink_actual(adapter, wb, "hyperlinks", expected)
        assert "hyperlink" in result
        adapter.close_workbook(wb)


class TestReadImageActual:
    def test_returns_image(self, adapter: OpenpyxlAdapter) -> None:
        wb = adapter.open_workbook(FIXTURES_DIR / "tier2/14_images.xlsx")
        expected: JSONDict = {"image": {"cell": "B2", "path": "test.png"}}
        result = read_image_actual(adapter, wb, "images", expected)
        assert "image" in result
        adapter.close_workbook(wb)


class TestReadCommentActual:
    def test_returns_comment(self, adapter: OpenpyxlAdapter) -> None:
        wb = adapter.open_workbook(FIXTURES_DIR / "tier2/16_comments.xlsx")
        expected: JSONDict = {"comment": {"cell": "B2", "text": "hello"}}
        result = read_comment_actual(adapter, wb, "comments", expected)
        assert "comment" in result
        adapter.close_workbook(wb)


class TestReadFreezePanesActual:
    def test_returns_freeze(self, adapter: OpenpyxlAdapter) -> None:
        wb = adapter.open_workbook(FIXTURES_DIR / "tier2/17_freeze_panes.xlsx")
        expected: JSONDict = {"freeze": {"mode": "freeze"}}
        result = read_freeze_panes_actual(adapter, wb, "freeze_panes", expected)
        assert "freeze" in result
        adapter.close_workbook(wb)


class TestReadSheetNamesActual:
    def test_returns_sheet_names(self, adapter: OpenpyxlAdapter) -> None:
        wb = adapter.open_workbook(FIXTURES_DIR / "tier1/09_multiple_sheets.xlsx")
        result = read_sheet_names_actual(adapter, wb)
        assert "sheet_names" in result
        assert isinstance(result["sheet_names"], list)
        assert len(result["sheet_names"]) >= 2
        adapter.close_workbook(wb)


# ═════════════════════════════════════════════════
# _strip_cf_priority
# ═════════════════════════════════════════════════


class TestStripCfPriority:
    def test_removes_priority(self) -> None:
        expected: JSONDict = {"cf_rule": {"range": "B2:B6", "priority": 1, "rule_type": "cellIs"}}
        result = _strip_cf_priority(expected)
        assert "priority" not in result["cf_rule"]
        assert result["cf_rule"]["range"] == "B2:B6"

    def test_no_cf_rule(self) -> None:
        expected: JSONDict = {"value": 42}
        assert _strip_cf_priority(expected) is expected

    def test_non_dict_cf_rule(self) -> None:
        expected: JSONDict = {"cf_rule": "not_a_dict"}
        assert _strip_cf_priority(expected) is expected


# ═════════════════════════════════════════════════
# _collect_sheet_names
# ═════════════════════════════════════════════════


class TestCollectSheetNames:
    def test_explicit_sheet_names(self) -> None:
        tf = TestFile(
            path="test.xlsx",
            feature="multiple_sheets",
            tier=1,
            test_cases=[
                TestCase(
                    id="tc1",
                    label="",
                    row=1,
                    expected={"sheet_names": ["Sheet1", "Sheet2"]},
                    importance=Importance.BASIC,
                ),
            ],
        )
        result = _collect_sheet_names(tf)
        assert result == ["Sheet1", "Sheet2"]

    def test_feature_name_included(self) -> None:
        tf = TestFile(
            path="test.xlsx",
            feature="cell_values",
            tier=1,
            test_cases=[
                TestCase(
                    id="tc1",
                    label="",
                    row=2,
                    expected={"type": "string", "value": "Hi"},
                    importance=Importance.BASIC,
                ),
            ],
        )
        result = _collect_sheet_names(tf)
        assert "cell_values" in result

    def test_formula_sheet_extraction(self) -> None:
        tf = TestFile(
            path="test.xlsx",
            feature="formulas",
            tier=1,
            test_cases=[
                TestCase(
                    id="tc1",
                    label="",
                    row=2,
                    expected={"formula": "='References'!B2"},
                    importance=Importance.BASIC,
                ),
            ],
        )
        result = _collect_sheet_names(tf)
        assert "References" in result
        assert "formulas" in result

    def test_tc_sheet_appended(self) -> None:
        tf = TestFile(
            path="test.xlsx",
            feature="cell_values",
            tier=1,
            test_cases=[
                TestCase(
                    id="tc1",
                    label="",
                    row=2,
                    expected={"type": "string"},
                    importance=Importance.BASIC,
                    sheet="CustomSheet",
                ),
            ],
        )
        result = _collect_sheet_names(tf)
        assert "CustomSheet" in result

    def test_cf_formula_sheet_extraction(self) -> None:
        tf = TestFile(
            path="test.xlsx",
            feature="conditional_formatting",
            tier=1,
            test_cases=[
                TestCase(
                    id="tc1",
                    label="",
                    row=2,
                    expected={"cf_rule": {"formula": "='Data'!A1>10"}},
                    importance=Importance.BASIC,
                ),
            ],
        )
        result = _collect_sheet_names(tf)
        assert "Data" in result

    def test_dv_formula_sheet_extraction(self) -> None:
        tf = TestFile(
            path="test.xlsx",
            feature="data_validation",
            tier=1,
            test_cases=[
                TestCase(
                    id="tc1",
                    label="",
                    row=2,
                    expected={"validation": {"formula1": "='Lists'!A1:A5"}},
                    importance=Importance.BASIC,
                ),
            ],
        )
        result = _collect_sheet_names(tf)
        assert "Lists" in result


# ═════════════════════════════════════════════════
# _read_cell_scalar
# ═════════════════════════════════════════════════


class TestReadCellScalar:
    def test_string(self, adapter: OpenpyxlAdapter) -> None:
        wb = adapter.open_workbook(FIXTURES_DIR / "tier1/01_cell_values.xlsx")
        result = _read_cell_scalar(adapter, wb, "cell_values", "B2")
        assert result == "Hello World"
        adapter.close_workbook(wb)

    def test_blank(self, adapter: OpenpyxlAdapter) -> None:
        wb = adapter.open_workbook(FIXTURES_DIR / "tier1/01_cell_values.xlsx")
        result = _read_cell_scalar(adapter, wb, "cell_values", "B4")
        assert result is None
        adapter.close_workbook(wb)

    def test_date(self, adapter: OpenpyxlAdapter) -> None:
        wb = adapter.open_workbook(FIXTURES_DIR / "tier1/01_cell_values.xlsx")
        result = _read_cell_scalar(adapter, wb, "cell_values", "B12")
        assert result == "2026-02-04"
        adapter.close_workbook(wb)

    def test_datetime(self, adapter: OpenpyxlAdapter) -> None:
        wb = adapter.open_workbook(FIXTURES_DIR / "tier1/01_cell_values.xlsx")
        result = _read_cell_scalar(adapter, wb, "cell_values", "B13")
        assert result == "2026-02-04T10:30:45"
        adapter.close_workbook(wb)


# ═════════════════════════════════════════════════
# calculate_score
# ═════════════════════════════════════════════════


def _make_result(
    passed: bool, importance: Importance = Importance.BASIC
) -> TestResult:
    return TestResult(
        test_case_id="tc",
        operation=OperationType.READ,
        passed=passed,
        expected={},
        actual={},
        importance=importance,
    )


class TestCalculateScore:
    def test_empty_returns_0(self) -> None:
        assert calculate_score([]) == 0

    def test_all_basic_all_edge_pass_returns_3(self) -> None:
        results = [
            _make_result(True, Importance.BASIC),
            _make_result(True, Importance.BASIC),
            _make_result(True, Importance.EDGE),
        ]
        assert calculate_score(results) == 3

    def test_all_basic_pass_edge_fail_returns_2(self) -> None:
        results = [
            _make_result(True, Importance.BASIC),
            _make_result(False, Importance.EDGE),
        ]
        assert calculate_score(results) == 2

    def test_partial_basic_returns_1(self) -> None:
        results = [
            _make_result(True, Importance.BASIC),
            _make_result(False, Importance.BASIC),
        ]
        assert calculate_score(results) == 1

    def test_no_basic_pass_returns_0(self) -> None:
        results = [
            _make_result(False, Importance.BASIC),
            _make_result(False, Importance.BASIC),
        ]
        assert calculate_score(results) == 0

    def test_only_edge_no_basic_returns_0(self) -> None:
        results = [_make_result(True, Importance.EDGE)]
        assert calculate_score(results) == 0

    def test_all_basic_no_edge_returns_3(self) -> None:
        results = [_make_result(True, Importance.BASIC)]
        assert calculate_score(results) == 3


# ═════════════════════════════════════════════════
# get_write_verifier / get_write_verifier_for_feature
# ═════════════════════════════════════════════════


class TestGetWriteVerifier:
    def test_default_returns_openpyxl(self) -> None:
        with patch.dict(os.environ, {"EXCELBENCH_WRITE_ORACLE": "openpyxl"}):
            v = get_write_verifier()
            assert v.name == "openpyxl"

    def test_auto_on_darwin_returns_openpyxl(self) -> None:
        with patch.dict(os.environ, {"EXCELBENCH_WRITE_ORACLE": "auto"}):
            with patch("excelbench.harness.runner._excel_available", return_value=False):
                v = get_write_verifier()
                assert v.name == "openpyxl"

    def test_excel_oracle_none_fallback(self) -> None:
        with patch.dict(os.environ, {"EXCELBENCH_WRITE_ORACLE": "excel"}):
            with patch("excelbench.harness.runner.ExcelOracleAdapter", None):
                v = get_write_verifier()
                assert v.name == "openpyxl"


class TestGetWriteVerifierForFeature:
    def test_openpyxl_override(self) -> None:
        with patch.dict(os.environ, {"EXCELBENCH_WRITE_ORACLE": "openpyxl"}):
            v = get_write_verifier_for_feature("images")
            assert v.name == "openpyxl"

    def test_darwin_returns_openpyxl(self) -> None:
        with patch.dict(os.environ, {"EXCELBENCH_WRITE_ORACLE": "auto"}):
            with patch("excelbench.harness.runner.platform") as mock_platform:
                mock_platform.system.return_value = "Darwin"
                v = get_write_verifier_for_feature("conditional_formatting")
                assert v.name == "openpyxl"


class TestGetWriteVerifierForAdapter:
    def test_xls_adapter_returns_xlrd(self) -> None:
        mock_adapter = MagicMock()
        mock_adapter.output_extension = ".xls"
        v = get_write_verifier_for_adapter(mock_adapter, "cell_values")
        assert v.name == "xlrd"

    def test_xlsx_adapter_delegates(self) -> None:
        with patch.dict(os.environ, {"EXCELBENCH_WRITE_ORACLE": "openpyxl"}):
            mock_adapter = MagicMock()
            mock_adapter.output_extension = ".xlsx"
            v = get_write_verifier_for_adapter(mock_adapter, "cell_values")
            assert v.name == "openpyxl"


# ═════════════════════════════════════════════════
# test_read_case integration
# ═════════════════════════════════════════════════


class TestTestReadCase:
    def test_cell_value_pass(self, adapter: OpenpyxlAdapter) -> None:
        wb = adapter.open_workbook(FIXTURES_DIR / "tier1/01_cell_values.xlsx")
        tc = TestCase(
            id="cv_str",
            label="simple string",
            row=2,
            expected={"type": "string", "value": "Hello World"},
            importance=Importance.BASIC,
        )
        result = _test_read_case(adapter, wb, "cell_values", tc, "cell_values", OperationType.READ)
        assert result.passed is True
        assert result.test_case_id == "cv_str"
        adapter.close_workbook(wb)

    def test_cell_value_fail(self, adapter: OpenpyxlAdapter) -> None:
        wb = adapter.open_workbook(FIXTURES_DIR / "tier1/01_cell_values.xlsx")
        tc = TestCase(
            id="cv_wrong",
            label="",
            row=2,
            expected={"type": "string", "value": "Wrong Value"},
            importance=Importance.BASIC,
        )
        result = _test_read_case(adapter, wb, "cell_values", tc, "cell_values", OperationType.READ)
        assert result.passed is False
        adapter.close_workbook(wb)

    def test_multiple_sheets(self, adapter: OpenpyxlAdapter) -> None:
        wb = adapter.open_workbook(FIXTURES_DIR / "tier1/09_multiple_sheets.xlsx")
        sheet_names = adapter.get_sheet_names(wb)
        tc = TestCase(
            id="ms_names",
            label="",
            row=1,
            expected={"sheet_names": sheet_names},
            importance=Importance.BASIC,
        )
        result = _test_read_case(
            adapter, wb, "multiple_sheets", tc, "multiple_sheets", OperationType.READ
        )
        assert result.passed is True
        adapter.close_workbook(wb)


# ═════════════════════════════════════════════════
# _write_*_case functions via openpyxl roundtrip
# ═════════════════════════════════════════════════


class TestWriteCaseFunctions:
    def test_write_cell_value_case(self, adapter: OpenpyxlAdapter, tmp_path: Path) -> None:
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        _write_cell_value_case(adapter, wb, "S1", "A1", {"type": "string", "value": "Test"})
        adapter.save_workbook(wb, tmp_path / "out.xlsx")
        wb2 = adapter.open_workbook(tmp_path / "out.xlsx")
        v = adapter.read_cell_value(wb2, "S1", "A1")
        assert v.value == "Test"
        adapter.close_workbook(wb2)

    def test_write_cell_value_date(self, adapter: OpenpyxlAdapter, tmp_path: Path) -> None:
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        _write_cell_value_case(
            adapter, wb, "S1", "A1", {"type": "date", "value": "2026-01-15"}
        )
        adapter.save_workbook(wb, tmp_path / "out.xlsx")
        wb2 = adapter.open_workbook(tmp_path / "out.xlsx")
        v = adapter.read_cell_value(wb2, "S1", "A1")
        assert v.type == CellType.DATE
        adapter.close_workbook(wb2)

    def test_write_formula_case(self, adapter: OpenpyxlAdapter, tmp_path: Path) -> None:
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        _write_formula_case(adapter, wb, "S1", "A1", {"formula": "=1+1"})
        adapter.save_workbook(wb, tmp_path / "out.xlsx")
        wb2 = adapter.open_workbook(tmp_path / "out.xlsx")
        v = adapter.read_cell_value(wb2, "S1", "A1")
        assert v.type == CellType.FORMULA
        adapter.close_workbook(wb2)

    def test_write_text_format_case(self, adapter: OpenpyxlAdapter, tmp_path: Path) -> None:
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        tc = TestCase(
            id="fmt_1",
            label="Bold 14pt",
            row=1,
            expected={"bold": True, "font_size": 14},
            importance=Importance.BASIC,
        )
        _write_text_format_case(adapter, wb, "S1", "A1", tc)
        adapter.save_workbook(wb, tmp_path / "out.xlsx")
        wb2 = adapter.open_workbook(tmp_path / "out.xlsx")
        fmt = adapter.read_cell_format(wb2, "S1", "A1")
        assert fmt.bold is True
        assert fmt.font_size == 14
        adapter.close_workbook(wb2)

    def test_write_background_color_case(
        self, adapter: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        _write_background_color_case(adapter, wb, "S1", "A1", {"bg_color": "#FF0000"})
        adapter.save_workbook(wb, tmp_path / "out.xlsx")
        wb2 = adapter.open_workbook(tmp_path / "out.xlsx")
        fmt = adapter.read_cell_format(wb2, "S1", "A1")
        assert fmt.bg_color is not None
        adapter.close_workbook(wb2)

    def test_write_number_format_case(
        self, adapter: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        _write_number_format_case(adapter, wb, "S1", "A1", {"number_format": "#,##0.00"})
        adapter.save_workbook(wb, tmp_path / "out.xlsx")
        wb2 = adapter.open_workbook(tmp_path / "out.xlsx")
        fmt = adapter.read_cell_format(wb2, "S1", "A1")
        assert fmt.number_format is not None
        adapter.close_workbook(wb2)

    def test_write_number_format_date(
        self, adapter: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        _write_number_format_case(adapter, wb, "S1", "A1", {"number_format": "yyyy-mm-dd"})
        adapter.save_workbook(wb, tmp_path / "out.xlsx")
        wb2 = adapter.open_workbook(tmp_path / "out.xlsx")
        v = adapter.read_cell_value(wb2, "S1", "A1")
        assert v.type == CellType.DATE
        adapter.close_workbook(wb2)

    def test_write_alignment_case(self, adapter: OpenpyxlAdapter, tmp_path: Path) -> None:
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        _write_alignment_case(adapter, wb, "S1", "A1", {"h_align": "center"})
        adapter.save_workbook(wb, tmp_path / "out.xlsx")
        wb2 = adapter.open_workbook(tmp_path / "out.xlsx")
        fmt = adapter.read_cell_format(wb2, "S1", "A1")
        assert fmt.h_align == "center"
        adapter.close_workbook(wb2)

    def test_write_border_case(self, adapter: OpenpyxlAdapter, tmp_path: Path) -> None:
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        _write_border_case(adapter, wb, "S1", "A1", {"border_style": "thin"})
        adapter.save_workbook(wb, tmp_path / "out.xlsx")
        wb2 = adapter.open_workbook(tmp_path / "out.xlsx")
        border = adapter.read_cell_border(wb2, "S1", "A1")
        assert border.top is not None
        adapter.close_workbook(wb2)

    def test_write_dimensions_row_height(
        self, adapter: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        tc = TestCase(
            id="dim_rh",
            label="",
            row=3,
            expected={"row_height": 40.0},
            importance=Importance.BASIC,
        )
        _write_dimensions_case(adapter, wb, "S1", "B3", tc)
        adapter.save_workbook(wb, tmp_path / "out.xlsx")
        wb2 = adapter.open_workbook(tmp_path / "out.xlsx")
        h = adapter.read_row_height(wb2, "S1", 3)
        assert h is not None and abs(h - 40.0) < 1.0
        adapter.close_workbook(wb2)

    def test_write_multi_sheet_skip_sheet_names(self, adapter: OpenpyxlAdapter) -> None:
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        # Should be a no-op for sheet_names expected
        _write_multi_sheet_case(adapter, wb, "S1", "A1", {"sheet_names": ["S1"]})

    def test_write_multi_sheet_cell_value(
        self, adapter: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        _write_multi_sheet_case(
            adapter, wb, "S1", "A1", {"type": "string", "value": "Multi"}
        )
        adapter.save_workbook(wb, tmp_path / "out.xlsx")
        wb2 = adapter.open_workbook(tmp_path / "out.xlsx")
        v = adapter.read_cell_value(wb2, "S1", "A1")
        assert v.value == "Multi"
        adapter.close_workbook(wb2)

    def test_write_merged_cells_case(
        self, adapter: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        _write_merged_cells_case(
            adapter, wb, "S1", {"merged_range": "A1:C1", "top_left_value": "Hello"}
        )
        adapter.save_workbook(wb, tmp_path / "out.xlsx")
        wb2 = adapter.open_workbook(tmp_path / "out.xlsx")
        ranges = adapter.read_merged_ranges(wb2, "S1")
        assert len(ranges) >= 1
        adapter.close_workbook(wb2)

    def test_write_merged_cells_with_bg_color(
        self, adapter: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        _write_merged_cells_case(
            adapter,
            wb,
            "S1",
            {"merged_range": "A1:B2", "top_left_bg_color": "#00FF00"},
        )
        adapter.save_workbook(wb, tmp_path / "out.xlsx")
        wb2 = adapter.open_workbook(tmp_path / "out.xlsx")
        fmt = adapter.read_cell_format(wb2, "S1", "A1")
        assert fmt.bg_color is not None
        adapter.close_workbook(wb2)

    def test_write_conditional_format_case(
        self, adapter: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        _write_conditional_format_case(
            adapter,
            wb,
            "S1",
            {
                "cf_rule": {
                    "range": "A1:A5",
                    "rule_type": "cellIs",
                    "operator": "greaterThan",
                    "formula": "10",
                }
            },
        )
        adapter.save_workbook(wb, tmp_path / "out.xlsx")
        # Just verifying it doesn't crash

    def test_write_data_validation_case(
        self, adapter: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        _write_data_validation_case(
            adapter,
            wb,
            "S1",
            {"validation": {"range": "A1:A5", "validation_type": "list", "formula1": '"A,B,C"'}},
        )
        adapter.save_workbook(wb, tmp_path / "out.xlsx")

    def test_write_hyperlink_case(
        self, adapter: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        _write_hyperlink_case(
            adapter, wb, "S1", {"hyperlink": {"cell": "A1", "target": "https://example.com"}}
        )
        adapter.save_workbook(wb, tmp_path / "out.xlsx")
        wb2 = adapter.open_workbook(tmp_path / "out.xlsx")
        links = adapter.read_hyperlinks(wb2, "S1")
        assert len(links) >= 1
        adapter.close_workbook(wb2)

    def test_write_comment_case(self, adapter: OpenpyxlAdapter, tmp_path: Path) -> None:
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        _write_comment_case(
            adapter, wb, "S1", {"comment": {"cell": "A1", "text": "Note", "author": "Bot"}}
        )
        adapter.save_workbook(wb, tmp_path / "out.xlsx")
        wb2 = adapter.open_workbook(tmp_path / "out.xlsx")
        comments = adapter.read_comments(wb2, "S1")
        assert len(comments) >= 1
        adapter.close_workbook(wb2)

    def test_write_freeze_panes_case(
        self, adapter: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        _write_freeze_panes_case(
            adapter, wb, "S1", {"freeze": {"mode": "freeze", "top_left_cell": "B2"}}
        )
        adapter.save_workbook(wb, tmp_path / "out.xlsx")
        wb2 = adapter.open_workbook(tmp_path / "out.xlsx")
        settings = adapter.read_freeze_panes(wb2, "S1")
        assert settings.get("mode") == "freeze" or settings.get("top_left_cell") is not None
        adapter.close_workbook(wb2)
