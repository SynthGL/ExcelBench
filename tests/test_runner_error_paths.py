"""Tests for runner error handling and edge-case code paths."""

from __future__ import annotations

from pathlib import Path
from typing import Any
from unittest.mock import MagicMock, patch

from excelbench.harness.runner import (
    _annotate_known_limitations,
    read_alignment_actual,
    read_border_actual,
    read_formula_actual,
    read_text_format_actual,
)
from excelbench.harness.runner import test_read as _test_read
from excelbench.harness.runner import test_read_case as _test_read_case
from excelbench.harness.runner import test_write as _test_write
from excelbench.models import (
    BorderEdge,
    BorderInfo,
    BorderStyle,
    CellFormat,
    CellType,
    CellValue,
    FeatureScore,
    OperationType,
    TestCase,
    TestFile,
)

JSONDict = dict[str, Any]


def _tc(
    tc_id: str,
    expected: JSONDict,
    *,
    sheet: str | None = None,
    cell: str | None = None,
) -> TestCase:
    return TestCase(
        id=tc_id, label=tc_id, row=2, expected=expected, sheet=sheet, cell=cell
    )


def _tf(feature: str, test_cases: list[TestCase] | None = None) -> TestFile:
    return TestFile(
        path="test.xlsx",
        feature=feature,
        tier=1,
        test_cases=test_cases or [],
    )


def _mock_adapter(*, can_read: bool = True, can_write: bool = True) -> MagicMock:
    adapter = MagicMock()
    adapter.name = "mock-adapter"
    adapter.can_read.return_value = can_read
    adapter.can_write.return_value = can_write
    adapter.output_extension = ".xlsx"
    adapter.info = MagicMock(
        name="mock-adapter",
        version="1.0",
        language="python",
        capabilities={"read", "write"},
    )
    return adapter


# ═════════════════════════════════════════════════
# _annotate_known_limitations — write-side branch
# ═════════════════════════════════════════════════


class TestAnnotateKnownLimitations:
    def test_write_side_annotation(self) -> None:
        """Write-side limitation note should be injected when write_score < 3."""
        score = FeatureScore(
            feature="alignment",
            library="pylightxl",
            write_score=1,
        )
        result = _annotate_known_limitations(score)
        assert result.notes is not None
        assert "pylightxl" in result.notes

    def test_write_side_skipped_when_score_3(self) -> None:
        """Write-side note should NOT be injected when write_score is 3 (perfect)."""
        score = FeatureScore(
            feature="alignment",
            library="pylightxl",
            write_score=3,
        )
        result = _annotate_known_limitations(score)
        assert result.notes is None

    def test_write_side_skipped_when_notes_exist(self) -> None:
        """Existing notes should not be overwritten."""
        score = FeatureScore(
            feature="alignment",
            library="pylightxl",
            write_score=1,
            notes="Custom note",
        )
        result = _annotate_known_limitations(score)
        assert result.notes == "Custom note"

    def test_no_matching_limitation(self) -> None:
        """Unknown library+feature combo returns unchanged score."""
        score = FeatureScore(
            feature="cell_values",
            library="nonexistent-lib",
            read_score=2,
        )
        result = _annotate_known_limitations(score)
        assert result.notes is None


# ═════════════════════════════════════════════════
# test_read — exception paths
# ═════════════════════════════════════════════════


class TestReadExceptionPaths:
    def test_open_workbook_failure(self, tmp_path: Path) -> None:
        """When open_workbook raises, all TCs should get error results."""
        adapter = _mock_adapter()
        adapter.open_workbook.side_effect = OSError("File corrupt")

        tf = _tf("cell_values", [_tc("t1", {"type": "string", "value": "x"})])
        results = _test_read(adapter, tf, tmp_path / "test.xlsx")

        assert len(results) == 1
        assert results[0].passed is False
        assert "Failed to open workbook" in (results[0].notes or "")
        assert results[0].diagnostics

    def test_get_sheet_names_empty(self, tmp_path: Path) -> None:
        """When get_sheet_names returns empty list, should raise ValueError."""
        adapter = _mock_adapter()
        adapter.open_workbook.return_value = MagicMock()
        adapter.get_sheet_names.return_value = []

        tf = _tf("cell_values", [_tc("t1", {"type": "string", "value": "x"})])
        results = _test_read(adapter, tf, tmp_path / "test.xlsx")

        assert len(results) == 1
        assert results[0].passed is False
        actual = results[0].actual
        error_msg = actual.get("error", "") if isinstance(actual, dict) else ""
        assert "No sheets found" in error_msg

    def test_exception_during_read(self, tmp_path: Path) -> None:
        """When get_sheet_names raises, inner exception handler catches it."""
        adapter = _mock_adapter()
        adapter.open_workbook.return_value = MagicMock()
        adapter.get_sheet_names.side_effect = RuntimeError("Corrupt index")

        tf = _tf("cell_values", [_tc("t1", {"type": "string", "value": "x"})])
        results = _test_read(adapter, tf, tmp_path / "test.xlsx")

        assert len(results) == 1
        assert results[0].passed is False


# ═════════════════════════════════════════════════
# test_read_case — edge cases
# ═════════════════════════════════════════════════


class TestReadCaseEdgeCases:
    def test_unknown_feature(self) -> None:
        """Unknown feature should return error dict, not crash."""
        adapter = _mock_adapter()
        wb = MagicMock()

        tc = _tc("t1", {"value": "x"})
        result = _test_read_case(
            adapter, wb, "Sheet1", tc, "nonexistent_feature", OperationType.READ
        )
        assert result.passed is False
        actual = result.actual
        error_msg = actual.get("error", "") if isinstance(actual, dict) else ""
        assert "Unknown feature" in error_msg

    def test_exception_in_read_case(self) -> None:
        """Exception during read_cell_value_actual should be caught."""
        adapter = _mock_adapter()
        adapter.read_cell_value.side_effect = RuntimeError("Read error")
        wb = MagicMock()

        tc = _tc("t1", {"type": "string", "value": "x"})
        result = _test_read_case(
            adapter, wb, "Sheet1", tc, "cell_values", OperationType.READ
        )
        assert result.passed is False
        assert "Exception" in (result.notes or "")
        assert result.diagnostics


# ═════════════════════════════════════════════════
# test_write — exception paths
# ═════════════════════════════════════════════════


class TestWriteExceptionPaths:
    def test_write_create_workbook_failure(self, tmp_path: Path) -> None:
        """When create_workbook raises, all TCs get error results."""
        adapter = _mock_adapter()
        adapter.create_workbook.side_effect = RuntimeError("Cannot create")

        tf = _tf("cell_values", [_tc("t1", {"type": "string", "value": "x"})])
        results = _test_write(adapter, tf, tmp_path / "test.xlsx")

        assert len(results) == 1
        assert results[0].passed is False
        assert results[0].operation == OperationType.WRITE

    def test_write_save_failure(self, tmp_path: Path) -> None:
        """When save_workbook raises, all TCs get error results."""
        adapter = _mock_adapter()
        adapter.create_workbook.return_value = MagicMock()
        adapter.save_workbook.side_effect = OSError("Disk full")

        tf = _tf("cell_values", [_tc("t1", {"type": "string", "value": "x"})])

        with patch(
            "excelbench.harness.runner.get_write_verifier_for_adapter"
        ) as mock_verifier:
            mock_verifier.return_value = _mock_adapter()
            results = _test_write(adapter, tf, tmp_path / "test.xlsx")

        assert len(results) == 1
        assert results[0].passed is False
        assert "Write failed" in (results[0].notes or "")
        assert results[0].diagnostics

    def test_write_verification_open_failure(self, tmp_path: Path) -> None:
        """When verifier.open_workbook raises after write, TCs get error results."""
        adapter = _mock_adapter()
        adapter.create_workbook.return_value = MagicMock()
        adapter.save_workbook.return_value = None

        verifier = _mock_adapter()
        verifier.open_workbook.side_effect = OSError("Cannot verify")

        tf = _tf("cell_values", [_tc("t1", {"type": "string", "value": "x"})])

        with patch(
            "excelbench.harness.runner.get_write_verifier_for_adapter",
            return_value=verifier,
        ):
            results = _test_write(adapter, tf, tmp_path / "test.xlsx")

        assert len(results) == 1
        assert results[0].passed is False
        assert "Failed to open workbook" in (results[0].notes or "")
        assert results[0].diagnostics


# ═════════════════════════════════════════════════
# read_*_actual functions — branch coverage
# ═════════════════════════════════════════════════


class TestReadFormulaActual:
    def test_non_formula_cell(self) -> None:
        """When cell is not a formula, should return error dict."""
        adapter = _mock_adapter()
        adapter.read_cell_value.return_value = CellValue(
            type=CellType.STRING, value="hello"
        )
        result = read_formula_actual(adapter, MagicMock(), "Sheet1", "B2")
        assert "error" in result
        assert "formula" in result["error"]


class TestReadTextFormatActual:
    def test_underline_and_strikethrough(self) -> None:
        """Underline and strikethrough branches should be included."""
        adapter = _mock_adapter()
        adapter.read_cell_format.return_value = CellFormat(
            underline="single", strikethrough=True
        )
        result = read_text_format_actual(adapter, MagicMock(), "Sheet1", "B2")
        assert result["underline"] == "single"
        assert result["strikethrough"] is True


class TestReadAlignmentActual:
    def test_wrap_rotation_indent(self) -> None:
        """Wrap, rotation, and indent branches should all appear."""
        adapter = _mock_adapter()
        adapter.read_cell_format.return_value = CellFormat(
            h_align="center", v_align="top", wrap=True, rotation=45, indent=2
        )
        result = read_alignment_actual(adapter, MagicMock(), "Sheet1", "B2")
        assert result["h_align"] == "center"
        assert result["v_align"] == "top"
        assert result["wrap"] is True
        assert result["rotation"] == 45
        assert result["indent"] == 2


class TestReadBorderActualDiagonal:
    def test_diagonal_borders(self) -> None:
        """Diagonal up/down border branches should be included."""
        adapter = _mock_adapter()
        adapter.read_cell_border.return_value = BorderInfo(
            diagonal_up=BorderEdge(style=BorderStyle.THIN, color="#000000"),
            diagonal_down=BorderEdge(style=BorderStyle.MEDIUM, color="#FF0000"),
        )
        result = read_border_actual(adapter, MagicMock(), "Sheet1", "B2")
        assert result["border_diagonal_up"] == "thin"
        assert result["border_diagonal_down"] == "medium"


def test_failed_assertion_has_data_mismatch_diagnostic() -> None:
    adapter = _mock_adapter()
    adapter.read_cell_value.return_value = CellValue(type=CellType.STRING, value="y")
    wb = MagicMock()
    tc = _tc("t1", {"type": "string", "value": "x"})
    result = _test_read_case(adapter, wb, "Sheet1", tc, "cell_values", OperationType.READ)
    assert result.passed is False
    assert result.diagnostics
    adapter.build_mismatch_diagnostic.assert_called_once()
