"""Tests for runner.py feature-specific read functions and run_benchmark edge cases.

Targets uncovered branches:
- read_merged_cells_actual: top_left_value, non_top_left_nonempty, bg_color
- read_conditional_format_actual: formula normalization
- read_data_validation_actual: formula normalization
- read_hyperlink_actual: internal target normalization
- read_comment_actual, read_image_actual, read_pivot_actual: not-found paths
- read_freeze_panes_actual
- run_benchmark: missing manifest, no matching features, file not found
- _deep_compare: color string, tuple comparison
- test_read_case: feature dispatch branches (images, hyperlinks, etc.)
"""

from __future__ import annotations

from pathlib import Path
from typing import Any
from unittest.mock import MagicMock

import pytest

from excelbench.harness.adapters.openpyxl_adapter import OpenpyxlAdapter
from excelbench.harness.runner import (
    compare_results,
    read_comment_actual,
    read_conditional_format_actual,
    read_data_validation_actual,
    read_freeze_panes_actual,
    read_hyperlink_actual,
    read_image_actual,
    read_merged_cells_actual,
    read_pivot_actual,
    run_benchmark,
)
from excelbench.harness.runner import test_read_case as _test_read_case
from excelbench.models import (
    CellFormat,
    CellType,
    CellValue,
    OperationType,
    TestCase,
)

JSONDict = dict[str, Any]


def _tc(
    tc_id: str,
    expected: JSONDict,
    *,
    sheet: str | None = None,
    cell: str | None = None,
    row: int = 2,
) -> TestCase:
    return TestCase(id=tc_id, label=tc_id, row=row, expected=expected, sheet=sheet, cell=cell)


@pytest.fixture
def opxl() -> OpenpyxlAdapter:
    return OpenpyxlAdapter()


# ═════════════════════════════════════════════════
# read_merged_cells_actual — branch coverage
# ═════════════════════════════════════════════════


class TestReadMergedCellsActual:
    def test_top_left_value(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        """top_left_value branch: reads the value of the top-left cell in a merge."""
        path = tmp_path / "merge.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="Merged"))
        opxl.merge_cells(wb, "S1", "A1:C1")
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        tc = _tc("t1", {"merged_range": "A1:C1", "top_left_value": "Merged"}, row=1)
        result = read_merged_cells_actual(opxl, wb2, "S1", tc)
        assert result["merged_range"] == "A1:C1"
        assert result["top_left_value"] == "Merged"
        opxl.close_workbook(wb2)

    def test_non_top_left_nonempty(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        """non_top_left_nonempty branch: counts non-blank cells in merge (excluding top-left)."""
        path = tmp_path / "merge_nonempty.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="Top"))
        opxl.merge_cells(wb, "S1", "A1:B2")
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        tc = _tc("t1", {"merged_range": "A1:B2", "non_top_left_nonempty": 0}, row=1)
        result = read_merged_cells_actual(opxl, wb2, "S1", tc)
        assert result.get("non_top_left_nonempty") == 0
        opxl.close_workbook(wb2)

    def test_top_left_bg_color(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        """top_left_bg_color branch: reads background color of merge top-left cell."""
        path = tmp_path / "merge_bg.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="x"))
        opxl.write_cell_format(wb, "S1", "A1", CellFormat(bg_color="#FF0000"))
        opxl.merge_cells(wb, "S1", "A1:B1")
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        tc = _tc("t1", {"merged_range": "A1:B1", "top_left_bg_color": "#FFFF0000"}, row=1)
        result = read_merged_cells_actual(opxl, wb2, "S1", tc)
        assert result.get("top_left_bg_color") is not None
        opxl.close_workbook(wb2)

    def test_non_top_left_bg_color(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        """non_top_left_bg_color branch: reads bg color of a non-top-left cell."""
        path = tmp_path / "merge_bg2.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="x"))
        opxl.merge_cells(wb, "S1", "A1:B2")
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        # expected value must be non-None to enter the branch
        tc = _tc("t1", {"merged_range": "A1:B2", "non_top_left_bg_color": "#000000"}, row=1)
        result = read_merged_cells_actual(opxl, wb2, "S1", tc)
        assert "non_top_left_bg_color" in result
        opxl.close_workbook(wb2)

    def test_range_not_found(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        """When expected range is not in merged_ranges, result has no match."""
        path = tmp_path / "no_merge.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="x"))
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        tc = _tc("t1", {"merged_range": "A1:B2"}, row=1)
        result = read_merged_cells_actual(opxl, wb2, "S1", tc)
        assert result.get("merged_range") is None
        opxl.close_workbook(wb2)


# ═════════════════════════════════════════════════
# read_conditional_format_actual
# ═════════════════════════════════════════════════


class TestReadConditionalFormatActual:
    def test_formula_normalization(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        """When formula matches after normalization, result uses expected formula."""
        path = tmp_path / "cf.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.NUMBER, value=5))
        opxl.add_conditional_format(
            wb,
            "S1",
            {
                "cf_rule": {
                    "range": "A1:A10",
                    "rule_type": "cellIs",
                    "operator": "greaterThan",
                    "formula": "5",
                    "format": {"bg_color": "#FF0000"},
                }
            },
        )
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        expected: JSONDict = {
            "cf_rule": {
                "rule_type": "cellIs",
                "operator": "greaterThan",
                "formula": "5",
            }
        }
        result = read_conditional_format_actual(opxl, wb2, "S1", expected)
        assert "cf_rule" in result
        opxl.close_workbook(wb2)

    def test_no_matching_rule(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        """No matching CF rule → empty dict."""
        path = tmp_path / "cf_empty.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        expected: JSONDict = {
            "cf_rule": {"rule_type": "cellIs", "operator": "equal", "formula": "999"}
        }
        result = read_conditional_format_actual(opxl, wb2, "S1", expected)
        assert result == {}
        opxl.close_workbook(wb2)


# ═════════════════════════════════════════════════
# read_data_validation_actual
# ═════════════════════════════════════════════════


class TestReadDataValidationActual:
    def test_validation_with_formula_match(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        """DV roundtrip with formula normalization."""
        path = tmp_path / "dv.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.add_data_validation(
            wb,
            "S1",
            {
                "validation": {
                    "range": "A1:A10",
                    "validation_type": "whole",
                    "operator": "between",
                    "formula1": "1",
                    "formula2": "100",
                }
            },
        )
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        expected: JSONDict = {
            "validation": {
                "validation_type": "whole",
                "operator": "between",
                "formula1": "1",
                "formula2": "100",
            }
        }
        result = read_data_validation_actual(opxl, wb2, "S1", expected)
        assert "validation" in result
        opxl.close_workbook(wb2)

    def test_no_matching_validation(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        """No matching DV → empty dict."""
        path = tmp_path / "dv_empty.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        expected: JSONDict = {"validation": {"validation_type": "list"}}
        result = read_data_validation_actual(opxl, wb2, "S1", expected)
        assert result == {}
        opxl.close_workbook(wb2)


# ═════════════════════════════════════════════════
# read_hyperlink_actual
# ═════════════════════════════════════════════════


class TestReadHyperlinkActual:
    def test_external_hyperlink(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "link.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.add_hyperlink(
            wb,
            "S1",
            {"hyperlink": {"cell": "A1", "target": "https://example.com", "display": "Click"}},
        )
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        expected: JSONDict = {"hyperlink": {"cell": "A1", "target": "https://example.com"}}
        result = read_hyperlink_actual(opxl, wb2, "S1", expected)
        assert "hyperlink" in result
        opxl.close_workbook(wb2)

    def test_internal_hyperlink(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        """Internal hyperlink → target normalization (strip #, remove quotes)."""
        path = tmp_path / "link_int.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.add_hyperlink(
            wb,
            "S1",
            {
                "hyperlink": {
                    "cell": "A1",
                    "target": "#Sheet2!A1",
                    "display": "Go to Sheet2",
                    "internal": True,
                }
            },
        )
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        expected: JSONDict = {"hyperlink": {"cell": "A1", "target": "#Sheet2!A1", "internal": True}}
        result = read_hyperlink_actual(opxl, wb2, "S1", expected)
        assert "hyperlink" in result
        opxl.close_workbook(wb2)

    def test_no_matching_hyperlink(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        """No matching hyperlink → empty dict."""
        path = tmp_path / "link_empty.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        expected: JSONDict = {"hyperlink": {"cell": "A1"}}
        result = read_hyperlink_actual(opxl, wb2, "S1", expected)
        assert result == {}
        opxl.close_workbook(wb2)


# ═════════════════════════════════════════════════
# read_comment_actual
# ═════════════════════════════════════════════════


class TestReadCommentActual:
    def test_comment_roundtrip(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "comment.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.add_comment(wb, "S1", {"comment": {"cell": "A1", "text": "Hello", "author": "Test"}})
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        expected: JSONDict = {"comment": {"cell": "A1", "text": "Hello"}}
        result = read_comment_actual(opxl, wb2, "S1", expected)
        assert "comment" in result
        assert result["comment"]["text"] == "Hello"
        opxl.close_workbook(wb2)

    def test_no_matching_comment(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        """No matching comment → empty dict."""
        path = tmp_path / "comment_empty.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        expected: JSONDict = {"comment": {"cell": "A1"}}
        result = read_comment_actual(opxl, wb2, "S1", expected)
        assert result == {}
        opxl.close_workbook(wb2)


# ═════════════════════════════════════════════════
# read_image_actual
# ═════════════════════════════════════════════════


class TestReadImageActual:
    def test_no_matching_image(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        """No matching image → empty dict."""
        path = tmp_path / "img_empty.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        expected: JSONDict = {"image": {"cell": "A1"}}
        result = read_image_actual(opxl, wb2, "S1", expected)
        assert result == {}
        opxl.close_workbook(wb2)


# ═════════════════════════════════════════════════
# read_pivot_actual
# ═════════════════════════════════════════════════


class TestReadPivotActual:
    def test_no_matching_pivot(self) -> None:
        """No matching pivot → empty dict."""
        adapter = MagicMock()
        adapter.read_pivot_tables.return_value = []
        expected: JSONDict = {"pivot": {"name": "PivotTable1"}}
        result = read_pivot_actual(adapter, MagicMock(), "S1", expected)
        assert result == {}

    def test_pivot_match_by_target_cell(self) -> None:
        """Match pivot by target_cell when name doesn't match."""
        adapter = MagicMock()
        adapter.read_pivot_tables.return_value = [
            {"name": "Other", "target_cell": "Sheet1!A1", "source_range": "Data!A1:D10"}
        ]
        expected: JSONDict = {"pivot": {"name": "Missing", "target_cell": "Sheet1!A1"}}
        result = read_pivot_actual(adapter, MagicMock(), "S1", expected)
        assert "pivot" in result

    def test_pivot_target_cell_normalization(self) -> None:
        """Pivot target_cell with $ and ':' gets normalized."""
        adapter = MagicMock()
        adapter.read_pivot_tables.return_value = [
            {"name": "PT1", "target_cell": "$A$1:$B$2", "source_range": "A1:D10"}
        ]
        expected: JSONDict = {"pivot": {"name": "PT1", "target_cell": "Sheet1!A1"}}
        result = read_pivot_actual(adapter, MagicMock(), "S1", expected)
        assert "pivot" in result
        # target_cell should be cleaned: $ stripped, ':'split, '!' prefixed
        assert "!" in result["pivot"]["target_cell"]


# ═════════════════════════════════════════════════
# read_freeze_panes_actual
# ═════════════════════════════════════════════════


class TestReadFreezePanesActual:
    def test_freeze_panes_roundtrip(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "freeze.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.set_freeze_panes(wb, "S1", {"mode": "freeze", "top_left_cell": "B2"})
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        expected: JSONDict = {"freeze": {"mode": "freeze", "top_left_cell": "B2"}}
        result = read_freeze_panes_actual(opxl, wb2, "S1", expected)
        assert "freeze" in result
        opxl.close_workbook(wb2)


# ═════════════════════════════════════════════════
# test_read_case — feature dispatch branches
# ═════════════════════════════════════════════════


class TestReadCaseFeatureDispatch:
    """Exercise the feature dispatch branches in test_read_case."""

    def test_conditional_formatting_dispatch(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "cf.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.add_conditional_format(
            wb,
            "S1",
            {
                "cf_rule": {
                    "range": "A1:A10",
                    "rule_type": "cellIs",
                    "operator": "greaterThan",
                    "formula": "5",
                }
            },
        )
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        tc = _tc(
            "t1", {"cf_rule": {"rule_type": "cellIs", "operator": "greaterThan", "formula": "5"}}
        )
        result = _test_read_case(opxl, wb2, "S1", tc, "conditional_formatting", OperationType.READ)
        # The dispatch to read_conditional_format_actual is exercised
        assert result is not None
        opxl.close_workbook(wb2)

    def test_data_validation_dispatch(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "dv.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.add_data_validation(
            wb,
            "S1",
            {
                "validation": {
                    "range": "A1:A5",
                    "validation_type": "list",
                    "formula1": '"A,B,C"',
                }
            },
        )
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        tc = _tc("t1", {"validation": {"validation_type": "list"}})
        result = _test_read_case(opxl, wb2, "S1", tc, "data_validation", OperationType.READ)
        assert result is not None
        opxl.close_workbook(wb2)

    def test_hyperlinks_dispatch(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "link.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.add_hyperlink(
            wb,
            "S1",
            {"hyperlink": {"cell": "A1", "target": "https://example.com", "display": "Link"}},
        )
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        tc = _tc("t1", {"hyperlink": {"cell": "A1", "target": "https://example.com"}})
        result = _test_read_case(opxl, wb2, "S1", tc, "hyperlinks", OperationType.READ)
        assert result is not None
        opxl.close_workbook(wb2)

    def test_images_dispatch(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        """Images dispatch — no matching image returns empty actual."""
        path = tmp_path / "img.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        tc = _tc("t1", {"image": {"cell": "A1", "path": "test.png"}})
        result = _test_read_case(opxl, wb2, "S1", tc, "images", OperationType.READ)
        assert result.passed is False
        opxl.close_workbook(wb2)

    def test_comments_dispatch(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "comment.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.add_comment(wb, "S1", {"comment": {"cell": "A1", "text": "Note"}})
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        tc = _tc("t1", {"comment": {"cell": "A1", "text": "Note"}})
        result = _test_read_case(opxl, wb2, "S1", tc, "comments", OperationType.READ)
        assert result is not None
        opxl.close_workbook(wb2)

    def test_freeze_panes_dispatch(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "freeze.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.set_freeze_panes(wb, "S1", {"mode": "freeze", "top_left_cell": "B2"})
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        tc = _tc("t1", {"freeze": {"mode": "freeze", "top_left_cell": "B2"}})
        result = _test_read_case(opxl, wb2, "S1", tc, "freeze_panes", OperationType.READ)
        assert result is not None
        opxl.close_workbook(wb2)


# ═════════════════════════════════════════════════
# compare_results / _deep_compare edge cases
# ═════════════════════════════════════════════════


class TestDeepCompare:
    def test_color_string_match(self) -> None:
        """Color strings starting with # should be case-insensitive."""
        assert compare_results({"color": "#FF0000"}, {"color": "#ff0000"}) is True

    def test_color_string_non_string_actual(self) -> None:
        """# prefixed expected with non-string actual → False."""
        assert compare_results({"color": "#FF0000"}, {"color": 123}) is False

    def test_tuple_vs_list(self) -> None:
        """Tuple expected should match list actual."""
        assert compare_results({"values": (1, 2, 3)}, {"values": [1, 2, 3]}) is True

    def test_tuple_vs_non_sequence(self) -> None:
        """Tuple expected vs non-sequence actual → False."""
        assert compare_results({"values": (1, 2)}, {"values": "nope"}) is False

    def test_numeric_tolerance(self) -> None:
        """Float comparison with tolerance."""
        assert compare_results({"height": 15.0}, {"height": 15.00005}) is True

    def test_error_in_actual(self) -> None:
        """Error key in actual always returns False."""
        assert compare_results({"value": "x"}, {"error": "oops"}) is False


# ═════════════════════════════════════════════════
# run_benchmark edge cases
# ═════════════════════════════════════════════════


class TestRunBenchmarkEdgeCases:
    def test_missing_manifest(self, tmp_path: Path) -> None:
        """Missing manifest.json should raise FileNotFoundError."""
        with pytest.raises(FileNotFoundError, match="Manifest not found"):
            run_benchmark(test_dir=tmp_path)

    def test_no_matching_features(self, tmp_path: Path) -> None:
        """Requesting non-existent feature should raise ValueError."""
        import json

        manifest = {
            "generated_at": "2024-01-01T00:00:00",
            "excel_version": "test",
            "generator_version": "1.0",
            "files": [
                {
                    "path": "cell_values.xlsx",
                    "feature": "cell_values",
                    "tier": 1,
                    "test_cases": [],
                }
            ],
        }
        (tmp_path / "manifest.json").write_text(json.dumps(manifest))

        with pytest.raises(ValueError, match="No matching features"):
            run_benchmark(test_dir=tmp_path, features=["nonexistent_feature"])

    def test_missing_test_file(self, tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
        """Missing .xlsx test file should print warning and skip."""
        import json

        manifest = {
            "generated_at": "2024-01-01T00:00:00",
            "excel_version": "test",
            "generator_version": "1.0",
            "files": [
                {
                    "path": "missing_file.xlsx",
                    "feature": "cell_values",
                    "tier": 1,
                    "test_cases": [
                        {
                            "id": "t1",
                            "label": "test",
                            "row": 2,
                            "expected": {"type": "string", "value": "x"},
                        },
                    ],
                }
            ],
        }
        (tmp_path / "manifest.json").write_text(json.dumps(manifest))

        result = run_benchmark(test_dir=tmp_path, features=["cell_values"])
        captured = capsys.readouterr()
        assert "Warning" in captured.out or len(result.scores) == 0
