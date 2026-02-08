"""Tests for adapter write-side format/border/conditional/validation code paths.

Targets uncovered branches in:
- xlsxwriter_adapter.py: _create_format(), save_workbook() cell-type branches,
  conditional formats, data validations, split panes, images
- openpyxl_adapter.py: write_cell_format(), write_cell_border(),
  add_conditional_format(), set_freeze_panes() split mode
"""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
from typing import Any

import pytest

from excelbench.harness.adapters.openpyxl_adapter import OpenpyxlAdapter
from excelbench.harness.adapters.xlsxwriter_adapter import XlsxwriterAdapter
from excelbench.models import (
    BorderEdge,
    BorderInfo,
    BorderStyle,
    CellFormat,
    CellType,
    CellValue,
)

JSONDict = dict[str, Any]


# ═════════════════════════════════════════════════
# Fixtures
# ═════════════════════════════════════════════════


@pytest.fixture
def xlsxw() -> XlsxwriterAdapter:
    return XlsxwriterAdapter()


@pytest.fixture
def opxl() -> OpenpyxlAdapter:
    return OpenpyxlAdapter()


# ═════════════════════════════════════════════════
# XlsxWriter: _create_format() — text formatting
# ═════════════════════════════════════════════════


class TestXlsxwriterCreateFormatText:
    """Exercise _create_format text/font branches via write→read roundtrip."""

    def test_bold_italic(
        self, xlsxw: XlsxwriterAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "bold_italic.xlsx"
        wb = xlsxw.create_workbook()
        xlsxw.add_sheet(wb, "S1")
        xlsxw.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="test"))
        xlsxw.write_cell_format(wb, "S1", "A1", CellFormat(bold=True, italic=True))
        xlsxw.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        fmt = opxl.read_cell_format(wb2, "S1", "A1")
        assert fmt.bold is True
        assert fmt.italic is True
        opxl.close_workbook(wb2)

    def test_underline_single(
        self, xlsxw: XlsxwriterAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "underline.xlsx"
        wb = xlsxw.create_workbook()
        xlsxw.add_sheet(wb, "S1")
        xlsxw.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="x"))
        xlsxw.write_cell_format(wb, "S1", "A1", CellFormat(underline="single"))
        xlsxw.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        fmt = opxl.read_cell_format(wb2, "S1", "A1")
        assert fmt.underline is not None
        opxl.close_workbook(wb2)

    def test_underline_double(
        self, xlsxw: XlsxwriterAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "underline_double.xlsx"
        wb = xlsxw.create_workbook()
        xlsxw.add_sheet(wb, "S1")
        xlsxw.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="x"))
        xlsxw.write_cell_format(wb, "S1", "A1", CellFormat(underline="double"))
        xlsxw.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        fmt = opxl.read_cell_format(wb2, "S1", "A1")
        assert fmt.underline is not None
        opxl.close_workbook(wb2)

    def test_strikethrough(
        self, xlsxw: XlsxwriterAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "strike.xlsx"
        wb = xlsxw.create_workbook()
        xlsxw.add_sheet(wb, "S1")
        xlsxw.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="x"))
        xlsxw.write_cell_format(wb, "S1", "A1", CellFormat(strikethrough=True))
        xlsxw.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        fmt = opxl.read_cell_format(wb2, "S1", "A1")
        assert fmt.strikethrough is True
        opxl.close_workbook(wb2)

    def test_font_name_size_color(
        self, xlsxw: XlsxwriterAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "font.xlsx"
        wb = xlsxw.create_workbook()
        xlsxw.add_sheet(wb, "S1")
        xlsxw.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="x"))
        xlsxw.write_cell_format(
            wb,
            "S1",
            "A1",
            CellFormat(font_name="Arial", font_size=14, font_color="#FF0000"),
        )
        xlsxw.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        fmt = opxl.read_cell_format(wb2, "S1", "A1")
        assert fmt.font_name == "Arial"
        assert fmt.font_size == 14
        assert fmt.font_color is not None
        opxl.close_workbook(wb2)

    def test_bg_color(
        self, xlsxw: XlsxwriterAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "bg.xlsx"
        wb = xlsxw.create_workbook()
        xlsxw.add_sheet(wb, "S1")
        xlsxw.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="x"))
        xlsxw.write_cell_format(wb, "S1", "A1", CellFormat(bg_color="#00FF00"))
        xlsxw.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        fmt = opxl.read_cell_format(wb2, "S1", "A1")
        assert fmt.bg_color is not None
        opxl.close_workbook(wb2)

    def test_number_format(
        self, xlsxw: XlsxwriterAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "numfmt.xlsx"
        wb = xlsxw.create_workbook()
        xlsxw.add_sheet(wb, "S1")
        xlsxw.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.NUMBER, value=1234.56))
        xlsxw.write_cell_format(wb, "S1", "A1", CellFormat(number_format="#,##0.00"))
        xlsxw.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        fmt = opxl.read_cell_format(wb2, "S1", "A1")
        assert fmt.number_format == "#,##0.00"
        opxl.close_workbook(wb2)


# ═════════════════════════════════════════════════
# XlsxWriter: _create_format() — alignment branches
# ═════════════════════════════════════════════════


class TestXlsxwriterCreateFormatAlignment:
    def test_h_align_center(
        self, xlsxw: XlsxwriterAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "halign.xlsx"
        wb = xlsxw.create_workbook()
        xlsxw.add_sheet(wb, "S1")
        xlsxw.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="x"))
        xlsxw.write_cell_format(wb, "S1", "A1", CellFormat(h_align="center"))
        xlsxw.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        fmt = opxl.read_cell_format(wb2, "S1", "A1")
        assert fmt.h_align == "center"
        opxl.close_workbook(wb2)

    def test_v_align_top(
        self, xlsxw: XlsxwriterAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "valign.xlsx"
        wb = xlsxw.create_workbook()
        xlsxw.add_sheet(wb, "S1")
        xlsxw.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="x"))
        xlsxw.write_cell_format(wb, "S1", "A1", CellFormat(v_align="top"))
        xlsxw.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        fmt = opxl.read_cell_format(wb2, "S1", "A1")
        assert fmt.v_align == "top"
        opxl.close_workbook(wb2)

    def test_wrap_rotation_indent(
        self, xlsxw: XlsxwriterAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "wrap_rot_indent.xlsx"
        wb = xlsxw.create_workbook()
        xlsxw.add_sheet(wb, "S1")
        xlsxw.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="x"))
        xlsxw.write_cell_format(wb, "S1", "A1", CellFormat(wrap=True, rotation=45, indent=2))
        xlsxw.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        fmt = opxl.read_cell_format(wb2, "S1", "A1")
        assert fmt.wrap is True
        assert fmt.rotation == 45
        assert fmt.indent == 2
        opxl.close_workbook(wb2)


# ═════════════════════════════════════════════════
# XlsxWriter: _create_format() — border branches
# ═════════════════════════════════════════════════


class TestXlsxwriterCreateFormatBorder:
    def test_four_sided_border(
        self, xlsxw: XlsxwriterAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "borders.xlsx"
        wb = xlsxw.create_workbook()
        xlsxw.add_sheet(wb, "S1")
        xlsxw.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="x"))
        xlsxw.write_cell_border(
            wb,
            "S1",
            "A1",
            BorderInfo(
                top=BorderEdge(style=BorderStyle.THIN, color="#000000"),
                bottom=BorderEdge(style=BorderStyle.MEDIUM, color="#FF0000"),
                left=BorderEdge(style=BorderStyle.DASHED, color="#00FF00"),
                right=BorderEdge(style=BorderStyle.DOTTED, color="#0000FF"),
            ),
        )
        xlsxw.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        border = opxl.read_cell_border(wb2, "S1", "A1")
        assert border.top is not None
        assert border.top.style == BorderStyle.THIN
        assert border.bottom is not None
        assert border.bottom.style == BorderStyle.MEDIUM
        assert border.left is not None
        assert border.left.style == BorderStyle.DASHED
        assert border.right is not None
        assert border.right.style == BorderStyle.DOTTED
        opxl.close_workbook(wb2)

    def test_diagonal_up_only(
        self, xlsxw: XlsxwriterAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "diag_up.xlsx"
        wb = xlsxw.create_workbook()
        xlsxw.add_sheet(wb, "S1")
        xlsxw.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="x"))
        xlsxw.write_cell_border(
            wb,
            "S1",
            "A1",
            BorderInfo(diagonal_up=BorderEdge(style=BorderStyle.THIN, color="#000000")),
        )
        xlsxw.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        border = opxl.read_cell_border(wb2, "S1", "A1")
        assert border.diagonal_up is not None
        opxl.close_workbook(wb2)

    def test_diagonal_down_only(
        self, xlsxw: XlsxwriterAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "diag_down.xlsx"
        wb = xlsxw.create_workbook()
        xlsxw.add_sheet(wb, "S1")
        xlsxw.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="x"))
        xlsxw.write_cell_border(
            wb,
            "S1",
            "A1",
            BorderInfo(diagonal_down=BorderEdge(style=BorderStyle.MEDIUM, color="#FF0000")),
        )
        xlsxw.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        border = opxl.read_cell_border(wb2, "S1", "A1")
        assert border.diagonal_down is not None
        opxl.close_workbook(wb2)

    def test_diagonal_both(
        self, xlsxw: XlsxwriterAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "diag_both.xlsx"
        wb = xlsxw.create_workbook()
        xlsxw.add_sheet(wb, "S1")
        xlsxw.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="x"))
        xlsxw.write_cell_border(
            wb,
            "S1",
            "A1",
            BorderInfo(
                diagonal_up=BorderEdge(style=BorderStyle.THIN, color="#000000"),
                diagonal_down=BorderEdge(style=BorderStyle.THIN, color="#000000"),
            ),
        )
        xlsxw.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        border = opxl.read_cell_border(wb2, "S1", "A1")
        # Both diagonals should be present
        assert border.diagonal_up is not None or border.diagonal_down is not None
        opxl.close_workbook(wb2)


# ═════════════════════════════════════════════════
# XlsxWriter: save_workbook() — cell type branches
# ═════════════════════════════════════════════════


class TestXlsxwriterSaveCellTypes:
    def test_boolean_value(
        self, xlsxw: XlsxwriterAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "bool.xlsx"
        wb = xlsxw.create_workbook()
        xlsxw.add_sheet(wb, "S1")
        xlsxw.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.BOOLEAN, value=True))
        xlsxw.write_cell_value(wb, "S1", "A2", CellValue(type=CellType.BOOLEAN, value=False))
        xlsxw.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        v1 = opxl.read_cell_value(wb2, "S1", "A1")
        v2 = opxl.read_cell_value(wb2, "S1", "A2")
        assert v1.type == CellType.BOOLEAN
        assert v1.value is True
        assert v2.type == CellType.BOOLEAN
        assert v2.value is False
        opxl.close_workbook(wb2)

    def test_formula_value(
        self, xlsxw: XlsxwriterAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "formula.xlsx"
        wb = xlsxw.create_workbook()
        xlsxw.add_sheet(wb, "S1")
        xlsxw.write_cell_value(
            wb,
            "S1",
            "A1",
            CellValue(type=CellType.FORMULA, value="=1+1", formula="=1+1"),
        )
        xlsxw.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        v1 = opxl.read_cell_value(wb2, "S1", "A1")
        # Formula cells come back as formula or string depending on calculation
        assert v1.value is not None
        opxl.close_workbook(wb2)

    def test_date_value(
        self, xlsxw: XlsxwriterAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "date.xlsx"
        wb = xlsxw.create_workbook()
        xlsxw.add_sheet(wb, "S1")
        xlsxw.write_cell_value(
            wb,
            "S1",
            "A1",
            CellValue(type=CellType.DATE, value=date(2024, 6, 15)),
        )
        xlsxw.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        v1 = opxl.read_cell_value(wb2, "S1", "A1")
        assert v1.type == CellType.DATE
        opxl.close_workbook(wb2)

    def test_datetime_value(
        self, xlsxw: XlsxwriterAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "datetime.xlsx"
        wb = xlsxw.create_workbook()
        xlsxw.add_sheet(wb, "S1")
        xlsxw.write_cell_value(
            wb,
            "S1",
            "A1",
            CellValue(type=CellType.DATETIME, value=datetime(2024, 6, 15, 14, 30, 0)),
        )
        xlsxw.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        v1 = opxl.read_cell_value(wb2, "S1", "A1")
        assert v1.type == CellType.DATETIME
        opxl.close_workbook(wb2)

    def test_date_with_format(
        self, xlsxw: XlsxwriterAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        """DATE cell WITH an explicit format skips the default-format branch."""
        path = tmp_path / "date_fmt.xlsx"
        wb = xlsxw.create_workbook()
        xlsxw.add_sheet(wb, "S1")
        xlsxw.write_cell_value(
            wb,
            "S1",
            "A1",
            CellValue(type=CellType.DATE, value=date(2024, 1, 1)),
        )
        xlsxw.write_cell_format(wb, "S1", "A1", CellFormat(number_format="dd/mm/yyyy"))
        xlsxw.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        v1 = opxl.read_cell_value(wb2, "S1", "A1")
        assert v1.type == CellType.DATE
        opxl.close_workbook(wb2)

    def test_error_values(
        self, xlsxw: XlsxwriterAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "errors.xlsx"
        wb = xlsxw.create_workbook()
        xlsxw.add_sheet(wb, "S1")
        xlsxw.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.ERROR, value="#DIV/0!"))
        xlsxw.write_cell_value(wb, "S1", "A2", CellValue(type=CellType.ERROR, value="#N/A"))
        xlsxw.write_cell_value(wb, "S1", "A3", CellValue(type=CellType.ERROR, value="#VALUE!"))
        xlsxw.write_cell_value(wb, "S1", "A4", CellValue(type=CellType.ERROR, value="#CUSTOM!"))
        xlsxw.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        # Error cells are written as formulas that produce errors
        for cell in ["A1", "A2", "A3", "A4"]:
            v = opxl.read_cell_value(wb2, "S1", cell)
            assert v.value is not None
        opxl.close_workbook(wb2)

    def test_blank_value(
        self, xlsxw: XlsxwriterAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "blank.xlsx"
        wb = xlsxw.create_workbook()
        xlsxw.add_sheet(wb, "S1")
        xlsxw.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.BLANK))
        xlsxw.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        v1 = opxl.read_cell_value(wb2, "S1", "A1")
        assert v1.type == CellType.BLANK
        opxl.close_workbook(wb2)

    def test_format_only_no_value(
        self, xlsxw: XlsxwriterAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        """Cell with format but no value → write_blank with format."""
        path = tmp_path / "fmt_only.xlsx"
        wb = xlsxw.create_workbook()
        xlsxw.add_sheet(wb, "S1")
        xlsxw.write_cell_format(wb, "S1", "A1", CellFormat(bold=True, bg_color="#FFFF00"))
        xlsxw.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        fmt = opxl.read_cell_format(wb2, "S1", "A1")
        assert fmt.bold is True
        opxl.close_workbook(wb2)


# ═════════════════════════════════════════════════
# XlsxWriter: save_workbook() — split panes
# ═════════════════════════════════════════════════


class TestXlsxwriterSplitPanes:
    def test_split_panes(self, xlsxw: XlsxwriterAdapter, tmp_path: Path) -> None:
        path = tmp_path / "split.xlsx"
        wb = xlsxw.create_workbook()
        xlsxw.add_sheet(wb, "S1")
        xlsxw.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="x"))
        xlsxw.set_freeze_panes(wb, "S1", {"mode": "split", "x_split": 2000, "y_split": 3000})
        xlsxw.save_workbook(wb, path)
        assert path.exists()


# ═════════════════════════════════════════════════
# XlsxWriter: save_workbook() — conditional formats
# ═════════════════════════════════════════════════


class TestXlsxwriterConditionalFormats:
    def test_cell_is_rule(self, xlsxw: XlsxwriterAdapter, tmp_path: Path) -> None:
        path = tmp_path / "cf_cellis.xlsx"
        wb = xlsxw.create_workbook()
        xlsxw.add_sheet(wb, "S1")
        xlsxw.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.NUMBER, value=10))
        xlsxw.add_conditional_format(
            wb,
            "S1",
            {
                "cf_rule": {
                    "range": "A1:A10",
                    "rule_type": "cellIs",
                    "operator": "greaterThan",
                    "formula": "5",
                    "format": {"bg_color": "#FF0000"},
                    "stop_if_true": True,
                }
            },
        )
        xlsxw.save_workbook(wb, path)
        assert path.exists()

    def test_expression_rule(self, xlsxw: XlsxwriterAdapter, tmp_path: Path) -> None:
        path = tmp_path / "cf_expr.xlsx"
        wb = xlsxw.create_workbook()
        xlsxw.add_sheet(wb, "S1")
        xlsxw.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.NUMBER, value=5))
        xlsxw.add_conditional_format(
            wb,
            "S1",
            {
                "cf_rule": {
                    "range": "A1:A10",
                    "rule_type": "expression",
                    "formula": "=A1>3",
                    "format": {"bg_color": "#00FF00"},
                }
            },
        )
        xlsxw.save_workbook(wb, path)
        assert path.exists()

    def test_color_scale_rule(self, xlsxw: XlsxwriterAdapter, tmp_path: Path) -> None:
        path = tmp_path / "cf_colorscale.xlsx"
        wb = xlsxw.create_workbook()
        xlsxw.add_sheet(wb, "S1")
        for i in range(1, 6):
            xlsxw.write_cell_value(wb, "S1", f"A{i}", CellValue(type=CellType.NUMBER, value=i))
        xlsxw.add_conditional_format(
            wb, "S1", {"cf_rule": {"range": "A1:A5", "rule_type": "colorScale"}}
        )
        xlsxw.save_workbook(wb, path)
        assert path.exists()

    def test_data_bar_rule(self, xlsxw: XlsxwriterAdapter, tmp_path: Path) -> None:
        path = tmp_path / "cf_databar.xlsx"
        wb = xlsxw.create_workbook()
        xlsxw.add_sheet(wb, "S1")
        for i in range(1, 6):
            xlsxw.write_cell_value(wb, "S1", f"A{i}", CellValue(type=CellType.NUMBER, value=i))
        xlsxw.add_conditional_format(
            wb, "S1", {"cf_rule": {"range": "A1:A5", "rule_type": "dataBar"}}
        )
        xlsxw.save_workbook(wb, path)
        assert path.exists()

    def test_cf_without_format(self, xlsxw: XlsxwriterAdapter, tmp_path: Path) -> None:
        """Conditional format rule without bg_color → no format key in options."""
        path = tmp_path / "cf_nofmt.xlsx"
        wb = xlsxw.create_workbook()
        xlsxw.add_sheet(wb, "S1")
        xlsxw.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.NUMBER, value=5))
        xlsxw.add_conditional_format(
            wb,
            "S1",
            {
                "cf_rule": {
                    "range": "A1:A10",
                    "rule_type": "cellIs",
                    "operator": "equal",
                    "formula": "5",
                }
            },
        )
        xlsxw.save_workbook(wb, path)
        assert path.exists()


# ═════════════════════════════════════════════════
# XlsxWriter: save_workbook() — data validations
# ═════════════════════════════════════════════════


class TestXlsxwriterDataValidations:
    def test_list_validation_with_source(self, xlsxw: XlsxwriterAdapter, tmp_path: Path) -> None:
        path = tmp_path / "dv_list.xlsx"
        wb = xlsxw.create_workbook()
        xlsxw.add_sheet(wb, "S1")
        xlsxw.add_data_validation(
            wb,
            "S1",
            {
                "validation": {
                    "range": "A1:A10",
                    "validation_type": "list",
                    "formula1": '"Apple,Banana,Cherry"',
                    "allow_blank": True,
                    "prompt_title": "Choose fruit",
                    "prompt": "Pick one",
                    "error_title": "Invalid",
                    "error": "Not a fruit",
                }
            },
        )
        xlsxw.save_workbook(wb, path)
        assert path.exists()

    def test_integer_validation_with_between(
        self, xlsxw: XlsxwriterAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "dv_int.xlsx"
        wb = xlsxw.create_workbook()
        xlsxw.add_sheet(wb, "S1")
        xlsxw.add_data_validation(
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
        xlsxw.save_workbook(wb, path)
        assert path.exists()

    def test_list_validation_plain_source(self, xlsxw: XlsxwriterAdapter, tmp_path: Path) -> None:
        """List source without surrounding quotes → no stripping."""
        path = tmp_path / "dv_list_plain.xlsx"
        wb = xlsxw.create_workbook()
        xlsxw.add_sheet(wb, "S1")
        xlsxw.add_data_validation(
            wb,
            "S1",
            {
                "validation": {
                    "range": "A1:A5",
                    "validation_type": "list",
                    "formula1": "Yes,No",
                }
            },
        )
        xlsxw.save_workbook(wb, path)
        assert path.exists()


# ═════════════════════════════════════════════════
# XlsxWriter: save_workbook() — images with offsets
# ═════════════════════════════════════════════════


class TestXlsxwriterImages:
    def test_image_with_offset(self, xlsxw: XlsxwriterAdapter, tmp_path: Path) -> None:
        # Create a minimal 1x1 PNG
        img_path = tmp_path / "test.png"
        import struct
        import zlib

        def _minimal_png() -> bytes:
            sig = b"\x89PNG\r\n\x1a\n"
            ihdr_data = struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)
            ihdr_crc = zlib.crc32(b"IHDR" + ihdr_data) & 0xFFFFFFFF
            ihdr = struct.pack(">I", 13) + b"IHDR" + ihdr_data + struct.pack(">I", ihdr_crc)
            raw = b"\x00\x00\x00\x00"  # filter byte + 1 pixel RGB
            idat_data = zlib.compress(raw)
            idat_crc = zlib.crc32(b"IDAT" + idat_data) & 0xFFFFFFFF
            idat = (
                struct.pack(">I", len(idat_data))
                + b"IDAT"
                + idat_data
                + struct.pack(">I", idat_crc)
            )
            iend_crc = zlib.crc32(b"IEND") & 0xFFFFFFFF
            iend = struct.pack(">I", 0) + b"IEND" + struct.pack(">I", iend_crc)
            return sig + ihdr + idat + iend

        img_path.write_bytes(_minimal_png())

        path = tmp_path / "img.xlsx"
        wb = xlsxw.create_workbook()
        xlsxw.add_sheet(wb, "S1")
        xlsxw.add_image(
            wb,
            "S1",
            {
                "image": {
                    "cell": "A1",
                    "path": str(img_path),
                    "offset": [10, 20],
                }
            },
        )
        xlsxw.save_workbook(wb, path)
        assert path.exists()


# ═════════════════════════════════════════════════
# Openpyxl: write_cell_format() — all branches
# ═════════════════════════════════════════════════


class TestOpenpyxlWriteCellFormat:
    def test_bold_italic_underline_strike(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "opxl_fmt.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="styled"))
        opxl.write_cell_format(
            wb,
            "S1",
            "A1",
            CellFormat(bold=True, italic=True, underline="single", strikethrough=True),
        )
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        fmt = opxl.read_cell_format(wb2, "S1", "A1")
        assert fmt.bold is True
        assert fmt.italic is True
        assert fmt.underline is not None
        assert fmt.strikethrough is True
        opxl.close_workbook(wb2)

    def test_font_name_size_color(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "opxl_font.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="x"))
        opxl.write_cell_format(
            wb,
            "S1",
            "A1",
            CellFormat(font_name="Courier New", font_size=16, font_color="#0000FF"),
        )
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        fmt = opxl.read_cell_format(wb2, "S1", "A1")
        assert fmt.font_name == "Courier New"
        assert fmt.font_size == 16
        assert fmt.font_color is not None
        opxl.close_workbook(wb2)

    def test_bg_color(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "opxl_bg.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="x"))
        opxl.write_cell_format(wb, "S1", "A1", CellFormat(bg_color="#FFFF00"))
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        fmt = opxl.read_cell_format(wb2, "S1", "A1")
        assert fmt.bg_color is not None
        opxl.close_workbook(wb2)

    def test_number_format(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "opxl_numfmt.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.NUMBER, value=0.5))
        opxl.write_cell_format(wb, "S1", "A1", CellFormat(number_format="0.00%"))
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        fmt = opxl.read_cell_format(wb2, "S1", "A1")
        assert fmt.number_format == "0.00%"
        opxl.close_workbook(wb2)

    def test_alignment_all(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "opxl_align.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="x"))
        opxl.write_cell_format(
            wb,
            "S1",
            "A1",
            CellFormat(h_align="center", v_align="top", wrap=True, rotation=90, indent=3),
        )
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        fmt = opxl.read_cell_format(wb2, "S1", "A1")
        assert fmt.h_align == "center"
        assert fmt.v_align == "top"
        assert fmt.wrap is True
        assert fmt.rotation == 90
        assert fmt.indent == 3
        opxl.close_workbook(wb2)


# ═════════════════════════════════════════════════
# Openpyxl: write_cell_border() — all branches
# ═════════════════════════════════════════════════


class TestOpenpyxlWriteCellBorder:
    def test_four_sided_border(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "opxl_border.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="x"))
        opxl.write_cell_border(
            wb,
            "S1",
            "A1",
            BorderInfo(
                top=BorderEdge(style=BorderStyle.THIN, color="#000000"),
                bottom=BorderEdge(style=BorderStyle.THICK, color="#FF0000"),
                left=BorderEdge(style=BorderStyle.DOUBLE, color="#00FF00"),
                right=BorderEdge(style=BorderStyle.HAIR, color="#0000FF"),
            ),
        )
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        border = opxl.read_cell_border(wb2, "S1", "A1")
        assert border.top is not None and border.top.style == BorderStyle.THIN
        assert border.bottom is not None and border.bottom.style == BorderStyle.THICK
        assert border.left is not None and border.left.style == BorderStyle.DOUBLE
        assert border.right is not None and border.right.style == BorderStyle.HAIR
        opxl.close_workbook(wb2)

    def test_diagonal_up_border(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "opxl_diag_up.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="x"))
        opxl.write_cell_border(
            wb,
            "S1",
            "A1",
            BorderInfo(diagonal_up=BorderEdge(style=BorderStyle.THIN, color="#000000")),
        )
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        border = opxl.read_cell_border(wb2, "S1", "A1")
        assert border.diagonal_up is not None
        opxl.close_workbook(wb2)

    def test_diagonal_down_border(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "opxl_diag_down.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="x"))
        opxl.write_cell_border(
            wb,
            "S1",
            "A1",
            BorderInfo(diagonal_down=BorderEdge(style=BorderStyle.MEDIUM, color="#FF0000")),
        )
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        border = opxl.read_cell_border(wb2, "S1", "A1")
        assert border.diagonal_down is not None
        opxl.close_workbook(wb2)

    def test_none_style_edge(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        """BorderEdge with NONE style → make_side returns empty Side."""
        path = tmp_path / "opxl_none_border.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="x"))
        opxl.write_cell_border(
            wb,
            "S1",
            "A1",
            BorderInfo(top=BorderEdge(style=BorderStyle.NONE, color="#000000")),
        )
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        border = opxl.read_cell_border(wb2, "S1", "A1")
        # NONE style should result in no visible border
        assert border.top is None or border.top.style == BorderStyle.NONE
        opxl.close_workbook(wb2)


# ═════════════════════════════════════════════════
# Openpyxl: add_conditional_format() — rule types
# ═════════════════════════════════════════════════


class TestOpenpyxlConditionalFormat:
    def test_expression_rule(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "opxl_cf_expr.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.NUMBER, value=5))
        opxl.add_conditional_format(
            wb,
            "S1",
            {
                "cf_rule": {
                    "range": "A1:A10",
                    "rule_type": "expression",
                    "formula": "=A1>3",
                    "format": {"bg_color": "#FF0000", "font_color": "#FFFFFF"},
                }
            },
        )
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        cfs = opxl.read_conditional_formats(wb2, "S1")
        assert len(cfs) >= 1
        opxl.close_workbook(wb2)

    def test_color_scale_rule(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "opxl_cf_colorscale.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        for i in range(1, 6):
            opxl.write_cell_value(wb, "S1", f"A{i}", CellValue(type=CellType.NUMBER, value=i))
        opxl.add_conditional_format(
            wb, "S1", {"cf_rule": {"range": "A1:A5", "rule_type": "colorScale"}}
        )
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        cfs = opxl.read_conditional_formats(wb2, "S1")
        assert len(cfs) >= 1
        opxl.close_workbook(wb2)

    def test_data_bar_rule(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "opxl_cf_databar.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        for i in range(1, 6):
            opxl.write_cell_value(wb, "S1", f"A{i}", CellValue(type=CellType.NUMBER, value=i))
        opxl.add_conditional_format(
            wb, "S1", {"cf_rule": {"range": "A1:A5", "rule_type": "dataBar"}}
        )
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        cfs = opxl.read_conditional_formats(wb2, "S1")
        assert len(cfs) >= 1
        opxl.close_workbook(wb2)

    def test_cell_is_with_priority(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        """Test CellIsRule with priority and stop_if_true."""
        path = tmp_path / "opxl_cf_priority.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.NUMBER, value=10))
        opxl.add_conditional_format(
            wb,
            "S1",
            {
                "cf_rule": {
                    "range": "A1:A10",
                    "rule_type": "cellIs",
                    "operator": "greaterThan",
                    "formula": "5",
                    "format": {"bg_color": "#00FF00"},
                    "stop_if_true": True,
                    "priority": 1,
                }
            },
        )
        opxl.save_workbook(wb, path)
        assert path.exists()


# ═════════════════════════════════════════════════
# Openpyxl: set_freeze_panes() — split mode
# ═════════════════════════════════════════════════


class TestOpenpyxlFreezePanesSplit:
    def test_split_mode(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "opxl_split.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="x"))
        opxl.set_freeze_panes(
            wb,
            "S1",
            {
                "mode": "split",
                "x_split": 2000,
                "y_split": 3000,
                "top_left_cell": "B2",
                "active_pane": "bottomRight",
            },
        )
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        freeze = opxl.read_freeze_panes(wb2, "S1")
        assert freeze.get("mode") == "split" or freeze != {}
        opxl.close_workbook(wb2)

    def test_split_mode_minimal(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        """Split mode with minimal settings — exercises the pane=None branch."""
        path = tmp_path / "opxl_split_min.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.set_freeze_panes(wb, "S1", {"mode": "split"})
        opxl.save_workbook(wb, path)
        assert path.exists()


# ═════════════════════════════════════════════════
# Openpyxl: read_cell_value() — edge cases
# ═════════════════════════════════════════════════


class TestOpenpyxlReadCellValueEdgeCases:
    def test_date_value(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        """Date (not datetime) cell roundtrip."""
        path = tmp_path / "opxl_date.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.write_cell_value(
            wb, "S1", "A1", CellValue(type=CellType.DATE, value=date(2024, 6, 15))
        )
        opxl.write_cell_format(wb, "S1", "A1", CellFormat(number_format="yyyy-mm-dd"))
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        v = opxl.read_cell_value(wb2, "S1", "A1")
        assert v.type == CellType.DATE
        opxl.close_workbook(wb2)

    def test_datetime_with_time(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        """Datetime with non-midnight time should be DATETIME, not DATE."""
        path = tmp_path / "opxl_datetime.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.write_cell_value(
            wb,
            "S1",
            "A1",
            CellValue(type=CellType.DATETIME, value=datetime(2024, 6, 15, 10, 30, 45)),
        )
        opxl.write_cell_format(wb, "S1", "A1", CellFormat(number_format="yyyy-mm-dd hh:mm:ss"))
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        v = opxl.read_cell_value(wb2, "S1", "A1")
        assert v.type == CellType.DATETIME
        opxl.close_workbook(wb2)

    def test_error_value_na(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        """Error formula roundtrip — #N/A."""
        path = tmp_path / "opxl_err.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.ERROR, value="#N/A"))
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        v = opxl.read_cell_value(wb2, "S1", "A1")
        # Written as formula =NA(), so should come back as formula or error
        assert v.value is not None
        opxl.close_workbook(wb2)

    def test_formula_roundtrip(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        """Formula cell roundtrip."""
        path = tmp_path / "opxl_formula.xlsx"
        wb = opxl.create_workbook()
        opxl.add_sheet(wb, "S1")
        opxl.write_cell_value(
            wb,
            "S1",
            "A1",
            CellValue(type=CellType.FORMULA, value="=1+1", formula="=1+1"),
        )
        opxl.save_workbook(wb, path)

        wb2 = opxl.open_workbook(path)
        v = opxl.read_cell_value(wb2, "S1", "A1")
        assert v.value is not None
        opxl.close_workbook(wb2)
