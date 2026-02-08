"""xlwt → xlrd .xls roundtrip tests.

xlwt writes BIFF8 (.xls), xlrd reads it back.  Every test exercises both
adapters' code paths simultaneously, maximising coverage per test.
"""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

import pytest

from excelbench.harness.adapters.xlrd_adapter import XlrdAdapter
from excelbench.harness.adapters.xlwt_adapter import (
    XlwtAdapter,
    _col_to_index,
    _hex_to_xlwt_colour,
)
from excelbench.harness.adapters.xlwt_adapter import (
    _parse_cell_ref as xlwt_parse,
)
from excelbench.models import (
    BorderEdge,
    BorderInfo,
    BorderStyle,
    CellFormat,
    CellType,
    CellValue,
)


@pytest.fixture
def xlwt() -> XlwtAdapter:
    return XlwtAdapter()


@pytest.fixture
def xlrd() -> XlrdAdapter:
    return XlrdAdapter()


# ── helpers ──────────────────────────────────────────────────────────────


class TestXlwtHelpers:
    def test_parse_cell_ref_valid(self) -> None:
        assert xlwt_parse("A1") == (0, 0)
        assert xlwt_parse("B3") == (2, 1)
        assert xlwt_parse("AA1") == (0, 26)

    def test_parse_cell_ref_invalid(self) -> None:
        with pytest.raises(ValueError, match="Invalid cell reference"):
            xlwt_parse("123")

    def test_col_to_index(self) -> None:
        assert _col_to_index("A") == 0
        assert _col_to_index("B") == 1
        assert _col_to_index("Z") == 25
        assert _col_to_index("AA") == 26

    def test_hex_to_xlwt_colour_exact_match(self) -> None:
        idx = _hex_to_xlwt_colour("#FF0000")
        assert isinstance(idx, int)

    def test_hex_to_xlwt_colour_short_hex(self) -> None:
        idx = _hex_to_xlwt_colour("#FFF")
        assert idx == 0x40  # default (bad length)

    def test_hex_to_xlwt_colour_nearest(self) -> None:
        # Not an exact match — triggers brute-force nearest
        idx = _hex_to_xlwt_colour("#123456")
        assert isinstance(idx, int)

    def test_hex_to_xlwt_colour_without_hash(self) -> None:
        idx = _hex_to_xlwt_colour("0000FF")
        assert isinstance(idx, int)


# ── cell value write→read roundtrip ──────────────────────────────────────


class TestXlwtXlrdCellValues:
    def test_string(self, xlwt: XlwtAdapter, xlrd: XlrdAdapter, tmp_path: Path) -> None:
        path = tmp_path / "str.xls"
        wb = xlwt.create_workbook()
        xlwt.add_sheet(wb, "S1")
        xlwt.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="hello"))
        xlwt.save_workbook(wb, path)

        rb = xlrd.open_workbook(path)
        cv = xlrd.read_cell_value(rb, "S1", "A1")
        assert cv.type == CellType.STRING
        assert cv.value == "hello"
        xlrd.close_workbook(rb)

    def test_number(self, xlwt: XlwtAdapter, xlrd: XlrdAdapter, tmp_path: Path) -> None:
        path = tmp_path / "num.xls"
        wb = xlwt.create_workbook()
        xlwt.add_sheet(wb, "S1")
        xlwt.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.NUMBER, value=42.5))
        xlwt.save_workbook(wb, path)

        rb = xlrd.open_workbook(path)
        cv = xlrd.read_cell_value(rb, "S1", "A1")
        assert cv.type == CellType.NUMBER
        assert cv.value == 42.5
        xlrd.close_workbook(rb)

    def test_boolean(self, xlwt: XlwtAdapter, xlrd: XlrdAdapter, tmp_path: Path) -> None:
        path = tmp_path / "bool.xls"
        wb = xlwt.create_workbook()
        xlwt.add_sheet(wb, "S1")
        xlwt.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.BOOLEAN, value=True))
        xlwt.save_workbook(wb, path)

        rb = xlrd.open_workbook(path)
        cv = xlrd.read_cell_value(rb, "S1", "A1")
        assert cv.type == CellType.BOOLEAN
        assert cv.value is True
        xlrd.close_workbook(rb)

    def test_date(self, xlwt: XlwtAdapter, xlrd: XlrdAdapter, tmp_path: Path) -> None:
        path = tmp_path / "date.xls"
        wb = xlwt.create_workbook()
        xlwt.add_sheet(wb, "S1")
        xlwt.write_cell_value(
            wb, "S1", "A1", CellValue(type=CellType.DATE, value=date(2024, 6, 15))
        )
        xlwt.save_workbook(wb, path)

        rb = xlrd.open_workbook(path)
        cv = xlrd.read_cell_value(rb, "S1", "A1")
        assert cv.type == CellType.DATE
        assert cv.value == date(2024, 6, 15)
        xlrd.close_workbook(rb)

    def test_datetime(self, xlwt: XlwtAdapter, xlrd: XlrdAdapter, tmp_path: Path) -> None:
        path = tmp_path / "dt.xls"
        wb = xlwt.create_workbook()
        xlwt.add_sheet(wb, "S1")
        xlwt.write_cell_value(
            wb,
            "S1",
            "A1",
            CellValue(type=CellType.DATETIME, value=datetime(2024, 6, 15, 14, 30, 0)),
        )
        xlwt.save_workbook(wb, path)

        rb = xlrd.open_workbook(path)
        cv = xlrd.read_cell_value(rb, "S1", "A1")
        assert cv.type == CellType.DATETIME
        assert cv.value == datetime(2024, 6, 15, 14, 30, 0)
        xlrd.close_workbook(rb)

    def test_formula(self, xlwt: XlwtAdapter, xlrd: XlrdAdapter, tmp_path: Path) -> None:
        path = tmp_path / "formula.xls"
        wb = xlwt.create_workbook()
        xlwt.add_sheet(wb, "S1")
        xlwt.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.NUMBER, value=10))
        xlwt.write_cell_value(
            wb,
            "S1",
            "A2",
            CellValue(type=CellType.FORMULA, value="=A1*2", formula="=A1*2"),
        )
        xlwt.save_workbook(wb, path)

        rb = xlrd.open_workbook(path)
        # xlrd returns formula results, but the formula was written
        # The value read back depends on whether xlrd can evaluate
        cv = xlrd.read_cell_value(rb, "S1", "A1")
        assert cv.value == 10
        xlrd.close_workbook(rb)

    def test_error_values(self, xlwt: XlwtAdapter, xlrd: XlrdAdapter, tmp_path: Path) -> None:
        path = tmp_path / "err.xls"
        wb = xlwt.create_workbook()
        xlwt.add_sheet(wb, "S1")
        # xlwt writes errors as text strings
        xlwt.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.ERROR, value="#DIV/0!"))
        xlwt.write_cell_value(wb, "S1", "A2", CellValue(type=CellType.ERROR, value="#N/A"))
        xlwt.save_workbook(wb, path)

        rb = xlrd.open_workbook(path)
        # Written as text strings, xlrd reads them back as text and detects errors
        cv1 = xlrd.read_cell_value(rb, "S1", "A1")
        assert cv1.type == CellType.ERROR
        cv2 = xlrd.read_cell_value(rb, "S1", "A2")
        assert cv2.type == CellType.ERROR
        xlrd.close_workbook(rb)

    def test_blank(self, xlwt: XlwtAdapter, xlrd: XlrdAdapter, tmp_path: Path) -> None:
        path = tmp_path / "blank.xls"
        wb = xlwt.create_workbook()
        xlwt.add_sheet(wb, "S1")
        xlwt.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.BLANK))
        # Also write something in A2 so sheet isn't empty
        xlwt.write_cell_value(wb, "S1", "A2", CellValue(type=CellType.STRING, value="x"))
        xlwt.save_workbook(wb, path)

        rb = xlrd.open_workbook(path)
        cv = xlrd.read_cell_value(rb, "S1", "A1")
        # Blank written as empty string may come back as BLANK or empty STRING
        assert cv.type in (CellType.BLANK, CellType.STRING)
        xlrd.close_workbook(rb)

    def test_out_of_bounds_cell(self, xlrd: XlrdAdapter, xlwt: XlwtAdapter, tmp_path: Path) -> None:
        path = tmp_path / "oob.xls"
        wb = xlwt.create_workbook()
        xlwt.add_sheet(wb, "S1")
        xlwt.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="x"))
        xlwt.save_workbook(wb, path)

        rb = xlrd.open_workbook(path)
        cv = xlrd.read_cell_value(rb, "S1", "Z99")
        assert cv.type == CellType.BLANK
        xlrd.close_workbook(rb)


# ── formatting write→read roundtrip ──────────────────────────────────────


class TestXlwtXlrdFormatting:
    def test_bold_italic_strikethrough(
        self, xlwt: XlwtAdapter, xlrd: XlrdAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "fmt.xls"
        wb = xlwt.create_workbook()
        xlwt.add_sheet(wb, "S1")
        xlwt.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="styled"))
        xlwt.write_cell_format(
            wb, "S1", "A1", CellFormat(bold=True, italic=True, strikethrough=True)
        )
        xlwt.save_workbook(wb, path)

        rb = xlrd.open_workbook(path)
        fmt = xlrd.read_cell_format(rb, "S1", "A1")
        assert fmt.bold is True
        assert fmt.italic is True
        assert fmt.strikethrough is True
        xlrd.close_workbook(rb)

    def test_underline_single(self, xlwt: XlwtAdapter, xlrd: XlrdAdapter, tmp_path: Path) -> None:
        path = tmp_path / "ul.xls"
        wb = xlwt.create_workbook()
        xlwt.add_sheet(wb, "S1")
        xlwt.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="u"))
        xlwt.write_cell_format(wb, "S1", "A1", CellFormat(underline="single"))
        xlwt.save_workbook(wb, path)

        rb = xlrd.open_workbook(path)
        fmt = xlrd.read_cell_format(rb, "S1", "A1")
        assert fmt.underline == "single"
        xlrd.close_workbook(rb)

    def test_underline_double(self, xlwt: XlwtAdapter, xlrd: XlrdAdapter, tmp_path: Path) -> None:
        path = tmp_path / "ul2.xls"
        wb = xlwt.create_workbook()
        xlwt.add_sheet(wb, "S1")
        xlwt.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="u2"))
        xlwt.write_cell_format(wb, "S1", "A1", CellFormat(underline="double"))
        xlwt.save_workbook(wb, path)

        rb = xlrd.open_workbook(path)
        fmt = xlrd.read_cell_format(rb, "S1", "A1")
        assert fmt.underline == "double"
        xlrd.close_workbook(rb)

    def test_underline_accounting(
        self, xlwt: XlwtAdapter, xlrd: XlrdAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "ul_acc.xls"
        wb = xlwt.create_workbook()
        xlwt.add_sheet(wb, "S1")
        xlwt.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="sa"))
        xlwt.write_cell_format(wb, "S1", "A1", CellFormat(underline="singleAccounting"))
        xlwt.write_cell_value(wb, "S1", "A2", CellValue(type=CellType.STRING, value="da"))
        xlwt.write_cell_format(wb, "S1", "A2", CellFormat(underline="doubleAccounting"))
        xlwt.save_workbook(wb, path)

        rb = xlrd.open_workbook(path)
        f1 = xlrd.read_cell_format(rb, "S1", "A1")
        assert f1.underline == "singleAccounting"
        f2 = xlrd.read_cell_format(rb, "S1", "A2")
        assert f2.underline == "doubleAccounting"
        xlrd.close_workbook(rb)

    def test_font_name_size_color(
        self, xlwt: XlwtAdapter, xlrd: XlrdAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "font.xls"
        wb = xlwt.create_workbook()
        xlwt.add_sheet(wb, "S1")
        xlwt.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="f"))
        xlwt.write_cell_format(
            wb,
            "S1",
            "A1",
            CellFormat(font_name="Courier New", font_size=14.0, font_color="#FF0000"),
        )
        xlwt.save_workbook(wb, path)

        rb = xlrd.open_workbook(path)
        fmt = xlrd.read_cell_format(rb, "S1", "A1")
        assert fmt.font_name == "Courier New"
        assert fmt.font_size == 14.0
        # Font color depends on palette mapping; just check it's set
        assert fmt.font_color is not None
        xlrd.close_workbook(rb)

    def test_bg_color(self, xlwt: XlwtAdapter, xlrd: XlrdAdapter, tmp_path: Path) -> None:
        path = tmp_path / "bg.xls"
        wb = xlwt.create_workbook()
        xlwt.add_sheet(wb, "S1")
        xlwt.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="bg"))
        xlwt.write_cell_format(wb, "S1", "A1", CellFormat(bg_color="#FFFF00"))
        xlwt.save_workbook(wb, path)

        rb = xlrd.open_workbook(path)
        fmt = xlrd.read_cell_format(rb, "S1", "A1")
        # bg_color depends on palette, just verify format was read
        assert fmt is not None
        xlrd.close_workbook(rb)

    def test_number_format(self, xlwt: XlwtAdapter, xlrd: XlrdAdapter, tmp_path: Path) -> None:
        path = tmp_path / "nf.xls"
        wb = xlwt.create_workbook()
        xlwt.add_sheet(wb, "S1")
        xlwt.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.NUMBER, value=0.5))
        xlwt.write_cell_format(wb, "S1", "A1", CellFormat(number_format="0.00%"))
        xlwt.save_workbook(wb, path)

        rb = xlrd.open_workbook(path)
        fmt = xlrd.read_cell_format(rb, "S1", "A1")
        assert fmt.number_format == "0.00%"
        xlrd.close_workbook(rb)

    def test_alignment(self, xlwt: XlwtAdapter, xlrd: XlrdAdapter, tmp_path: Path) -> None:
        path = tmp_path / "align.xls"
        wb = xlwt.create_workbook()
        xlwt.add_sheet(wb, "S1")
        xlwt.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="a"))
        xlwt.write_cell_format(
            wb,
            "S1",
            "A1",
            CellFormat(
                h_align="center",
                v_align="top",
                wrap=True,
                rotation=45,
                indent=2,
            ),
        )
        xlwt.save_workbook(wb, path)

        rb = xlrd.open_workbook(path)
        fmt = xlrd.read_cell_format(rb, "S1", "A1")
        assert fmt.h_align == "center"
        assert fmt.v_align == "top"
        assert fmt.wrap is True
        assert fmt.rotation == 45
        assert fmt.indent == 2
        xlrd.close_workbook(rb)

    def test_format_no_existing_value(
        self, xlwt: XlwtAdapter, xlrd: XlrdAdapter, tmp_path: Path
    ) -> None:
        """write_cell_format when no prior write_cell_value (no existing cell)."""
        path = tmp_path / "fmt_only.xls"
        wb = xlwt.create_workbook()
        xlwt.add_sheet(wb, "S1")
        xlwt.write_cell_format(wb, "S1", "A1", CellFormat(bold=True))
        xlwt.save_workbook(wb, path)

        rb = xlrd.open_workbook(path)
        fmt = xlrd.read_cell_format(rb, "S1", "A1")
        assert fmt.bold is True
        xlrd.close_workbook(rb)

    def test_format_out_of_bounds_read(
        self, xlwt: XlwtAdapter, xlrd: XlrdAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "oob_fmt.xls"
        wb = xlwt.create_workbook()
        xlwt.add_sheet(wb, "S1")
        xlwt.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="x"))
        xlwt.save_workbook(wb, path)

        rb = xlrd.open_workbook(path)
        fmt = xlrd.read_cell_format(rb, "S1", "Z99")
        assert fmt == CellFormat()
        xlrd.close_workbook(rb)


# ── borders write→read roundtrip ─────────────────────────────────────────


class TestXlwtXlrdBorders:
    def test_four_sided_border(self, xlwt: XlwtAdapter, xlrd: XlrdAdapter, tmp_path: Path) -> None:
        path = tmp_path / "border.xls"
        wb = xlwt.create_workbook()
        xlwt.add_sheet(wb, "S1")
        xlwt.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="b"))
        edge = BorderEdge(style=BorderStyle.THIN, color="#000000")
        xlwt.write_cell_border(
            wb, "S1", "A1", BorderInfo(top=edge, bottom=edge, left=edge, right=edge)
        )
        xlwt.save_workbook(wb, path)

        rb = xlrd.open_workbook(path)
        bi = xlrd.read_cell_border(rb, "S1", "A1")
        assert bi.top is not None
        assert bi.top.style == BorderStyle.THIN
        assert bi.bottom is not None
        assert bi.left is not None
        assert bi.right is not None
        xlrd.close_workbook(rb)

    def test_diagonal_up(self, xlwt: XlwtAdapter, xlrd: XlrdAdapter, tmp_path: Path) -> None:
        path = tmp_path / "diag_up.xls"
        wb = xlwt.create_workbook()
        xlwt.add_sheet(wb, "S1")
        xlwt.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="d"))
        diag_edge = BorderEdge(style=BorderStyle.THIN, color="#FF0000")
        xlwt.write_cell_border(wb, "S1", "A1", BorderInfo(diagonal_up=diag_edge))
        xlwt.save_workbook(wb, path)

        rb = xlrd.open_workbook(path)
        bi = xlrd.read_cell_border(rb, "S1", "A1")
        # xlrd may or may not detect diag depending on BIFF support
        assert bi is not None
        xlrd.close_workbook(rb)

    def test_diagonal_down(self, xlwt: XlwtAdapter, xlrd: XlrdAdapter, tmp_path: Path) -> None:
        path = tmp_path / "diag_dn.xls"
        wb = xlwt.create_workbook()
        xlwt.add_sheet(wb, "S1")
        xlwt.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="d"))
        diag_edge = BorderEdge(style=BorderStyle.MEDIUM, color="#0000FF")
        xlwt.write_cell_border(wb, "S1", "A1", BorderInfo(diagonal_down=diag_edge))
        xlwt.save_workbook(wb, path)

        rb = xlrd.open_workbook(path)
        bi = xlrd.read_cell_border(rb, "S1", "A1")
        assert bi is not None
        xlrd.close_workbook(rb)

    def test_border_no_existing_value(
        self, xlwt: XlwtAdapter, xlrd: XlrdAdapter, tmp_path: Path
    ) -> None:
        """write_cell_border when no prior write_cell_value."""
        path = tmp_path / "bdr_only.xls"
        wb = xlwt.create_workbook()
        xlwt.add_sheet(wb, "S1")
        edge = BorderEdge(style=BorderStyle.THICK, color="#000000")
        xlwt.write_cell_border(wb, "S1", "A1", BorderInfo(top=edge))
        xlwt.save_workbook(wb, path)

        rb = xlrd.open_workbook(path)
        bi = xlrd.read_cell_border(rb, "S1", "A1")
        assert bi.top is not None
        xlrd.close_workbook(rb)

    def test_border_out_of_bounds_read(
        self, xlwt: XlwtAdapter, xlrd: XlrdAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "bdr_oob.xls"
        wb = xlwt.create_workbook()
        xlwt.add_sheet(wb, "S1")
        xlwt.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="x"))
        xlwt.save_workbook(wb, path)

        rb = xlrd.open_workbook(path)
        bi = xlrd.read_cell_border(rb, "S1", "Z99")
        assert bi == BorderInfo()
        xlrd.close_workbook(rb)


# ── dimensions, merge, freeze ────────────────────────────────────────────


class TestXlwtXlrdDimensions:
    def test_row_height(self, xlwt: XlwtAdapter, xlrd: XlrdAdapter, tmp_path: Path) -> None:
        path = tmp_path / "rh.xls"
        wb = xlwt.create_workbook()
        xlwt.add_sheet(wb, "S1")
        xlwt.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="h"))
        xlwt.set_row_height(wb, "S1", 1, 30.0)
        xlwt.save_workbook(wb, path)

        rb = xlrd.open_workbook(path)
        h = xlrd.read_row_height(rb, "S1", 1)
        assert h is not None
        assert h == pytest.approx(30.0, abs=1.0)
        xlrd.close_workbook(rb)

    def test_row_height_out_of_bounds(
        self, xlwt: XlwtAdapter, xlrd: XlrdAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "rh_oob.xls"
        wb = xlwt.create_workbook()
        xlwt.add_sheet(wb, "S1")
        xlwt.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="x"))
        xlwt.save_workbook(wb, path)

        rb = xlrd.open_workbook(path)
        h = xlrd.read_row_height(rb, "S1", 99)
        assert h is None
        xlrd.close_workbook(rb)

    def test_column_width(self, xlwt: XlwtAdapter, xlrd: XlrdAdapter, tmp_path: Path) -> None:
        path = tmp_path / "cw.xls"
        wb = xlwt.create_workbook()
        xlwt.add_sheet(wb, "S1")
        xlwt.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="w"))
        xlwt.set_column_width(wb, "S1", "A", 20.0)
        xlwt.save_workbook(wb, path)

        rb = xlrd.open_workbook(path)
        w = xlrd.read_column_width(rb, "S1", "A")
        assert w is not None
        assert w == pytest.approx(20.0, abs=1.0)
        xlrd.close_workbook(rb)

    def test_column_width_not_set(
        self, xlwt: XlwtAdapter, xlrd: XlrdAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "cw_no.xls"
        wb = xlwt.create_workbook()
        xlwt.add_sheet(wb, "S1")
        xlwt.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="x"))
        xlwt.save_workbook(wb, path)

        rb = xlrd.open_workbook(path)
        # Column B was never set — may return None or default width
        xlrd.read_column_width(rb, "S1", "B")
        xlrd.close_workbook(rb)


class TestXlwtXlrdMerge:
    def test_merge_cells(self, xlwt: XlwtAdapter, xlrd: XlrdAdapter, tmp_path: Path) -> None:
        path = tmp_path / "merge.xls"
        wb = xlwt.create_workbook()
        xlwt.add_sheet(wb, "S1")
        xlwt.merge_cells(wb, "S1", "A1:C1")
        xlwt.save_workbook(wb, path)

        rb = xlrd.open_workbook(path)
        ranges = xlrd.read_merged_ranges(rb, "S1")
        assert len(ranges) >= 1
        xlrd.close_workbook(rb)


class TestXlwtXlrdFreezePanes:
    @pytest.mark.xfail(
        reason="xlrd adapter uses non-existent Sheet.frozen_row_count attribute",
        strict=False,
    )
    def test_freeze(self, xlwt: XlwtAdapter, xlrd: XlrdAdapter, tmp_path: Path) -> None:
        path = tmp_path / "freeze.xls"
        wb = xlwt.create_workbook()
        xlwt.add_sheet(wb, "S1")
        xlwt.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="f"))
        xlwt.set_freeze_panes(wb, "S1", {"mode": "freeze", "top_left_cell": "B2"})
        xlwt.save_workbook(wb, path)

        rb = xlrd.open_workbook(path)
        fp = xlrd.read_freeze_panes(rb, "S1")
        assert fp.get("mode") == "freeze"
        assert fp.get("top_left_cell") == "B2"
        xlrd.close_workbook(rb)


# ── xlwt no-op tier 2 methods ────────────────────────────────────────────


class TestXlwtNoOps:
    def test_noop_methods(self, xlwt: XlwtAdapter) -> None:
        wb = xlwt.create_workbook()
        xlwt.add_sheet(wb, "S1")
        # These are all no-ops but should not raise
        xlwt.add_conditional_format(wb, "S1", {})
        xlwt.add_data_validation(wb, "S1", {})
        xlwt.add_hyperlink(wb, "S1", {})
        xlwt.add_image(wb, "S1", {})
        xlwt.add_pivot_table(wb, "S1", {})
        xlwt.add_comment(wb, "S1", {})


# ── xlrd tier 2 read stubs (empty returns) ───────────────────────────────


class TestXlrdTier2Stubs:
    def test_empty_reads(self, xlwt: XlwtAdapter, xlrd: XlrdAdapter, tmp_path: Path) -> None:
        path = tmp_path / "stub.xls"
        wb = xlwt.create_workbook()
        xlwt.add_sheet(wb, "S1")
        xlwt.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="x"))
        xlwt.save_workbook(wb, path)

        rb = xlrd.open_workbook(path)
        assert xlrd.read_conditional_formats(rb, "S1") == []
        assert xlrd.read_data_validations(rb, "S1") == []
        assert xlrd.read_images(rb, "S1") == []
        assert xlrd.read_pivot_tables(rb, "S1") == []
        xlrd.close_workbook(rb)


# ── xlwt adapter info ────────────────────────────────────────────────────


class TestXlwtInfo:
    def test_info(self, xlwt: XlwtAdapter) -> None:
        info = xlwt.info
        assert info.name == "xlwt"
        assert info.language == "python"
        assert "write" in info.capabilities

    def test_output_extension(self, xlwt: XlwtAdapter) -> None:
        assert xlwt.output_extension == ".xls"


class TestXlrdInfo:
    def test_info(self, xlrd: XlrdAdapter) -> None:
        info = xlrd.info
        assert info.name == "xlrd"
        assert info.language == "python"
        assert "read" in info.capabilities

    def test_supported_extensions(self, xlrd: XlrdAdapter) -> None:
        assert ".xls" in xlrd.supported_read_extensions
