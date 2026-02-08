"""Tests for PandasAdapter (read+write, value-only via pd.read_excel / pd.ExcelWriter)."""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

import pytest

from excelbench.harness.adapters.openpyxl_adapter import OpenpyxlAdapter
from excelbench.harness.adapters.pandas_adapter import PandasAdapter
from excelbench.models import (
    BorderInfo,
    CellFormat,
    CellType,
    CellValue,
)


@pytest.fixture
def opxl() -> OpenpyxlAdapter:
    return OpenpyxlAdapter()


@pytest.fixture
def pdxl() -> PandasAdapter:
    return PandasAdapter()


def _write_openpyxl_fixture(opxl: OpenpyxlAdapter, path: Path) -> None:
    """Write a multi-type .xlsx fixture using openpyxl for read tests."""
    wb = opxl.create_workbook()
    opxl.add_sheet(wb, "S1")
    opxl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="hello"))
    opxl.write_cell_value(wb, "S1", "A2", CellValue(type=CellType.NUMBER, value=42.5))
    opxl.write_cell_value(wb, "S1", "A3", CellValue(type=CellType.BOOLEAN, value=True))
    opxl.write_cell_value(wb, "S1", "A4", CellValue(type=CellType.DATE, value=date(2024, 6, 15)))
    opxl.write_cell_value(
        wb,
        "S1",
        "A5",
        CellValue(type=CellType.DATETIME, value=datetime(2024, 6, 15, 14, 30, 0)),
    )
    opxl.write_cell_value(wb, "S1", "A6", CellValue(type=CellType.ERROR, value="#N/A"))
    opxl.write_cell_value(
        wb, "S1", "A7", CellValue(type=CellType.FORMULA, value="=1+1", formula="=1+1")
    )
    opxl.write_cell_value(wb, "S1", "A8", CellValue(type=CellType.BLANK))
    opxl.save_workbook(wb, path)


# ═════════════════════════════════════════════════════════════════════════
# TestPandasInfo
# ═════════════════════════════════════════════════════════════════════════


class TestPandasInfo:
    def test_name(self, pdxl: PandasAdapter) -> None:
        assert pdxl.info.name == "pandas"

    def test_version(self, pdxl: PandasAdapter) -> None:
        assert pdxl.info.version != "unknown"

    def test_capabilities(self, pdxl: PandasAdapter) -> None:
        assert "read" in pdxl.info.capabilities
        assert "write" in pdxl.info.capabilities

    def test_language(self, pdxl: PandasAdapter) -> None:
        assert pdxl.info.language == "python"


# ═════════════════════════════════════════════════════════════════════════
# TestPandasReadCellValue
# ═════════════════════════════════════════════════════════════════════════


class TestPandasReadCellValue:
    """Read via pandas from openpyxl-written fixtures."""

    def test_string(self, pdxl: PandasAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        wb = pdxl.open_workbook(path)
        cv = pdxl.read_cell_value(wb, "S1", "A1")
        assert cv.type == CellType.STRING
        assert cv.value == "hello"
        pdxl.close_workbook(wb)

    def test_number(self, pdxl: PandasAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        wb = pdxl.open_workbook(path)
        cv = pdxl.read_cell_value(wb, "S1", "A2")
        assert cv.type == CellType.NUMBER
        assert cv.value == 42.5
        pdxl.close_workbook(wb)

    def test_boolean(self, pdxl: PandasAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        wb = pdxl.open_workbook(path)
        cv = pdxl.read_cell_value(wb, "S1", "A3")
        assert cv.type == CellType.BOOLEAN
        assert cv.value is True
        pdxl.close_workbook(wb)

    def test_date(self, pdxl: PandasAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        wb = pdxl.open_workbook(path)
        cv = pdxl.read_cell_value(wb, "S1", "A4")
        assert cv.type == CellType.DATE
        assert cv.value == date(2024, 6, 15)
        pdxl.close_workbook(wb)

    def test_datetime(self, pdxl: PandasAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        wb = pdxl.open_workbook(path)
        cv = pdxl.read_cell_value(wb, "S1", "A5")
        assert cv.type == CellType.DATETIME
        pdxl.close_workbook(wb)

    def test_error(self, pdxl: PandasAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        wb = pdxl.open_workbook(path)
        cv = pdxl.read_cell_value(wb, "S1", "A6")
        # pandas may return error as string or blank (openpyxl error type not round-tripped)
        assert cv.type in (CellType.ERROR, CellType.STRING, CellType.BLANK)
        pdxl.close_workbook(wb)

    def test_blank_out_of_bounds(
        self, pdxl: PandasAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        wb = pdxl.open_workbook(path)
        # Row out of bounds
        cv = pdxl.read_cell_value(wb, "S1", "A99")
        assert cv.type == CellType.BLANK
        # Col out of bounds
        cv2 = pdxl.read_cell_value(wb, "S1", "Z1")
        assert cv2.type == CellType.BLANK
        pdxl.close_workbook(wb)

    def test_sheet_not_found(
        self, pdxl: PandasAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        wb = pdxl.open_workbook(path)
        cv = pdxl.read_cell_value(wb, "NoSheet", "A1")
        assert cv.type == CellType.BLANK
        pdxl.close_workbook(wb)

    def test_sheet_names(
        self, pdxl: PandasAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        wb = pdxl.open_workbook(path)
        names = pdxl.get_sheet_names(wb)
        assert "S1" in names
        pdxl.close_workbook(wb)


# ═════════════════════════════════════════════════════════════════════════
# TestPandasReadStubs
# ═════════════════════════════════════════════════════════════════════════


class TestPandasReadStubs:
    """Tier-2 reads all return empty."""

    def test_all_stubs(self, pdxl: PandasAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        wb = pdxl.open_workbook(path)
        assert pdxl.read_cell_format(wb, "S1", "A1") == CellFormat()
        assert pdxl.read_cell_border(wb, "S1", "A1") == BorderInfo()
        assert pdxl.read_merged_ranges(wb, "S1") == []
        assert pdxl.read_conditional_formats(wb, "S1") == []
        assert pdxl.read_data_validations(wb, "S1") == []
        assert pdxl.read_hyperlinks(wb, "S1") == []
        assert pdxl.read_images(wb, "S1") == []
        assert pdxl.read_pivot_tables(wb, "S1") == []
        assert pdxl.read_comments(wb, "S1") == []
        assert pdxl.read_freeze_panes(wb, "S1") == {}
        assert pdxl.read_row_height(wb, "S1", 1) is None
        assert pdxl.read_column_width(wb, "S1", "A") is None
        pdxl.close_workbook(wb)


# ═════════════════════════════════════════════════════════════════════════
# TestPandasWriteRoundtrip
# ═════════════════════════════════════════════════════════════════════════


class TestPandasWriteRoundtrip:
    """Write via pandas, read back via openpyxl."""

    def test_string(self, pdxl: PandasAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "pd_out.xlsx"
        wb = pdxl.create_workbook()
        pdxl.add_sheet(wb, "S1")
        pdxl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="hi"))
        pdxl.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        cv = opxl.read_cell_value(rb, "S1", "A1")
        assert cv.value == "hi"
        opxl.close_workbook(rb)

    def test_number(self, pdxl: PandasAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "pd_num.xlsx"
        wb = pdxl.create_workbook()
        pdxl.add_sheet(wb, "S1")
        pdxl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.NUMBER, value=99))
        pdxl.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        cv = opxl.read_cell_value(rb, "S1", "A1")
        assert cv.value == 99
        opxl.close_workbook(rb)

    def test_boolean(self, pdxl: PandasAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "pd_bool.xlsx"
        wb = pdxl.create_workbook()
        pdxl.add_sheet(wb, "S1")
        pdxl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.BOOLEAN, value=False))
        pdxl.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        cv = opxl.read_cell_value(rb, "S1", "A1")
        assert cv.value is False
        opxl.close_workbook(rb)

    def test_blank(self, pdxl: PandasAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "pd_blank.xlsx"
        wb = pdxl.create_workbook()
        pdxl.add_sheet(wb, "S1")
        pdxl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.BLANK))
        pdxl.write_cell_value(wb, "S1", "A2", CellValue(type=CellType.STRING, value="x"))
        pdxl.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        cv = opxl.read_cell_value(rb, "S1", "A1")
        # Blank may come through as BLANK or STRING("") depending on pandas/openpyxl
        assert cv.type in (CellType.BLANK, CellType.STRING)
        opxl.close_workbook(rb)

    def test_date(self, pdxl: PandasAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "pd_date.xlsx"
        wb = pdxl.create_workbook()
        pdxl.add_sheet(wb, "S1")
        pdxl.write_cell_value(
            wb, "S1", "A1", CellValue(type=CellType.DATE, value=date(2024, 1, 1))
        )
        pdxl.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        cv = opxl.read_cell_value(rb, "S1", "A1")
        assert cv.type in (CellType.DATE, CellType.DATETIME)
        opxl.close_workbook(rb)

    def test_datetime(self, pdxl: PandasAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "pd_dt.xlsx"
        wb = pdxl.create_workbook()
        pdxl.add_sheet(wb, "S1")
        pdxl.write_cell_value(
            wb,
            "S1",
            "A1",
            CellValue(type=CellType.DATETIME, value=datetime(2024, 1, 1, 12, 0, 0)),
        )
        pdxl.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        cv = opxl.read_cell_value(rb, "S1", "A1")
        assert cv.type == CellType.DATETIME
        opxl.close_workbook(rb)

    def test_error(self, pdxl: PandasAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "pd_err.xlsx"
        wb = pdxl.create_workbook()
        pdxl.add_sheet(wb, "S1")
        pdxl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.ERROR, value="#VALUE!"))
        pdxl.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        cv = opxl.read_cell_value(rb, "S1", "A1")
        # Error strings written through pandas come back as strings
        assert cv.type in (CellType.ERROR, CellType.STRING)
        opxl.close_workbook(rb)

    def test_formula(self, pdxl: PandasAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "pd_form.xlsx"
        wb = pdxl.create_workbook()
        pdxl.add_sheet(wb, "S1")
        pdxl.write_cell_value(
            wb,
            "S1",
            "A1",
            CellValue(type=CellType.FORMULA, value="=1+1", formula="=1+1"),
        )
        pdxl.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        cv = opxl.read_cell_value(rb, "S1", "A1")
        # pandas writes formula as string — openpyxl may see it as STRING or FORMULA
        assert cv.type in (CellType.FORMULA, CellType.STRING)
        opxl.close_workbook(rb)

    def test_empty_sheet(
        self, pdxl: PandasAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "pd_empty.xlsx"
        wb = pdxl.create_workbook()
        pdxl.add_sheet(wb, "EmptySheet")
        pdxl.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        assert "EmptySheet" in opxl.get_sheet_names(rb)
        opxl.close_workbook(rb)

    def test_write_auto_creates_sheet(self, pdxl: PandasAdapter) -> None:
        wb = pdxl.create_workbook()
        # Don't call add_sheet — write_cell_value should auto-create
        pdxl.write_cell_value(wb, "Auto", "A1", CellValue(type=CellType.STRING, value="x"))
        assert "Auto" in wb["sheets"]


# ═════════════════════════════════════════════════════════════════════════
# TestPandasWriteNoops
# ═════════════════════════════════════════════════════════════════════════


class TestPandasWriteNoops:
    """Formatting/tier-2 writes don't raise."""

    def test_noop_methods(self, pdxl: PandasAdapter) -> None:
        wb = pdxl.create_workbook()
        pdxl.add_sheet(wb, "S1")
        # All no-ops — should not raise
        pdxl.write_cell_format(wb, "S1", "A1", CellFormat())
        pdxl.write_cell_border(wb, "S1", "A1", BorderInfo())
        pdxl.set_row_height(wb, "S1", 1, 30.0)
        pdxl.set_column_width(wb, "S1", "A", 20.0)
        pdxl.merge_cells(wb, "S1", "A1:B2")
        pdxl.add_conditional_format(wb, "S1", {})
        pdxl.add_data_validation(wb, "S1", {})
        pdxl.add_hyperlink(wb, "S1", {})
        pdxl.add_image(wb, "S1", {})
        pdxl.add_pivot_table(wb, "S1", {})
        pdxl.add_comment(wb, "S1", {})
        pdxl.set_freeze_panes(wb, "S1", {})
