"""Tests for PolarsAdapter (read-only, value-only via pl.read_excel)."""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

import pytest

from excelbench.harness.adapters.openpyxl_adapter import OpenpyxlAdapter
from excelbench.harness.adapters.polars_adapter import PolarsAdapter
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
def plxl() -> PolarsAdapter:
    return PolarsAdapter()


def _write_fixture(opxl: OpenpyxlAdapter, path: Path) -> None:
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
# TestPolarsInfo
# ═════════════════════════════════════════════════════════════════════════


class TestPolarsInfo:
    def test_name(self, plxl: PolarsAdapter) -> None:
        assert plxl.info.name == "polars"

    def test_version(self, plxl: PolarsAdapter) -> None:
        assert plxl.info.version != "unknown"

    def test_capabilities(self, plxl: PolarsAdapter) -> None:
        assert plxl.info.capabilities == {"read"}

    def test_is_read_only(self, plxl: PolarsAdapter) -> None:
        assert plxl.can_read()
        assert not plxl.can_write()


# ═════════════════════════════════════════════════════════════════════════
# TestPolarsReadCellValue
# ═════════════════════════════════════════════════════════════════════════


class TestPolarsReadCellValue:
    """Read via polars from openpyxl-written fixtures."""

    def test_string(self, plxl: PolarsAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_fixture(opxl, path)
        wb = plxl.open_workbook(path)
        cv = plxl.read_cell_value(wb, "S1", "A1")
        assert cv.type == CellType.STRING
        assert cv.value == "hello"
        plxl.close_workbook(wb)

    def test_number(self, plxl: PolarsAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_fixture(opxl, path)
        wb = plxl.open_workbook(path)
        cv = plxl.read_cell_value(wb, "S1", "A2")
        assert cv.type == CellType.NUMBER
        assert cv.value == 42.5
        plxl.close_workbook(wb)

    def test_boolean(self, plxl: PolarsAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_fixture(opxl, path)
        wb = plxl.open_workbook(path)
        cv = plxl.read_cell_value(wb, "S1", "A3")
        assert cv.type == CellType.BOOLEAN
        assert cv.value is True
        plxl.close_workbook(wb)

    def test_date(self, plxl: PolarsAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_fixture(opxl, path)
        wb = plxl.open_workbook(path)
        cv = plxl.read_cell_value(wb, "S1", "A4")
        assert cv.type == CellType.DATE
        assert cv.value == date(2024, 6, 15)
        plxl.close_workbook(wb)

    def test_datetime(self, plxl: PolarsAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_fixture(opxl, path)
        wb = plxl.open_workbook(path)
        cv = plxl.read_cell_value(wb, "S1", "A5")
        assert cv.type == CellType.DATETIME
        plxl.close_workbook(wb)

    def test_error(self, plxl: PolarsAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_fixture(opxl, path)
        wb = plxl.open_workbook(path)
        cv = plxl.read_cell_value(wb, "S1", "A6")
        # polars may see the error formula as a formula string
        assert cv.type in (CellType.ERROR, CellType.STRING, CellType.FORMULA, CellType.BLANK)
        plxl.close_workbook(wb)

    def test_blank_out_of_bounds(
        self, plxl: PolarsAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_fixture(opxl, path)
        wb = plxl.open_workbook(path)
        cv = plxl.read_cell_value(wb, "S1", "A99")
        assert cv.type == CellType.BLANK
        cv2 = plxl.read_cell_value(wb, "S1", "Z1")
        assert cv2.type == CellType.BLANK
        plxl.close_workbook(wb)

    def test_sheet_not_found(
        self, plxl: PolarsAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_fixture(opxl, path)
        wb = plxl.open_workbook(path)
        cv = plxl.read_cell_value(wb, "NoSheet", "A1")
        assert cv.type == CellType.BLANK
        plxl.close_workbook(wb)

    def test_sheet_names(
        self, plxl: PolarsAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_fixture(opxl, path)
        wb = plxl.open_workbook(path)
        names = plxl.get_sheet_names(wb)
        assert "S1" in names
        plxl.close_workbook(wb)


# ═════════════════════════════════════════════════════════════════════════
# TestPolarsReadStubs
# ═════════════════════════════════════════════════════════════════════════


class TestPolarsReadStubs:
    """Tier-2 reads return empty, write ops raise."""

    def test_all_stubs(self, plxl: PolarsAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_fixture(opxl, path)
        wb = plxl.open_workbook(path)
        assert plxl.read_cell_format(wb, "S1", "A1") == CellFormat()
        assert plxl.read_cell_border(wb, "S1", "A1") == BorderInfo()
        assert plxl.read_merged_ranges(wb, "S1") == []
        assert plxl.read_conditional_formats(wb, "S1") == []
        assert plxl.read_data_validations(wb, "S1") == []
        assert plxl.read_hyperlinks(wb, "S1") == []
        assert plxl.read_images(wb, "S1") == []
        assert plxl.read_pivot_tables(wb, "S1") == []
        assert plxl.read_comments(wb, "S1") == []
        assert plxl.read_freeze_panes(wb, "S1") == {}
        assert plxl.read_row_height(wb, "S1", 1) is None
        assert plxl.read_column_width(wb, "S1", "A") is None
        plxl.close_workbook(wb)

    def test_write_raises(self, plxl: PolarsAdapter) -> None:
        with pytest.raises(NotImplementedError):
            plxl.create_workbook()
        with pytest.raises(NotImplementedError):
            plxl.save_workbook(None, Path("/tmp/x.xlsx"))
