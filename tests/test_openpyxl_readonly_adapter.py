"""Tests for OpenpyxlReadonlyAdapter (read-only, streaming mode)."""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

import pytest

from excelbench.harness.adapters.openpyxl_adapter import OpenpyxlAdapter
from excelbench.harness.adapters.openpyxl_readonly_adapter import OpenpyxlReadonlyAdapter
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
def ro() -> OpenpyxlReadonlyAdapter:
    return OpenpyxlReadonlyAdapter()


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
# TestReadonlyInfo
# ═════════════════════════════════════════════════════════════════════════


class TestReadonlyInfo:
    def test_name(self, ro: OpenpyxlReadonlyAdapter) -> None:
        assert ro.info.name == "openpyxl-readonly"

    def test_version(self, ro: OpenpyxlReadonlyAdapter) -> None:
        assert ro.info.version != "unknown"

    def test_capabilities(self, ro: OpenpyxlReadonlyAdapter) -> None:
        assert ro.info.capabilities == {"read"}

    def test_is_read_only(self, ro: OpenpyxlReadonlyAdapter) -> None:
        assert ro.can_read()
        assert not ro.can_write()


# ═════════════════════════════════════════════════════════════════════════
# TestReadonlyReadCellValue
# ═════════════════════════════════════════════════════════════════════════


class TestReadonlyReadCellValue:
    """Read via openpyxl read-only mode from openpyxl-written fixtures."""

    def test_string(
        self, ro: OpenpyxlReadonlyAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_fixture(opxl, path)
        wb = ro.open_workbook(path)
        cv = ro.read_cell_value(wb, "S1", "A1")
        assert cv.type == CellType.STRING
        assert cv.value == "hello"
        ro.close_workbook(wb)

    def test_number(
        self, ro: OpenpyxlReadonlyAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_fixture(opxl, path)
        wb = ro.open_workbook(path)
        cv = ro.read_cell_value(wb, "S1", "A2")
        assert cv.type == CellType.NUMBER
        assert cv.value == 42.5
        ro.close_workbook(wb)

    def test_boolean(
        self, ro: OpenpyxlReadonlyAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_fixture(opxl, path)
        wb = ro.open_workbook(path)
        cv = ro.read_cell_value(wb, "S1", "A3")
        assert cv.type == CellType.BOOLEAN
        assert cv.value is True
        ro.close_workbook(wb)

    def test_date(
        self, ro: OpenpyxlReadonlyAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_fixture(opxl, path)
        wb = ro.open_workbook(path)
        cv = ro.read_cell_value(wb, "S1", "A4")
        assert cv.type == CellType.DATE
        assert cv.value == date(2024, 6, 15)
        ro.close_workbook(wb)

    def test_datetime(
        self, ro: OpenpyxlReadonlyAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_fixture(opxl, path)
        wb = ro.open_workbook(path)
        cv = ro.read_cell_value(wb, "S1", "A5")
        assert cv.type == CellType.DATETIME
        ro.close_workbook(wb)

    def test_error(
        self, ro: OpenpyxlReadonlyAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_fixture(opxl, path)
        wb = ro.open_workbook(path)
        cv = ro.read_cell_value(wb, "S1", "A6")
        # openpyxl read-only may see error formula or error string
        assert cv.type in (CellType.ERROR, CellType.FORMULA, CellType.STRING)
        ro.close_workbook(wb)

    def test_formula(
        self, ro: OpenpyxlReadonlyAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_fixture(opxl, path)
        wb = ro.open_workbook(path)
        cv = ro.read_cell_value(wb, "S1", "A7")
        assert cv.type == CellType.FORMULA
        ro.close_workbook(wb)

    def test_blank(
        self, ro: OpenpyxlReadonlyAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_fixture(opxl, path)
        wb = ro.open_workbook(path)
        cv = ro.read_cell_value(wb, "S1", "A8")
        assert cv.type == CellType.BLANK
        ro.close_workbook(wb)

    def test_sheet_names(
        self, ro: OpenpyxlReadonlyAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_fixture(opxl, path)
        wb = ro.open_workbook(path)
        names = ro.get_sheet_names(wb)
        assert "S1" in names
        ro.close_workbook(wb)

    def test_close_required(
        self, ro: OpenpyxlReadonlyAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        """Verify close_workbook doesn't raise."""
        path = tmp_path / "fixture.xlsx"
        _write_fixture(opxl, path)
        wb = ro.open_workbook(path)
        ro.close_workbook(wb)  # Should not raise


# ═════════════════════════════════════════════════════════════════════════
# TestReadonlyStubs
# ═════════════════════════════════════════════════════════════════════════


class TestReadonlyStubs:
    """Formatting/tier-2 reads return defaults; writes raise."""

    def test_read_stubs(
        self, ro: OpenpyxlReadonlyAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_fixture(opxl, path)
        wb = ro.open_workbook(path)
        assert ro.read_cell_format(wb, "S1", "A1") == CellFormat()
        assert ro.read_cell_border(wb, "S1", "A1") == BorderInfo()
        assert ro.read_row_height(wb, "S1", 1) is None
        assert ro.read_column_width(wb, "S1", "A") is None
        assert ro.read_conditional_formats(wb, "S1") == []
        assert ro.read_data_validations(wb, "S1") == []
        assert ro.read_hyperlinks(wb, "S1") == []
        assert ro.read_images(wb, "S1") == []
        assert ro.read_pivot_tables(wb, "S1") == []
        assert ro.read_comments(wb, "S1") == []
        assert ro.read_freeze_panes(wb, "S1") == {}
        ro.close_workbook(wb)

    def test_write_raises(self, ro: OpenpyxlReadonlyAdapter) -> None:
        with pytest.raises(NotImplementedError):
            ro.create_workbook()
        with pytest.raises(NotImplementedError):
            ro.save_workbook(None, Path("/tmp/x.xlsx"))
