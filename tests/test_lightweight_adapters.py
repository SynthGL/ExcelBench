"""Tests for pyexcel, pylightxl, calamine, and Rust-backed adapters.

Each adapter is tested via write→read roundtrip (or read-only via openpyxl-
written fixtures) to cover as many code paths as possible.
"""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

import pytest

from excelbench.harness.adapters.openpyxl_adapter import OpenpyxlAdapter
from excelbench.harness.adapters.pyexcel_adapter import PyexcelAdapter
from excelbench.harness.adapters.pylightxl_adapter import PylightxlAdapter
from excelbench.models import (
    CellFormat,
    CellType,
    CellValue,
)

# Optional Rust adapters
try:
    from excelbench.harness.adapters.umya_adapter import UmyaAdapter
except ImportError:
    UmyaAdapter = None  # type: ignore[assignment,misc]
try:
    from excelbench.harness.adapters.rust_xlsxwriter_adapter import RustXlsxWriterAdapter
except ImportError:
    RustXlsxWriterAdapter = None  # type: ignore[assignment,misc]
try:
    from excelbench.harness.adapters.rust_calamine_adapter import RustCalamineAdapter
except ImportError:
    RustCalamineAdapter = None  # type: ignore[assignment,misc]
try:
    from excelbench.harness.adapters.calamine_adapter import CalamineAdapter
except ImportError:
    CalamineAdapter = None  # type: ignore[assignment,misc]


@pytest.fixture
def opxl() -> OpenpyxlAdapter:
    return OpenpyxlAdapter()


@pytest.fixture
def pyxl() -> PyexcelAdapter:
    return PyexcelAdapter()


@pytest.fixture
def plxl() -> PylightxlAdapter:
    return PylightxlAdapter()


def _write_openpyxl_fixture(opxl: OpenpyxlAdapter, path: Path) -> None:
    """Write a multi-type .xlsx fixture using openpyxl for read-only adapters."""
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
# PyexcelAdapter tests
# ═════════════════════════════════════════════════════════════════════════


class TestPyexcelInfo:
    def test_info(self, pyxl: PyexcelAdapter) -> None:
        info = pyxl.info
        assert info.name == "pyexcel"
        assert "read" in info.capabilities
        assert "write" in info.capabilities


class TestPyexcelReadCellValue:
    """Read via pyexcel from openpyxl-written fixtures."""

    def test_string(self, pyxl: PyexcelAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        wb = pyxl.open_workbook(path)
        cv = pyxl.read_cell_value(wb, "S1", "A1")
        assert cv.type == CellType.STRING
        assert cv.value == "hello"
        pyxl.close_workbook(wb)

    def test_number(self, pyxl: PyexcelAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        wb = pyxl.open_workbook(path)
        cv = pyxl.read_cell_value(wb, "S1", "A2")
        assert cv.type == CellType.NUMBER
        assert cv.value == 42.5
        pyxl.close_workbook(wb)

    def test_boolean(self, pyxl: PyexcelAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        wb = pyxl.open_workbook(path)
        cv = pyxl.read_cell_value(wb, "S1", "A3")
        assert cv.type == CellType.BOOLEAN
        assert cv.value is True
        pyxl.close_workbook(wb)

    def test_date(self, pyxl: PyexcelAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        wb = pyxl.open_workbook(path)
        cv = pyxl.read_cell_value(wb, "S1", "A4")
        assert cv.type == CellType.DATE
        assert cv.value == date(2024, 6, 15)
        pyxl.close_workbook(wb)

    def test_datetime(self, pyxl: PyexcelAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        wb = pyxl.open_workbook(path)
        cv = pyxl.read_cell_value(wb, "S1", "A5")
        assert cv.type == CellType.DATETIME
        pyxl.close_workbook(wb)

    def test_error(self, pyxl: PyexcelAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        wb = pyxl.open_workbook(path)
        cv = pyxl.read_cell_value(wb, "S1", "A6")
        # pyexcel may return error as string, or blank (openpyxl error type not round-tripped)
        assert cv.type in (CellType.ERROR, CellType.STRING, CellType.BLANK)
        pyxl.close_workbook(wb)

    def test_blank_out_of_bounds(
        self, pyxl: PyexcelAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        wb = pyxl.open_workbook(path)
        # Row out of bounds
        cv = pyxl.read_cell_value(wb, "S1", "A99")
        assert cv.type == CellType.BLANK
        # Col out of bounds
        cv2 = pyxl.read_cell_value(wb, "S1", "Z1")
        assert cv2.type == CellType.BLANK
        pyxl.close_workbook(wb)

    def test_sheet_not_found(
        self, pyxl: PyexcelAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        wb = pyxl.open_workbook(path)
        cv = pyxl.read_cell_value(wb, "NoSheet", "A1")
        assert cv.type == CellType.BLANK
        pyxl.close_workbook(wb)

    def test_sheet_names(self, pyxl: PyexcelAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        wb = pyxl.open_workbook(path)
        names = pyxl.get_sheet_names(wb)
        assert "S1" in names
        pyxl.close_workbook(wb)

    def test_tier2_stubs(self, pyxl: PyexcelAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        wb = pyxl.open_workbook(path)
        assert pyxl.read_cell_format(wb, "S1", "A1") == CellFormat()
        assert pyxl.read_merged_ranges(wb, "S1") == []
        assert pyxl.read_conditional_formats(wb, "S1") == []
        assert pyxl.read_data_validations(wb, "S1") == []
        assert pyxl.read_hyperlinks(wb, "S1") == []
        assert pyxl.read_images(wb, "S1") == []
        assert pyxl.read_pivot_tables(wb, "S1") == []
        assert pyxl.read_comments(wb, "S1") == []
        assert pyxl.read_freeze_panes(wb, "S1") == {}
        assert pyxl.read_row_height(wb, "S1", 1) is None
        assert pyxl.read_column_width(wb, "S1", "A") is None
        pyxl.close_workbook(wb)


class TestPyexcelWriteRoundtrip:
    """Write via pyexcel, read back via openpyxl."""

    def test_write_cell_types(
        self, pyxl: PyexcelAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "pyxl_out.xlsx"
        wb = pyxl.create_workbook()
        pyxl.add_sheet(wb, "S1")
        pyxl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="hi"))
        pyxl.write_cell_value(wb, "S1", "A2", CellValue(type=CellType.NUMBER, value=99))
        pyxl.write_cell_value(wb, "S1", "A3", CellValue(type=CellType.BOOLEAN, value=False))
        pyxl.write_cell_value(wb, "S1", "A4", CellValue(type=CellType.BLANK))
        pyxl.write_cell_value(wb, "S1", "A5", CellValue(type=CellType.ERROR, value="#VALUE!"))
        pyxl.write_cell_value(
            wb,
            "S1",
            "A6",
            CellValue(type=CellType.DATE, value=date(2024, 1, 1)),
        )
        pyxl.write_cell_value(
            wb,
            "S1",
            "A7",
            CellValue(type=CellType.DATETIME, value=datetime(2024, 1, 1, 12, 0, 0)),
        )
        pyxl.write_cell_value(
            wb,
            "S1",
            "A8",
            CellValue(type=CellType.FORMULA, value="=1+1", formula="=1+1"),
        )
        pyxl.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        cv1 = opxl.read_cell_value(rb, "S1", "A1")
        assert cv1.value == "hi"
        cv2 = opxl.read_cell_value(rb, "S1", "A2")
        assert cv2.value == 99
        opxl.close_workbook(rb)

    def test_write_auto_creates_sheet(self, pyxl: PyexcelAdapter, tmp_path: Path) -> None:
        wb = pyxl.create_workbook()
        # Don't call add_sheet — write_cell_value should auto-create
        pyxl.write_cell_value(wb, "Auto", "A1", CellValue(type=CellType.STRING, value="x"))
        assert "Auto" in wb["sheets"]

    def test_write_noop_methods(self, pyxl: PyexcelAdapter) -> None:
        wb = pyxl.create_workbook()
        pyxl.add_sheet(wb, "S1")
        # All no-ops — should not raise
        pyxl.write_cell_format(wb, "S1", "A1", CellFormat())
        pyxl.set_row_height(wb, "S1", 1, 30.0)
        pyxl.set_column_width(wb, "S1", "A", 20.0)
        pyxl.merge_cells(wb, "S1", "A1:B2")
        pyxl.add_conditional_format(wb, "S1", {})
        pyxl.add_data_validation(wb, "S1", {})
        pyxl.add_hyperlink(wb, "S1", {})
        pyxl.add_image(wb, "S1", {})
        pyxl.add_pivot_table(wb, "S1", {})
        pyxl.add_comment(wb, "S1", {})
        pyxl.set_freeze_panes(wb, "S1", {})

    def test_save_empty_sheet(
        self, pyxl: PyexcelAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "empty.xlsx"
        wb = pyxl.create_workbook()
        pyxl.add_sheet(wb, "EmptySheet")
        pyxl.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        assert "EmptySheet" in opxl.get_sheet_names(rb)
        opxl.close_workbook(rb)


# ═════════════════════════════════════════════════════════════════════════
# PylightxlAdapter tests
# ═════════════════════════════════════════════════════════════════════════


class TestPylightxlInfo:
    def test_info(self, plxl: PylightxlAdapter) -> None:
        info = plxl.info
        assert info.name == "pylightxl"
        assert "read" in info.capabilities
        assert "write" in info.capabilities

    def test_extensions(self, plxl: PylightxlAdapter) -> None:
        assert ".xlsx" in plxl.supported_read_extensions


class TestPylightxlReadCellValue:
    """Self-roundtrip: write via pylightxl, read back via pylightxl."""

    def test_string(self, plxl: PylightxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "plxl.xlsx"
        wb = plxl.create_workbook()
        plxl.add_sheet(wb, "S1")
        plxl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="hello"))
        plxl.save_workbook(wb, path)

        rb = plxl.open_workbook(path)
        names = plxl.get_sheet_names(rb)
        assert "S1" in names
        cv = plxl.read_cell_value(rb, "S1", "A1")
        assert cv.type == CellType.STRING
        assert cv.value == "hello"
        plxl.close_workbook(rb)

    def test_number(self, plxl: PylightxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "plxl_num.xlsx"
        wb = plxl.create_workbook()
        plxl.add_sheet(wb, "S1")
        plxl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.NUMBER, value=42.5))
        plxl.save_workbook(wb, path)

        rb = plxl.open_workbook(path)
        cv = plxl.read_cell_value(rb, "S1", "A1")
        assert cv.type == CellType.NUMBER
        assert cv.value == 42.5
        plxl.close_workbook(rb)

    def test_boolean_as_int(self, plxl: PylightxlAdapter, tmp_path: Path) -> None:
        """pylightxl writes booleans as 0/1 integers."""
        path = tmp_path / "plxl_bool.xlsx"
        wb = plxl.create_workbook()
        plxl.add_sheet(wb, "S1")
        plxl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.BOOLEAN, value=True))
        plxl.save_workbook(wb, path)

        rb = plxl.open_workbook(path)
        cv = plxl.read_cell_value(rb, "S1", "A1")
        # pylightxl writes bool as int(True)=1
        assert cv.type in (CellType.BOOLEAN, CellType.NUMBER)
        plxl.close_workbook(rb)

    def test_blank(self, plxl: PylightxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "plxl_blank.xlsx"
        wb = plxl.create_workbook()
        plxl.add_sheet(wb, "S1")
        plxl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.BLANK))
        plxl.write_cell_value(wb, "S1", "A2", CellValue(type=CellType.STRING, value="x"))
        plxl.save_workbook(wb, path)

        rb = plxl.open_workbook(path)
        cv = plxl.read_cell_value(rb, "S1", "A1")
        assert cv.type == CellType.BLANK
        plxl.close_workbook(rb)

    def test_date_iso(self, plxl: PylightxlAdapter, tmp_path: Path) -> None:
        """pylightxl writes dates as ISO strings."""
        path = tmp_path / "plxl_date.xlsx"
        wb = plxl.create_workbook()
        plxl.add_sheet(wb, "S1")
        plxl.write_cell_value(
            wb, "S1", "A1", CellValue(type=CellType.DATE, value=date(2024, 6, 15))
        )
        plxl.save_workbook(wb, path)

        rb = plxl.open_workbook(path)
        cv = plxl.read_cell_value(rb, "S1", "A1")
        # Comes back as string "2024-06-15"
        assert cv.type in (CellType.DATE, CellType.STRING)
        plxl.close_workbook(rb)

    def test_datetime_iso(self, plxl: PylightxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "plxl_dt.xlsx"
        wb = plxl.create_workbook()
        plxl.add_sheet(wb, "S1")
        plxl.write_cell_value(
            wb,
            "S1",
            "A1",
            CellValue(type=CellType.DATETIME, value=datetime(2024, 6, 15, 14, 30)),
        )
        plxl.save_workbook(wb, path)

        rb = plxl.open_workbook(path)
        cv = plxl.read_cell_value(rb, "S1", "A1")
        assert cv.type in (CellType.DATETIME, CellType.STRING)
        plxl.close_workbook(rb)

    def test_error(self, plxl: PylightxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "plxl_err.xlsx"
        wb = plxl.create_workbook()
        plxl.add_sheet(wb, "S1")
        plxl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.ERROR, value="#N/A"))
        plxl.save_workbook(wb, path)

        rb = plxl.open_workbook(path)
        cv = plxl.read_cell_value(rb, "S1", "A1")
        assert cv.type in (CellType.ERROR, CellType.STRING)
        plxl.close_workbook(rb)

    def test_formula(self, plxl: PylightxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "plxl_form.xlsx"
        wb = plxl.create_workbook()
        plxl.add_sheet(wb, "S1")
        plxl.write_cell_value(
            wb,
            "S1",
            "A1",
            CellValue(type=CellType.FORMULA, value="=SUM(B1)", formula="=SUM(B1)"),
        )
        plxl.save_workbook(wb, path)

        rb = plxl.open_workbook(path)
        cv = plxl.read_cell_value(rb, "S1", "A1")
        assert cv.type in (CellType.FORMULA, CellType.STRING, CellType.BLANK)
        plxl.close_workbook(rb)

    def test_tier2_stubs(self, plxl: PylightxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "plxl_stubs.xlsx"
        wb = plxl.create_workbook()
        plxl.add_sheet(wb, "S1")
        plxl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="x"))
        plxl.save_workbook(wb, path)

        rb = plxl.open_workbook(path)
        assert plxl.read_cell_format(rb, "S1", "A1") == CellFormat()
        assert plxl.read_merged_ranges(rb, "S1") == []
        assert plxl.read_conditional_formats(rb, "S1") == []
        assert plxl.read_data_validations(rb, "S1") == []
        assert plxl.read_hyperlinks(rb, "S1") == []
        assert plxl.read_images(rb, "S1") == []
        assert plxl.read_pivot_tables(rb, "S1") == []
        assert plxl.read_comments(rb, "S1") == []
        assert plxl.read_freeze_panes(rb, "S1") == {}
        assert plxl.read_row_height(rb, "S1", 1) is None
        assert plxl.read_column_width(rb, "S1", "A") is None
        plxl.close_workbook(rb)


class TestPylightxlWriteRoundtrip:
    def test_write_cell_types(
        self, plxl: PylightxlAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "plxl_out.xlsx"
        wb = plxl.create_workbook()
        plxl.add_sheet(wb, "S1")
        plxl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="text"))
        plxl.write_cell_value(wb, "S1", "A2", CellValue(type=CellType.NUMBER, value=3.14))
        plxl.write_cell_value(wb, "S1", "A3", CellValue(type=CellType.BOOLEAN, value=True))
        plxl.write_cell_value(wb, "S1", "A4", CellValue(type=CellType.BLANK))
        plxl.write_cell_value(wb, "S1", "A5", CellValue(type=CellType.ERROR, value="#DIV/0!"))
        plxl.write_cell_value(wb, "S1", "A6", CellValue(type=CellType.DATE, value=date(2024, 3, 1)))
        plxl.write_cell_value(
            wb,
            "S1",
            "A7",
            CellValue(type=CellType.DATETIME, value=datetime(2024, 3, 1, 10, 0)),
        )
        plxl.write_cell_value(
            wb,
            "S1",
            "A8",
            CellValue(type=CellType.FORMULA, value="=SUM(A1)", formula="=SUM(A1)"),
        )
        plxl.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        cv1 = opxl.read_cell_value(rb, "S1", "A1")
        assert cv1.value == "text"
        opxl.close_workbook(rb)

    def test_write_noop_methods(self, plxl: PylightxlAdapter) -> None:
        wb = plxl.create_workbook()
        plxl.add_sheet(wb, "S1")
        plxl.write_cell_format(wb, "S1", "A1", CellFormat())
        plxl.set_row_height(wb, "S1", 1, 20.0)
        plxl.set_column_width(wb, "S1", "A", 15.0)
        plxl.merge_cells(wb, "S1", "A1:B2")
        plxl.add_conditional_format(wb, "S1", {})
        plxl.add_data_validation(wb, "S1", {})
        plxl.add_hyperlink(wb, "S1", {})
        plxl.add_image(wb, "S1", {})
        plxl.add_pivot_table(wb, "S1", {})
        plxl.add_comment(wb, "S1", {})
        plxl.set_freeze_panes(wb, "S1", {})

    def test_save_overwrites_existing(self, plxl: PylightxlAdapter, tmp_path: Path) -> None:
        """pylightxl removes existing file before writing (path.unlink)."""
        path = tmp_path / "overwrite.xlsx"
        # Create a dummy file first
        path.write_bytes(b"not a zip")
        wb = plxl.create_workbook()
        plxl.add_sheet(wb, "S1")
        plxl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="ok"))
        plxl.save_workbook(wb, path)
        assert path.exists()


# ═════════════════════════════════════════════════════════════════════════
# CalamineAdapter tests (read-only, Python binding)
# ═════════════════════════════════════════════════════════════════════════


@pytest.mark.skipif(CalamineAdapter is None, reason="python-calamine not installed")
class TestCalamineReadCellValue:
    def test_info(self) -> None:
        adapter = CalamineAdapter()
        assert adapter.info.name == "python-calamine"
        assert "read" in adapter.info.capabilities
        assert ".xlsx" in adapter.supported_read_extensions

    def test_string(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        adapter = CalamineAdapter()
        wb = adapter.open_workbook(path)
        names = adapter.get_sheet_names(wb)
        assert "S1" in names
        cv = adapter.read_cell_value(wb, "S1", "A1")
        assert cv.type == CellType.STRING
        assert cv.value == "hello"
        adapter.close_workbook(wb)

    def test_number(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        adapter = CalamineAdapter()
        wb = adapter.open_workbook(path)
        cv = adapter.read_cell_value(wb, "S1", "A2")
        assert cv.type == CellType.NUMBER
        assert cv.value == 42.5
        adapter.close_workbook(wb)

    def test_boolean(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        adapter = CalamineAdapter()
        wb = adapter.open_workbook(path)
        cv = adapter.read_cell_value(wb, "S1", "A3")
        assert cv.type == CellType.BOOLEAN
        assert cv.value is True
        adapter.close_workbook(wb)

    def test_date(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        adapter = CalamineAdapter()
        wb = adapter.open_workbook(path)
        cv = adapter.read_cell_value(wb, "S1", "A4")
        assert cv.type in (CellType.DATE, CellType.DATETIME)
        adapter.close_workbook(wb)

    def test_datetime(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        adapter = CalamineAdapter()
        wb = adapter.open_workbook(path)
        cv = adapter.read_cell_value(wb, "S1", "A5")
        assert cv.type == CellType.DATETIME
        adapter.close_workbook(wb)

    def test_blank_out_of_bounds(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        adapter = CalamineAdapter()
        wb = adapter.open_workbook(path)
        cv = adapter.read_cell_value(wb, "S1", "Z99")
        assert cv.type == CellType.BLANK
        adapter.close_workbook(wb)

    def test_tier2_stubs(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        adapter = CalamineAdapter()
        wb = adapter.open_workbook(path)
        assert adapter.read_cell_format(wb, "S1", "A1") == CellFormat()
        assert adapter.read_merged_ranges(wb, "S1") == []
        assert adapter.read_conditional_formats(wb, "S1") == []
        assert adapter.read_data_validations(wb, "S1") == []
        assert adapter.read_hyperlinks(wb, "S1") == []
        assert adapter.read_images(wb, "S1") == []
        assert adapter.read_pivot_tables(wb, "S1") == []
        assert adapter.read_comments(wb, "S1") == []
        assert adapter.read_freeze_panes(wb, "S1") == {}
        assert adapter.read_row_height(wb, "S1", 1) is None
        assert adapter.read_column_width(wb, "S1", "A") is None
        adapter.close_workbook(wb)


# ═════════════════════════════════════════════════════════════════════════
# UmyaAdapter tests (read/write via PyO3)
# ═════════════════════════════════════════════════════════════════════════


@pytest.mark.skipif(UmyaAdapter is None, reason="excelbench_rust umya not available")
class TestUmyaWriteRoundtrip:
    def test_info(self) -> None:
        adapter = UmyaAdapter()
        assert adapter.info.name == "umya-spreadsheet"
        assert adapter.info.language == "rust"
        assert "read" in adapter.info.capabilities
        assert "write" in adapter.info.capabilities
        assert ".xlsx" in adapter.supported_read_extensions

    def test_string_roundtrip(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        adapter = UmyaAdapter()
        path = tmp_path / "umya.xlsx"
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        adapter.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="umya"))
        adapter.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        cv = opxl.read_cell_value(rb, "S1", "A1")
        assert cv.value == "umya"
        opxl.close_workbook(rb)

    def test_number_roundtrip(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        adapter = UmyaAdapter()
        path = tmp_path / "umya_num.xlsx"
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        adapter.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.NUMBER, value=3.14))
        adapter.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        cv = opxl.read_cell_value(rb, "S1", "A1")
        assert cv.value == pytest.approx(3.14)
        opxl.close_workbook(rb)

    def test_boolean_roundtrip(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        adapter = UmyaAdapter()
        path = tmp_path / "umya_bool.xlsx"
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        adapter.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.BOOLEAN, value=True))
        adapter.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        cv = opxl.read_cell_value(rb, "S1", "A1")
        assert cv.value is True or cv.value == 1
        opxl.close_workbook(rb)

    def test_blank(self, tmp_path: Path) -> None:
        adapter = UmyaAdapter()
        path = tmp_path / "umya_blank.xlsx"
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        adapter.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.BLANK))
        adapter.write_cell_value(wb, "S1", "A2", CellValue(type=CellType.STRING, value="x"))
        adapter.save_workbook(wb, path)

        rb = adapter.open_workbook(path)
        cv = adapter.read_cell_value(rb, "S1", "A1")
        assert cv.type == CellType.BLANK
        adapter.close_workbook(rb)

    def test_formula(self, tmp_path: Path) -> None:
        adapter = UmyaAdapter()
        path = tmp_path / "umya_formula.xlsx"
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        adapter.write_cell_value(
            wb,
            "S1",
            "A1",
            CellValue(type=CellType.FORMULA, value="=1+1", formula="=1+1"),
        )
        adapter.save_workbook(wb, path)

        rb = adapter.open_workbook(path)
        cv = adapter.read_cell_value(rb, "S1", "A1")
        assert cv.type == CellType.FORMULA
        adapter.close_workbook(rb)

    def test_error(self, tmp_path: Path) -> None:
        adapter = UmyaAdapter()
        path = tmp_path / "umya_err.xlsx"
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        adapter.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.ERROR, value="#N/A"))
        adapter.save_workbook(wb, path)

        rb = adapter.open_workbook(path)
        cv = adapter.read_cell_value(rb, "S1", "A1")
        assert cv.type in (CellType.ERROR, CellType.STRING)
        adapter.close_workbook(rb)

    def test_date_roundtrip(self, tmp_path: Path) -> None:
        adapter = UmyaAdapter()
        path = tmp_path / "umya_date.xlsx"
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        adapter.write_cell_value(
            wb, "S1", "A1", CellValue(type=CellType.DATE, value=date(2024, 6, 15))
        )
        adapter.save_workbook(wb, path)

        rb = adapter.open_workbook(path)
        cv = adapter.read_cell_value(rb, "S1", "A1")
        assert cv.type in (CellType.DATE, CellType.STRING)
        adapter.close_workbook(rb)

    def test_datetime_roundtrip(self, tmp_path: Path) -> None:
        adapter = UmyaAdapter()
        path = tmp_path / "umya_dt.xlsx"
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        adapter.write_cell_value(
            wb,
            "S1",
            "A1",
            CellValue(type=CellType.DATETIME, value=datetime(2024, 6, 15, 14, 30)),
        )
        adapter.save_workbook(wb, path)

        rb = adapter.open_workbook(path)
        cv = adapter.read_cell_value(rb, "S1", "A1")
        assert cv.type in (CellType.DATETIME, CellType.STRING)
        adapter.close_workbook(rb)

    def test_read_from_openpyxl_fixture(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        adapter = UmyaAdapter()
        wb = adapter.open_workbook(path)
        names = adapter.get_sheet_names(wb)
        assert "S1" in names
        cv = adapter.read_cell_value(wb, "S1", "A1")
        assert cv.value == "hello"
        adapter.close_workbook(wb)

    def test_write_noop_methods(self) -> None:
        adapter = UmyaAdapter()
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        # All no-ops should not raise
        adapter.write_cell_format(wb, "S1", "A1", CellFormat())
        adapter.set_row_height(wb, "S1", 1, 20.0)
        adapter.set_column_width(wb, "S1", "A", 15.0)
        adapter.merge_cells(wb, "S1", "A1:B2")
        adapter.add_conditional_format(wb, "S1", {})
        adapter.add_data_validation(wb, "S1", {})
        adapter.add_hyperlink(wb, "S1", {})
        adapter.add_image(wb, "S1", {})
        adapter.add_comment(wb, "S1", {})
        adapter.set_freeze_panes(wb, "S1", {})

    def test_pivot_table_raises(self) -> None:
        adapter = UmyaAdapter()
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        with pytest.raises(NotImplementedError):
            adapter.add_pivot_table(wb, "S1", {})

    def test_read_tier2_stubs(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        adapter = UmyaAdapter()
        wb = adapter.open_workbook(path)
        assert adapter.read_cell_format(wb, "S1", "A1") == CellFormat()
        assert adapter.read_merged_ranges(wb, "S1") == []
        assert adapter.read_conditional_formats(wb, "S1") == []
        assert adapter.read_data_validations(wb, "S1") == []
        assert adapter.read_hyperlinks(wb, "S1") == []
        assert adapter.read_images(wb, "S1") == []
        assert adapter.read_pivot_tables(wb, "S1") == []
        assert adapter.read_comments(wb, "S1") == []
        assert adapter.read_freeze_panes(wb, "S1") == {}
        assert adapter.read_row_height(wb, "S1", 1) is None
        assert adapter.read_column_width(wb, "S1", "A") is None
        adapter.close_workbook(wb)


# ═════════════════════════════════════════════════════════════════════════
# RustXlsxWriterAdapter tests (write-only via PyO3)
# ═════════════════════════════════════════════════════════════════════════


@pytest.mark.skipif(
    RustXlsxWriterAdapter is None, reason="excelbench_rust rust_xlsxwriter not available"
)
class TestRustXlsxWriterRoundtrip:
    def test_info(self) -> None:
        adapter = RustXlsxWriterAdapter()
        assert adapter.info.name == "rust_xlsxwriter"
        assert adapter.info.language == "rust"
        assert "write" in adapter.info.capabilities

    def test_string_roundtrip(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        adapter = RustXlsxWriterAdapter()
        path = tmp_path / "rxw.xlsx"
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        adapter.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="rust"))
        adapter.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        cv = opxl.read_cell_value(rb, "S1", "A1")
        assert cv.value == "rust"
        opxl.close_workbook(rb)

    def test_number_roundtrip(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        adapter = RustXlsxWriterAdapter()
        path = tmp_path / "rxw_num.xlsx"
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        adapter.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.NUMBER, value=2.71))
        adapter.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        cv = opxl.read_cell_value(rb, "S1", "A1")
        assert cv.value == pytest.approx(2.71)
        opxl.close_workbook(rb)

    def test_boolean_roundtrip(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        adapter = RustXlsxWriterAdapter()
        path = tmp_path / "rxw_bool.xlsx"
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        adapter.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.BOOLEAN, value=False))
        adapter.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        cv = opxl.read_cell_value(rb, "S1", "A1")
        assert cv.value is False or cv.value == 0
        opxl.close_workbook(rb)

    def test_formula(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        adapter = RustXlsxWriterAdapter()
        path = tmp_path / "rxw_formula.xlsx"
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        adapter.write_cell_value(
            wb,
            "S1",
            "A1",
            CellValue(type=CellType.FORMULA, value="=2+2", formula="=2+2"),
        )
        adapter.save_workbook(wb, path)
        assert path.exists()

    def test_blank(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        adapter = RustXlsxWriterAdapter()
        path = tmp_path / "rxw_blank.xlsx"
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        adapter.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.BLANK))
        adapter.write_cell_value(wb, "S1", "A2", CellValue(type=CellType.STRING, value="x"))
        adapter.save_workbook(wb, path)
        assert path.exists()

    def test_error(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        adapter = RustXlsxWriterAdapter()
        path = tmp_path / "rxw_err.xlsx"
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        adapter.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.ERROR, value="#N/A"))
        adapter.save_workbook(wb, path)
        assert path.exists()

    def test_date(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        adapter = RustXlsxWriterAdapter()
        path = tmp_path / "rxw_date.xlsx"
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        adapter.write_cell_value(
            wb, "S1", "A1", CellValue(type=CellType.DATE, value=date(2024, 1, 1))
        )
        adapter.save_workbook(wb, path)
        assert path.exists()

    def test_datetime(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        adapter = RustXlsxWriterAdapter()
        path = tmp_path / "rxw_dt.xlsx"
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        adapter.write_cell_value(
            wb,
            "S1",
            "A1",
            CellValue(type=CellType.DATETIME, value=datetime(2024, 1, 1, 12, 30)),
        )
        adapter.save_workbook(wb, path)
        assert path.exists()

    def test_noop_methods(self) -> None:
        adapter = RustXlsxWriterAdapter()
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        adapter.write_cell_format(wb, "S1", "A1", CellFormat())
        adapter.set_row_height(wb, "S1", 1, 20.0)
        adapter.set_column_width(wb, "S1", "A", 15.0)
        adapter.merge_cells(wb, "S1", "A1:B2")
        adapter.add_conditional_format(wb, "S1", {})
        adapter.add_data_validation(wb, "S1", {})
        adapter.add_hyperlink(wb, "S1", {})
        adapter.add_image(wb, "S1", {})
        adapter.add_comment(wb, "S1", {})
        adapter.set_freeze_panes(wb, "S1", {})

    def test_pivot_table_raises(self) -> None:
        adapter = RustXlsxWriterAdapter()
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S1")
        with pytest.raises(NotImplementedError):
            adapter.add_pivot_table(wb, "S1", {})


# ═════════════════════════════════════════════════════════════════════════
# RustCalamineAdapter tests (read-only via PyO3)
# ═════════════════════════════════════════════════════════════════════════


@pytest.mark.skipif(RustCalamineAdapter is None, reason="excelbench_rust calamine not available")
class TestRustCalamineRead:
    def test_info(self) -> None:
        adapter = RustCalamineAdapter()
        assert adapter.info.name == "calamine"
        assert adapter.info.language == "rust"
        assert "read" in adapter.info.capabilities
        assert ".xlsx" in adapter.supported_read_extensions

    def test_string(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        adapter = RustCalamineAdapter()
        wb = adapter.open_workbook(path)
        names = adapter.get_sheet_names(wb)
        assert "S1" in names
        cv = adapter.read_cell_value(wb, "S1", "A1")
        assert cv.value == "hello"
        adapter.close_workbook(wb)

    def test_number(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        adapter = RustCalamineAdapter()
        wb = adapter.open_workbook(path)
        cv = adapter.read_cell_value(wb, "S1", "A2")
        assert cv.type == CellType.NUMBER
        assert cv.value == 42.5
        adapter.close_workbook(wb)

    def test_boolean(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        adapter = RustCalamineAdapter()
        wb = adapter.open_workbook(path)
        cv = adapter.read_cell_value(wb, "S1", "A3")
        assert cv.type == CellType.BOOLEAN
        adapter.close_workbook(wb)

    def test_tier2_stubs(self, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        adapter = RustCalamineAdapter()
        wb = adapter.open_workbook(path)
        assert adapter.read_cell_format(wb, "S1", "A1") == CellFormat()
        assert adapter.read_merged_ranges(wb, "S1") == []
        assert adapter.read_conditional_formats(wb, "S1") == []
        assert adapter.read_data_validations(wb, "S1") == []
        assert adapter.read_hyperlinks(wb, "S1") == []
        assert adapter.read_images(wb, "S1") == []
        assert adapter.read_pivot_tables(wb, "S1") == []
        assert adapter.read_comments(wb, "S1") == []
        assert adapter.read_freeze_panes(wb, "S1") == {}
        assert adapter.read_row_height(wb, "S1", 1) is None
        assert adapter.read_column_width(wb, "S1", "A") is None
        adapter.close_workbook(wb)
