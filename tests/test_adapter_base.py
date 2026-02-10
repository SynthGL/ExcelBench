"""Tests for ReadOnlyAdapter and WriteOnlyAdapter base class contracts."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from excelbench.harness.adapters.base import ReadOnlyAdapter, WriteOnlyAdapter
from excelbench.models import (
    BorderInfo,
    CellFormat,
    CellType,
    CellValue,
    DiagnosticCategory,
    LibraryInfo,
    OperationType,
)

JSONDict = dict[str, Any]


class ConcreteReadOnly(ReadOnlyAdapter):
    """Minimal concrete ReadOnlyAdapter for testing write-method guards."""

    @property
    def info(self) -> LibraryInfo:
        return LibraryInfo(
            name="test-readonly", version="0.0", language="python", capabilities={"read"}
        )

    def open_workbook(self, path: Path) -> Any:
        return None

    def close_workbook(self, workbook: Any) -> None:
        pass

    def get_sheet_names(self, workbook: Any) -> list[str]:
        return []

    def read_cell_value(self, workbook: Any, sheet: str, cell: str) -> CellValue:
        return CellValue(type=CellType.BLANK)

    def read_cell_format(self, workbook: Any, sheet: str, cell: str) -> CellFormat:
        return CellFormat()

    def read_cell_border(self, workbook: Any, sheet: str, cell: str) -> BorderInfo:
        return BorderInfo()

    def read_row_height(self, workbook: Any, sheet: str, row: int) -> float | None:
        return None

    def read_column_width(self, workbook: Any, sheet: str, column: str) -> float | None:
        return None

    def read_merged_ranges(self, workbook: Any, sheet: str) -> list[str]:
        return []

    def read_conditional_formats(self, workbook: Any, sheet: str) -> list[JSONDict]:
        return []

    def read_data_validations(self, workbook: Any, sheet: str) -> list[JSONDict]:
        return []

    def read_hyperlinks(self, workbook: Any, sheet: str) -> list[JSONDict]:
        return []

    def read_images(self, workbook: Any, sheet: str) -> list[JSONDict]:
        return []

    def read_pivot_tables(self, workbook: Any, sheet: str) -> list[JSONDict]:
        return []

    def read_comments(self, workbook: Any, sheet: str) -> list[JSONDict]:
        return []

    def read_freeze_panes(self, workbook: Any, sheet: str) -> JSONDict:
        return {}


class ConcreteWriteOnly(WriteOnlyAdapter):
    """Minimal concrete WriteOnlyAdapter for testing read-method guards."""

    @property
    def info(self) -> LibraryInfo:
        return LibraryInfo(
            name="test-writeonly", version="0.0", language="python", capabilities={"write"}
        )

    def create_workbook(self) -> Any:
        return None

    def add_sheet(self, workbook: Any, name: str) -> None:
        pass

    def write_cell_value(self, workbook: Any, sheet: str, cell: str, value: CellValue) -> None:
        pass

    def write_cell_format(self, workbook: Any, sheet: str, cell: str, format: CellFormat) -> None:
        pass

    def write_cell_border(self, workbook: Any, sheet: str, cell: str, border: BorderInfo) -> None:
        pass

    def save_workbook(self, workbook: Any, path: Path) -> None:
        pass

    def set_row_height(self, workbook: Any, sheet: str, row: int, height: float) -> None:
        pass

    def set_column_width(self, workbook: Any, sheet: str, column: str, width: float) -> None:
        pass

    def merge_cells(self, workbook: Any, sheet: str, cell_range: str) -> None:
        pass

    def add_conditional_format(self, workbook: Any, sheet: str, rule: JSONDict) -> None:
        pass

    def add_data_validation(self, workbook: Any, sheet: str, validation: JSONDict) -> None:
        pass

    def add_hyperlink(self, workbook: Any, sheet: str, link: JSONDict) -> None:
        pass

    def add_image(self, workbook: Any, sheet: str, image: JSONDict) -> None:
        pass

    def add_pivot_table(self, workbook: Any, sheet: str, pivot: JSONDict) -> None:
        pass

    def add_comment(self, workbook: Any, sheet: str, comment: JSONDict) -> None:
        pass

    def set_freeze_panes(self, workbook: Any, sheet: str, settings: JSONDict) -> None:
        pass


class TestReadOnlyAdapter:
    """ReadOnlyAdapter should raise NotImplementedError for all write methods."""

    @pytest.fixture
    def adapter(self) -> ConcreteReadOnly:
        return ConcreteReadOnly()

    def test_create_workbook(self, adapter: ConcreteReadOnly) -> None:
        with pytest.raises(NotImplementedError, match="read-only"):
            adapter.create_workbook()

    def test_add_sheet(self, adapter: ConcreteReadOnly) -> None:
        with pytest.raises(NotImplementedError, match="read-only"):
            adapter.add_sheet(None, "Sheet1")

    def test_write_cell_value(self, adapter: ConcreteReadOnly) -> None:
        with pytest.raises(NotImplementedError, match="read-only"):
            adapter.write_cell_value(None, "S", "A1", CellValue(type=CellType.STRING, value="x"))

    def test_write_cell_format(self, adapter: ConcreteReadOnly) -> None:
        with pytest.raises(NotImplementedError, match="read-only"):
            adapter.write_cell_format(None, "S", "A1", CellFormat())

    def test_write_cell_border(self, adapter: ConcreteReadOnly) -> None:
        with pytest.raises(NotImplementedError, match="read-only"):
            adapter.write_cell_border(None, "S", "A1", BorderInfo())

    def test_save_workbook(self, adapter: ConcreteReadOnly) -> None:
        with pytest.raises(NotImplementedError, match="read-only"):
            adapter.save_workbook(None, Path("test.xlsx"))

    def test_set_row_height(self, adapter: ConcreteReadOnly) -> None:
        with pytest.raises(NotImplementedError, match="read-only"):
            adapter.set_row_height(None, "S", 1, 20.0)

    def test_set_column_width(self, adapter: ConcreteReadOnly) -> None:
        with pytest.raises(NotImplementedError, match="read-only"):
            adapter.set_column_width(None, "S", "A", 15.0)

    def test_merge_cells(self, adapter: ConcreteReadOnly) -> None:
        with pytest.raises(NotImplementedError, match="read-only"):
            adapter.merge_cells(None, "S", "A1:B2")

    def test_add_conditional_format(self, adapter: ConcreteReadOnly) -> None:
        with pytest.raises(NotImplementedError, match="read-only"):
            adapter.add_conditional_format(None, "S", {})

    def test_add_data_validation(self, adapter: ConcreteReadOnly) -> None:
        with pytest.raises(NotImplementedError, match="read-only"):
            adapter.add_data_validation(None, "S", {})

    def test_add_hyperlink(self, adapter: ConcreteReadOnly) -> None:
        with pytest.raises(NotImplementedError, match="read-only"):
            adapter.add_hyperlink(None, "S", {})

    def test_add_image(self, adapter: ConcreteReadOnly) -> None:
        with pytest.raises(NotImplementedError, match="read-only"):
            adapter.add_image(None, "S", {})

    def test_add_pivot_table(self, adapter: ConcreteReadOnly) -> None:
        with pytest.raises(NotImplementedError, match="read-only"):
            adapter.add_pivot_table(None, "S", {})

    def test_add_comment(self, adapter: ConcreteReadOnly) -> None:
        with pytest.raises(NotImplementedError, match="read-only"):
            adapter.add_comment(None, "S", {})

    def test_set_freeze_panes(self, adapter: ConcreteReadOnly) -> None:
        with pytest.raises(NotImplementedError, match="read-only"):
            adapter.set_freeze_panes(None, "S", {})


class TestWriteOnlyAdapter:
    """WriteOnlyAdapter should raise NotImplementedError for all read methods."""

    @pytest.fixture
    def adapter(self) -> ConcreteWriteOnly:
        return ConcreteWriteOnly()

    def test_open_workbook(self, adapter: ConcreteWriteOnly) -> None:
        with pytest.raises(NotImplementedError, match="write-only"):
            adapter.open_workbook(Path("test.xlsx"))

    def test_close_workbook_is_noop(self, adapter: ConcreteWriteOnly) -> None:
        adapter.close_workbook(None)

    def test_get_sheet_names(self, adapter: ConcreteWriteOnly) -> None:
        with pytest.raises(NotImplementedError, match="write-only"):
            adapter.get_sheet_names(None)

    def test_read_cell_value(self, adapter: ConcreteWriteOnly) -> None:
        with pytest.raises(NotImplementedError, match="write-only"):
            adapter.read_cell_value(None, "S", "A1")

    def test_read_cell_format(self, adapter: ConcreteWriteOnly) -> None:
        with pytest.raises(NotImplementedError, match="write-only"):
            adapter.read_cell_format(None, "S", "A1")

    def test_read_cell_border(self, adapter: ConcreteWriteOnly) -> None:
        with pytest.raises(NotImplementedError, match="write-only"):
            adapter.read_cell_border(None, "S", "A1")

    def test_read_row_height(self, adapter: ConcreteWriteOnly) -> None:
        with pytest.raises(NotImplementedError, match="write-only"):
            adapter.read_row_height(None, "S", 1)

    def test_read_column_width(self, adapter: ConcreteWriteOnly) -> None:
        with pytest.raises(NotImplementedError, match="write-only"):
            adapter.read_column_width(None, "S", "A")

    def test_read_merged_ranges(self, adapter: ConcreteWriteOnly) -> None:
        with pytest.raises(NotImplementedError, match="write-only"):
            adapter.read_merged_ranges(None, "S")

    def test_read_conditional_formats(self, adapter: ConcreteWriteOnly) -> None:
        with pytest.raises(NotImplementedError, match="write-only"):
            adapter.read_conditional_formats(None, "S")

    def test_read_data_validations(self, adapter: ConcreteWriteOnly) -> None:
        with pytest.raises(NotImplementedError, match="write-only"):
            adapter.read_data_validations(None, "S")

    def test_read_hyperlinks(self, adapter: ConcreteWriteOnly) -> None:
        with pytest.raises(NotImplementedError, match="write-only"):
            adapter.read_hyperlinks(None, "S")

    def test_read_images(self, adapter: ConcreteWriteOnly) -> None:
        with pytest.raises(NotImplementedError, match="write-only"):
            adapter.read_images(None, "S")

    def test_read_pivot_tables(self, adapter: ConcreteWriteOnly) -> None:
        with pytest.raises(NotImplementedError, match="write-only"):
            adapter.read_pivot_tables(None, "S")

    def test_read_comments(self, adapter: ConcreteWriteOnly) -> None:
        with pytest.raises(NotImplementedError, match="write-only"):
            adapter.read_comments(None, "S")

    def test_read_freeze_panes(self, adapter: ConcreteWriteOnly) -> None:
        with pytest.raises(NotImplementedError, match="write-only"):
            adapter.read_freeze_panes(None, "S")


class TestAdapterCapabilities:
    def test_readonly_can_read(self) -> None:
        a = ConcreteReadOnly()
        assert a.can_read() is True
        assert a.can_write() is False

    def test_writeonly_can_write(self) -> None:
        a = ConcreteWriteOnly()
        assert a.can_read() is False
        assert a.can_write() is True

    def test_readonly_name(self) -> None:
        a = ConcreteReadOnly()
        assert a.name == "test-readonly"

    def test_writeonly_name(self) -> None:
        a = ConcreteWriteOnly()
        assert a.name == "test-writeonly"

    def test_readonly_output_extension(self) -> None:
        a = ConcreteReadOnly()
        assert a.output_extension == ".xlsx"

    def test_writeonly_supports_read_path(self) -> None:
        a = ConcreteWriteOnly()
        assert a.supports_read_path(Path("test.xlsx")) is True
        assert a.supports_read_path(Path("test.csv")) is False


def test_map_error_to_diagnostic() -> None:
    adapter = ConcreteReadOnly()
    diag = adapter.map_error_to_diagnostic(
        exc=NotImplementedError("unsupported"),
        feature="pivot_tables",
        operation=OperationType.READ,
        test_case_id="t1",
    )
    assert diag.category == DiagnosticCategory.UNSUPPORTED_FEATURE
    assert diag.location.feature == "pivot_tables"


def test_build_mismatch_diagnostic() -> None:
    adapter = ConcreteReadOnly()
    diag = adapter.build_mismatch_diagnostic(
        feature="cell_values",
        operation=OperationType.READ,
        test_case_id="t1",
        expected={"value": "x"},
        actual={"value": "y"},
    )
    assert diag.category == DiagnosticCategory.DATA_MISMATCH
    assert "expected" in diag.adapter_message
