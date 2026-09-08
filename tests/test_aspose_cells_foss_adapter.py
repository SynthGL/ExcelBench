"""Tests for the aspose-cells-foss adapter using its installed implementation."""

from __future__ import annotations

from pathlib import Path

import pytest

from excelbench.harness.adapters.aspose_cells_foss_adapter import AsposeCellsFossAdapter
from excelbench.harness.adapters.base import UnsupportedAdapterOperationError
from excelbench.models import (
    BorderEdge,
    BorderInfo,
    BorderStyle,
    CellFormat,
    CellType,
    CellValue,
)


@pytest.fixture
def adapter() -> AsposeCellsFossAdapter:
    return AsposeCellsFossAdapter()


def test_values_formats_and_dimensions_round_trip(
    adapter: AsposeCellsFossAdapter, tmp_path: Path
) -> None:
    """Values, formula text, formatting, borders, and dimensions survive an XLSX round trip."""
    path = tmp_path / "aspose-cells-foss.xlsx"
    workbook = adapter.create_workbook()
    adapter.add_sheet(workbook, "Data")
    adapter.write_cell_value(
        workbook, "Data", "A1", CellValue(type=CellType.STRING, value="hello")
    )
    adapter.write_cell_value(
        workbook, "Data", "B1", CellValue(type=CellType.NUMBER, value=42.5)
    )
    adapter.write_cell_value(
        workbook,
        "Data",
        "C1",
        CellValue(type=CellType.FORMULA, value=None, formula="=B1*2"),
    )
    adapter.write_cell_format(
        workbook,
        "Data",
        "A1",
        CellFormat(
            bold=True,
            italic=True,
            underline="single",
            strikethrough=True,
            font_name="Arial",
            font_size=16,
            font_color="#FF0000",
            bg_color="#FFFF00",
            number_format="0.00",
            h_align="center",
            v_align="top",
            wrap=True,
            rotation=45,
            indent=2,
        ),
    )
    adapter.write_cell_border(
        workbook,
        "Data",
        "A1",
        BorderInfo(
            top=BorderEdge(style=BorderStyle.THIN, color="#FF0000"),
            bottom=BorderEdge(style=BorderStyle.DOUBLE, color="#00FF00"),
        ),
    )
    adapter.set_row_height(workbook, "Data", 1, 30)
    adapter.set_column_width(workbook, "Data", "A", 20)
    adapter.save_workbook(workbook, path)

    reopened = adapter.open_workbook(path)
    assert adapter.get_sheet_names(reopened) == ["Data"]
    assert adapter.read_cell_value(reopened, "Data", "A1") == CellValue(
        type=CellType.STRING, value="hello"
    )
    assert adapter.read_cell_value(reopened, "Data", "B1") == CellValue(
        type=CellType.NUMBER, value=42.5
    )
    formula = adapter.read_cell_value(reopened, "Data", "C1")
    assert formula.type == CellType.FORMULA
    assert formula.formula == "=B1*2"
    cell_format = adapter.read_cell_format(reopened, "Data", "A1")
    assert cell_format.bold is True
    assert cell_format.italic is True
    assert cell_format.underline == "single"
    assert cell_format.strikethrough is True
    assert cell_format.font_name == "Arial"
    assert cell_format.font_size == 16
    assert cell_format.font_color == "#FF0000"
    assert cell_format.bg_color == "#FFFF00"
    assert cell_format.number_format == "0.00"
    assert cell_format.h_align == "center"
    assert cell_format.v_align == "top"
    assert cell_format.wrap is True
    assert cell_format.rotation == 45
    assert cell_format.indent == 2
    border = adapter.read_cell_border(reopened, "Data", "A1")
    assert border.top == BorderEdge(style=BorderStyle.THIN, color="#FF0000")
    assert border.bottom == BorderEdge(style=BorderStyle.DOUBLE, color="#00FF00")
    assert adapter.read_row_height(reopened, "Data", 1) == 30
    assert adapter.read_column_width(reopened, "Data", "A") == 20
    adapter.close_workbook(reopened)


def test_pivot_operations_raise_structured_unsupported_error(
    adapter: AsposeCellsFossAdapter,
) -> None:
    """The adapter reports its unsupported pivot capability without a silent no-op."""
    workbook = adapter.create_workbook()
    adapter.add_sheet(workbook, "Data")

    with pytest.raises(UnsupportedAdapterOperationError) as error:
        adapter.add_pivot_table(workbook, "Data", {})

    assert error.value.adapter == "aspose-cells-foss"
    assert error.value.operation == "add_pivot_table"
    assert "no pivot-table API" in error.value.reason
