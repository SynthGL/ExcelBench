"""Round-trip contracts for the WolfxlAdapter through the public WolfXL API.

Verified against WolfXL 2.0.1 (PyPI) and 2.1.0 (local wheel): every
covered operation behaves identically on both versions.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

wolfxl = pytest.importorskip("wolfxl", reason="wolfxl not installed")

from excelbench.harness.adapters import WolfxlAdapter  # noqa: E402
from excelbench.models import (  # noqa: E402
    BorderEdge,
    BorderInfo,
    BorderStyle,
    CellFormat,
    CellType,
    CellValue,
)


@pytest.fixture
def adapter() -> Any:
    assert WolfxlAdapter is not None
    return WolfxlAdapter()


def _string(value: str) -> CellValue:
    return CellValue(type=CellType.STRING, value=value)


def _number(value: float) -> CellValue:
    return CellValue(type=CellType.NUMBER, value=value)


def _build_basic_workbook(adapter: Any) -> Any:
    workbook = adapter.create_workbook()
    adapter.add_sheet(workbook, "Data")
    adapter.write_cell_value(workbook, "Data", "A1", _string("hello"))
    adapter.write_cell_value(workbook, "Data", "B1", _number(7))
    adapter.write_cell_value(workbook, "Data", "B2", _number(2.5))
    return workbook


def test_value_round_trip(adapter: Any, tmp_path: Path) -> None:
    workbook = _build_basic_workbook(adapter)
    path = tmp_path / "values.xlsx"
    adapter.save_workbook(workbook, path)

    reopened = adapter.open_workbook(path)
    try:
        assert adapter.get_sheet_names(reopened) == ["Data"]
        assert adapter.read_cell_value(reopened, "Data", "A1").value == "hello"
        assert adapter.read_cell_value(reopened, "Data", "B1").value == 7
        assert adapter.read_cell_value(reopened, "Data", "B2").value == 2.5
    finally:
        adapter.close_workbook(reopened)


def test_boolean_and_blank_round_trip(adapter: Any, tmp_path: Path) -> None:
    workbook = adapter.create_workbook()
    adapter.add_sheet(workbook, "S")
    adapter.write_cell_value(
        workbook, "S", "A1", CellValue(type=CellType.BOOLEAN, value=True)
    )
    adapter.write_cell_value(
        workbook, "S", "A2", CellValue(type=CellType.BLANK, value=None)
    )
    path = tmp_path / "bool.xlsx"
    adapter.save_workbook(workbook, path)

    reopened = adapter.open_workbook(path)
    try:
        assert adapter.read_cell_value(reopened, "S", "A1").value is True
        assert adapter.read_cell_value(reopened, "S", "A2").value is None
    finally:
        adapter.close_workbook(reopened)


def test_opened_workbook_tracks_its_source_path(adapter: Any, tmp_path: Path) -> None:
    workbook = _build_basic_workbook(adapter)
    path = tmp_path / "bound.xlsx"
    adapter.save_workbook(workbook, path)

    reopened = adapter.open_workbook(path)
    try:
        assert adapter._workbook_paths.get(id(reopened)) == path
    finally:
        adapter.close_workbook(reopened)


def test_format_border_dimension_merge_round_trip(adapter: Any, tmp_path: Path) -> None:
    workbook = adapter.create_workbook()
    adapter.add_sheet(workbook, "F")
    adapter.write_cell_value(workbook, "F", "A1", _number(1))
    adapter.write_cell_format(
        workbook,
        "F",
        "A1",
        CellFormat(bold=True, font_size=14.0, bg_color="#FFFF0000"),
    )
    adapter.write_cell_border(
        workbook,
        "F",
        "A1",
        BorderInfo(bottom=BorderEdge(style=BorderStyle.MEDIUM, color="#FF0000")),
    )
    adapter.set_row_height(workbook, "F", 1, 30.5)
    adapter.set_column_width(workbook, "F", "A", 42.0)
    adapter.merge_cells(workbook, "F", "A3:B4")
    path = tmp_path / "formatted.xlsx"
    adapter.save_workbook(workbook, path)

    reopened = adapter.open_workbook(path)
    try:
        cell_format = adapter.read_cell_format(reopened, "F", "A1")
        assert cell_format.bold is True
        assert cell_format.font_size == 14.0

        border = adapter.read_cell_border(reopened, "F", "A1")
        assert border.bottom is not None
        assert border.bottom.style == BorderStyle.MEDIUM

        assert adapter.read_row_height(reopened, "F", 1) == 30.5
        assert adapter.read_column_width(reopened, "F", "A") == 42.0
        assert "A3:B4" in adapter.read_merged_ranges(reopened, "F")
    finally:
        adapter.close_workbook(reopened)


def test_formula_text_round_trip(adapter: Any, tmp_path: Path) -> None:
    workbook = adapter.create_workbook()
    adapter.add_sheet(workbook, "C")
    adapter.write_cell_value(workbook, "C", "A1", _number(2))
    adapter.write_cell_value(workbook, "C", "A2", _number(3))
    adapter.write_cell_value(
        workbook,
        "C",
        "A3",
        CellValue(type=CellType.FORMULA, value=5, formula="=A1+A2"),
    )
    path = tmp_path / "formula.xlsx"
    adapter.save_workbook(workbook, path)

    reopened = adapter.open_workbook(path)
    try:
        formula_cell = adapter.read_cell_value(reopened, "C", "A3")
        assert formula_cell.formula == "=A1+A2"
    finally:
        adapter.close_workbook(reopened)


def test_structure_round_trip(adapter: Any, tmp_path: Path) -> None:
    workbook = adapter.create_workbook()
    adapter.add_sheet(workbook, "S")
    adapter.write_cell_value(workbook, "S", "B2", _number(42))
    adapter.set_freeze_panes(workbook, "S", {"mode": "freeze", "top_left_cell": "B2"})
    adapter.add_named_range(
        workbook, "S", {"name": "SingleCell", "refers_to": "S!$B$2"}
    )
    adapter.add_comment(
        workbook, "S", {"cell": "A2", "text": "note", "author": "tester"}
    )
    adapter.add_table(workbook, "S", {"name": "T1", "ref": "A1:A2"})
    path = tmp_path / "structure.xlsx"
    adapter.save_workbook(workbook, path)

    reopened = adapter.open_workbook(path)
    try:
        assert adapter.read_freeze_panes(reopened, "S") == {
            "mode": "freeze",
            "top_left_cell": "B2",
        }
        ranges = adapter.read_named_ranges(reopened, "S")
        assert any(item["name"] == "SingleCell" for item in ranges)
        comments = adapter.read_comments(reopened, "S")
        assert comments[0]["text"] == "note"
        tables = adapter.read_tables(reopened, "S")
        assert tables[0]["name"] == "T1"
    finally:
        adapter.close_workbook(reopened)


def test_hyperlink_round_trip(adapter: Any, tmp_path: Path) -> None:
    workbook = adapter.create_workbook()
    adapter.add_sheet(workbook, "L")
    adapter.write_cell_value(workbook, "L", "A1", _string("site"))
    adapter.add_hyperlink(
        workbook,
        "L",
        {
            "cell": "A1",
            "target": "https://example.com",
            "display": "Example",
            "internal": False,
        },
    )
    path = tmp_path / "links.xlsx"
    adapter.save_workbook(workbook, path)

    reopened = adapter.open_workbook(path)
    try:
        links = adapter.read_hyperlinks(reopened, "L")
        assert any(link.get("target") == "https://example.com" for link in links)
    finally:
        adapter.close_workbook(reopened)
