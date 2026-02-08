"""Smoke/integration tests for optional Rust backends.

These tests are designed to be safe in environments where the PyO3 extension
module is not built/installed.
"""

from __future__ import annotations

import importlib.util
import tempfile
from datetime import date, datetime
from pathlib import Path
from typing import Any

import pytest


def _enabled_backends(excelbench_rust: Any) -> set[str]:
    info = excelbench_rust.build_info()
    if isinstance(info, dict):
        enabled = info.get("enabled_backends")
        if isinstance(enabled, list):
            return {str(x) for x in enabled}
    return set()


def test_registry_works_without_excelbench_rust() -> None:
    """If the native extension isn't installed, adapter discovery must still work."""
    if importlib.util.find_spec("excelbench_rust") is not None:
        pytest.skip("excelbench_rust installed; this test targets no-extension environments")

    from excelbench.harness.adapters import get_all_adapters

    names = [a.name for a in get_all_adapters()]
    assert "calamine" not in names
    assert "rust_xlsxwriter" not in names
    assert "umya-spreadsheet" not in names


def test_rust_calamine_datetime_semantics() -> None:
    excelbench_rust = pytest.importorskip("excelbench_rust")
    enabled = _enabled_backends(excelbench_rust)
    if "calamine" not in enabled:
        pytest.skip("excelbench_rust compiled without calamine backend")

    from excelbench.harness.adapters.openpyxl_adapter import OpenpyxlAdapter
    from excelbench.harness.adapters.rust_calamine_adapter import RustCalamineAdapter
    from excelbench.models import CellType, CellValue

    f = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    path = Path(f.name)
    f.close()
    try:
        openpyxl = OpenpyxlAdapter()
        wb = openpyxl.create_workbook()
        openpyxl.add_sheet(wb, "S")
        openpyxl.write_cell_value(
            wb, "S", "A1", CellValue(type=CellType.DATE, value=date(2024, 6, 15))
        )
        openpyxl.write_cell_value(
            wb,
            "S",
            "A2",
            CellValue(type=CellType.DATETIME, value=datetime(2024, 6, 15, 10, 30, 0)),
        )
        openpyxl.save_workbook(wb, path)

        adapter = RustCalamineAdapter()
        wb2 = adapter.open_workbook(path)
        v1 = adapter.read_cell_value(wb2, "S", "A1")
        v2 = adapter.read_cell_value(wb2, "S", "A2")
        assert v1.type == CellType.DATE
        assert v2.type == CellType.DATETIME
    finally:
        path.unlink(missing_ok=True)


def test_rust_xlsxwriter_preserves_sheet_insertion_order() -> None:
    excelbench_rust = pytest.importorskip("excelbench_rust")
    enabled = _enabled_backends(excelbench_rust)
    if "rust_xlsxwriter" not in enabled:
        pytest.skip("excelbench_rust compiled without rust_xlsxwriter backend")

    from excelbench.harness.adapters.rust_xlsxwriter_adapter import RustXlsxWriterAdapter
    from excelbench.models import CellType, CellValue

    import openpyxl

    f = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    path = Path(f.name)
    f.close()
    try:
        adapter = RustXlsxWriterAdapter()
        wb = adapter.create_workbook()

        # Insert in non-alphabetical order to detect ordering bugs.
        adapter.add_sheet(wb, "Sheet2")
        adapter.add_sheet(wb, "Sheet1")
        adapter.write_cell_value(wb, "Sheet2", "A1", CellValue(type=CellType.STRING, value="s2"))
        adapter.write_cell_value(wb, "Sheet1", "A1", CellValue(type=CellType.STRING, value="s1"))

        adapter.save_workbook(wb, path)

        wb2 = openpyxl.load_workbook(str(path), data_only=False)
        assert wb2.sheetnames == ["Sheet2", "Sheet1"]
        wb2.close()
    finally:
        path.unlink(missing_ok=True)


def test_rust_xlsxwriter_error_tokens_write_expected_formula() -> None:
    excelbench_rust = pytest.importorskip("excelbench_rust")
    enabled = _enabled_backends(excelbench_rust)
    if "rust_xlsxwriter" not in enabled:
        pytest.skip("excelbench_rust compiled without rust_xlsxwriter backend")

    from excelbench.harness.adapters.rust_xlsxwriter_adapter import RustXlsxWriterAdapter
    from excelbench.models import CellType, CellValue

    import openpyxl

    f = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    path = Path(f.name)
    f.close()
    try:
        adapter = RustXlsxWriterAdapter()
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S")
        adapter.write_cell_value(wb, "S", "A1", CellValue(type=CellType.ERROR, value="#REF!"))
        adapter.save_workbook(wb, path)

        wb2 = openpyxl.load_workbook(str(path), data_only=False)
        ws = wb2["S"]
        # For non-mapped error tokens we write the literal token as a string.
        assert ws["A1"].value == "#REF!"
        wb2.close()
    finally:
        path.unlink(missing_ok=True)


def test_rust_xlsxwriter_writes_dates_as_excel_dates() -> None:
    excelbench_rust = pytest.importorskip("excelbench_rust")
    enabled = _enabled_backends(excelbench_rust)
    if "rust_xlsxwriter" not in enabled:
        pytest.skip("excelbench_rust compiled without rust_xlsxwriter backend")

    from excelbench.harness.adapters.rust_xlsxwriter_adapter import RustXlsxWriterAdapter
    from excelbench.models import CellType, CellValue

    import openpyxl

    f = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    path = Path(f.name)
    f.close()
    try:
        adapter = RustXlsxWriterAdapter()
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S")
        adapter.write_cell_value(
            wb, "S", "A1", CellValue(type=CellType.DATE, value=date(2024, 6, 15))
        )
        adapter.write_cell_value(
            wb,
            "S",
            "A2",
            CellValue(type=CellType.DATETIME, value=datetime(2024, 6, 15, 10, 30, 0)),
        )
        adapter.save_workbook(wb, path)

        wb2 = openpyxl.load_workbook(str(path), data_only=False)
        ws = wb2["S"]
        assert ws["A1"].value == datetime(2024, 6, 15, 0, 0, 0)
        assert ws["A2"].value == datetime(2024, 6, 15, 10, 30, 0)
        wb2.close()
    finally:
        path.unlink(missing_ok=True)


def test_umya_write_date_datetime_and_error_encodings() -> None:
    excelbench_rust = pytest.importorskip("excelbench_rust")
    enabled = _enabled_backends(excelbench_rust)
    if "umya-spreadsheet" not in enabled:
        pytest.skip("excelbench_rust compiled without umya backend")

    from excelbench.harness.adapters.umya_adapter import UmyaAdapter
    from excelbench.models import CellType, CellValue

    import openpyxl

    f = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    path = Path(f.name)
    f.close()
    try:
        adapter = UmyaAdapter()
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "S")
        adapter.write_cell_value(
            wb, "S", "A1", CellValue(type=CellType.DATE, value=date(2024, 6, 15))
        )
        adapter.write_cell_value(
            wb,
            "S",
            "A2",
            CellValue(type=CellType.DATETIME, value=datetime(2024, 6, 15, 10, 30, 0)),
        )
        adapter.write_cell_value(wb, "S", "A3", CellValue(type=CellType.ERROR, value="#DIV/0!"))
        adapter.write_cell_value(wb, "S", "A4", CellValue(type=CellType.ERROR, value="#REF!"))
        adapter.save_workbook(wb, path)

        wb2 = openpyxl.load_workbook(str(path), data_only=False)
        ws = wb2["S"]
        assert ws["A1"].value == datetime(2024, 6, 15, 0, 0, 0)
        assert ws["A2"].value == datetime(2024, 6, 15, 10, 30, 0)
        assert ws["A3"].data_type == "f"
        assert ws["A3"].value == "=1/0"
        assert ws["A4"].data_type == "s"
        assert ws["A4"].value == "#REF!"
        wb2.close()
    finally:
        path.unlink(missing_ok=True)


def test_umya_reads_openpyxl_dates_as_date_and_datetime() -> None:
    excelbench_rust = pytest.importorskip("excelbench_rust")
    enabled = _enabled_backends(excelbench_rust)
    if "umya-spreadsheet" not in enabled:
        pytest.skip("excelbench_rust compiled without umya backend")

    from excelbench.harness.adapters.openpyxl_adapter import OpenpyxlAdapter
    from excelbench.harness.adapters.umya_adapter import UmyaAdapter
    from excelbench.models import CellType, CellValue

    f = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    path = Path(f.name)
    f.close()
    try:
        openpyxl = OpenpyxlAdapter()
        wb = openpyxl.create_workbook()
        openpyxl.add_sheet(wb, "S")
        openpyxl.write_cell_value(
            wb, "S", "A1", CellValue(type=CellType.DATE, value=date(2024, 6, 15))
        )
        openpyxl.write_cell_value(
            wb,
            "S",
            "A2",
            CellValue(type=CellType.DATETIME, value=datetime(2024, 6, 15, 10, 30, 0)),
        )
        openpyxl.save_workbook(wb, path)

        adapter = UmyaAdapter()
        wb2 = adapter.open_workbook(path)
        v1 = adapter.read_cell_value(wb2, "S", "A1")
        v2 = adapter.read_cell_value(wb2, "S", "A2")
        assert v1.type == CellType.DATE
        assert v2.type == CellType.DATETIME
    finally:
        path.unlink(missing_ok=True)
