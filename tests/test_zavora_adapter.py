"""Pure-Python payload tests for the Zavora write adapter."""

from __future__ import annotations

from pathlib import Path

import pytest

from excelbench.harness.adapters.zavora_adapter import ZavoraAdapter
from excelbench.harness.external_oracles import (
    ExternalOracleRequest,
    ExternalOracleResult,
)
from excelbench.models import CellType, CellValue


@pytest.fixture
def zavora() -> ZavoraAdapter:
    return ZavoraAdapter()


def test_info_declares_write_only_capability(zavora: ZavoraAdapter) -> None:
    assert zavora.info.name == "zavora-xlsx"
    assert zavora.info.version == "0.1.2"
    assert zavora.info.language == "rust"
    assert zavora.info.capabilities == {"write"}
    assert zavora.can_write()
    assert not zavora.can_read()


def test_cell_value_and_formula_entries(zavora: ZavoraAdapter) -> None:
    workbook = zavora.create_workbook()
    zavora.write_cell_value(
        workbook, "Data", "A1", CellValue(type=CellType.NUMBER, value=42)
    )
    zavora.write_cell_value(
        workbook,
        "Data",
        "B1",
        CellValue(type=CellType.FORMULA, value="=A1*2", formula="=A1*2"),
    )

    assert workbook["sheets"] == [{"name": "Data"}]
    assert workbook["cells"] == [
        {"sheet": "Data", "cell": "A1", "type": "number", "value": 42},
        {
            "sheet": "Data",
            "cell": "B1",
            "type": "formula",
            "formula": "=A1*2",
            "value": "=A1*2",
        },
    ]


def test_table_merge_and_freeze_payloads(zavora: ZavoraAdapter) -> None:
    workbook = zavora.create_workbook()
    zavora.merge_cells(workbook, "Data", "A1:B1")
    zavora.set_freeze_panes(
        workbook, "Data", {"freeze": {"mode": "freeze", "y_split": 1}}
    )
    zavora.add_table(
        workbook,
        "Data",
        {"table": {"name": "Items", "ref": "A1:B3", "columns": ["Name", "Amount"]}},
    )

    assert workbook["merges"] == [{"sheet": "Data", "range": "A1:B1"}]
    assert workbook["panes"] == [{"sheet": "Data", "mode": "freeze", "y_split": 1}]
    assert workbook["tables"] == [
        {
            "sheet": "Data",
            "range": "A1:B3",
            "name": "Items",
            "style": None,
            "show_header_row": True,
            "show_row_stripes": True,
            "totals_row": False,
            "columns": ["Name", "Amount"],
            "autofilter": False,
        }
    ]


def test_save_invokes_write_fixture_request(
    zavora: ZavoraAdapter, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    captured: dict[str, object] = {}

    def fake_run(
        tool: object, request: object, **kwargs: object
    ) -> ExternalOracleResult:
        captured["tool"] = tool
        captured["request"] = request
        captured["kwargs"] = kwargs
        return ExternalOracleResult(
            tool_name="zavora",
            passed=True,
            skipped=False,
            returncode=0,
            stdout='{"ok": true}',
            stderr="",
            payload={"ok": True},
        )

    monkeypatch.setattr(
        "excelbench.harness.adapters.zavora_adapter.run_external_oracle", fake_run
    )
    output_path = tmp_path / "result.xlsx"
    workbook = zavora.create_workbook()
    zavora.write_cell_value(
        workbook, "Data", "A1", CellValue(type=CellType.STRING, value="hello")
    )

    zavora.save_workbook(workbook, output_path)

    request = captured["request"]
    assert isinstance(request, ExternalOracleRequest)
    assert request.fixture_id == "zavora-adapter"
    assert request.operation == "write_fixture"
    assert request.output_path == output_path
    assert request.payload is workbook
    assert captured["kwargs"] == {"timeout_seconds": 300}
