"""Tests for optional external spreadsheet oracle helpers."""

from __future__ import annotations

import shutil
import sys
from pathlib import Path
from zipfile import ZipFile

import pytest

from excelbench.harness.external_oracles import (
    ExternalOracleRequest,
    ExternalOracleTool,
    external_oracle_catalog,
    run_external_oracle,
)


def test_catalog_lists_planned_external_oracles() -> None:
    catalog = external_oracle_catalog()

    assert set(catalog) == {"excelize", "libreoffice", "apache-poi", "closedxml"}
    assert "pivots" in catalog["excelize"].capabilities
    assert "open_save_validate" in catalog["libreoffice"].capabilities


def test_repo_catalog_points_excelize_at_go_helper() -> None:
    repo_root = Path(__file__).resolve().parents[1]
    tool = external_oracle_catalog(repo_root=repo_root)["excelize"]

    assert tool.command == ("go", "run", ".")
    assert tool.cwd == repo_root / "tools" / "external-oracles" / "excelize"


def test_missing_oracle_helper_is_structured_skip() -> None:
    tool = ExternalOracleTool(
        name="missing",
        command=("excelbench-helper-that-does-not-exist",),
        language="test",
        homepage="https://example.invalid",
        capabilities=frozenset({"write"}),
    )

    result = run_external_oracle(
        tool,
        ExternalOracleRequest(
            fixture_id="fixture",
            operation="write_fixture",
            payload={"cells": []},
        ),
    )

    assert result.skipped is True
    assert result.passed is False
    assert result.returncode is None
    assert result.notes is not None
    assert "not found" in result.notes


def test_oracle_request_serializes_paths() -> None:
    request = ExternalOracleRequest(
        fixture_id="pivot-cache",
        operation="open_save_validate",
        payload={"feature": "pivot_tables"},
        input_path=Path("input.xlsx"),
        output_path=Path("output.xlsx"),
    )

    assert request.to_json_dict() == {
        "fixture_id": "pivot-cache",
        "operation": "open_save_validate",
        "payload": {"feature": "pivot_tables"},
        "input_path": "input.xlsx",
        "output_path": "output.xlsx",
    }


def test_successful_oracle_helper_parses_json_stdout() -> None:
    tool = ExternalOracleTool(
        name="python-json-helper",
        command=(
            sys.executable,
            "-c",
            (
                "import json, sys; "
                "request = json.load(sys.stdin); "
                "print(json.dumps({'fixture_id': request['fixture_id'], 'ok': True}))"
            ),
        ),
        language="python",
        homepage="https://example.invalid",
        capabilities=frozenset({"write"}),
    )

    result = run_external_oracle(
        tool,
        ExternalOracleRequest(
            fixture_id="rich-text",
            operation="write_fixture",
            payload={"cells": [{"cell": "A1", "value": "Hello"}]},
        ),
    )

    assert result.passed is True
    assert result.skipped is False
    assert result.returncode == 0
    assert result.payload == {"fixture_id": "rich-text", "ok": True}


def test_nonzero_oracle_helper_is_failure() -> None:
    tool = ExternalOracleTool(
        name="python-failing-helper",
        command=(sys.executable, "-c", "import sys; sys.stderr.write('boom'); sys.exit(2)"),
        language="python",
        homepage="https://example.invalid",
        capabilities=frozenset({"write"}),
    )

    result = run_external_oracle(
        tool,
        ExternalOracleRequest(
            fixture_id="broken",
            operation="write_fixture",
            payload={},
        ),
    )

    assert result.passed is False
    assert result.skipped is False
    assert result.returncode == 2
    assert result.stderr == "boom"


def test_invalid_json_stdout_is_failure_payload() -> None:
    tool = ExternalOracleTool(
        name="python-invalid-json-helper",
        command=(sys.executable, "-c", "print('not json')"),
        language="python",
        homepage="https://example.invalid",
        capabilities=frozenset({"write"}),
    )

    result = run_external_oracle(
        tool,
        ExternalOracleRequest(
            fixture_id="bad-json",
            operation="write_fixture",
            payload={},
        ),
    )

    assert result.passed is False
    assert result.payload["error"] == "invalid_json_stdout"


@pytest.mark.skipif(shutil.which("go") is None, reason="Go is required for Excelize oracle smoke")
def test_excelize_go_helper_writes_advanced_fixture(tmp_path: Path) -> None:
    repo_root = Path(__file__).resolve().parents[1]
    output_path = tmp_path / "excelize-smoke.xlsx"
    tool = external_oracle_catalog(repo_root=repo_root)["excelize"]

    result = run_external_oracle(
        tool,
        ExternalOracleRequest(
            fixture_id="excelize-smoke",
            operation="write_fixture",
            output_path=output_path,
            payload={
                "sheets": [{"name": "Data"}, {"name": "Pivot"}],
                "cells": [
                    {"sheet": "Data", "cell": "A1", "value": "Region"},
                    {"sheet": "Data", "cell": "B1", "value": "Product"},
                    {"sheet": "Data", "cell": "C1", "value": "Sales"},
                    {"sheet": "Data", "cell": "A2", "value": "West"},
                    {"sheet": "Data", "cell": "B2", "value": "Widgets"},
                    {"sheet": "Data", "cell": "C2", "value": 120},
                    {"sheet": "Data", "cell": "A3", "value": "East"},
                    {"sheet": "Data", "cell": "B3", "value": "Services"},
                    {"sheet": "Data", "cell": "C3", "value": 95},
                ],
                "tables": [{"sheet": "Data", "range": "A1:C3", "name": "SalesTable"}],
                "conditional_formats": [
                    {"sheet": "Data", "range": "C2:C3", "type": "3_color_scale"},
                    {"sheet": "Data", "range": "C2:C3", "type": "data_bar"},
                    {
                        "sheet": "Data",
                        "range": "C2:C3",
                        "type": "icon_set",
                        "icon_style": "3TrafficLights1",
                    },
                ],
                "charts": [
                    {
                        "sheet": "Data",
                        "cell": "E2",
                        "type": "col",
                        "title": "Sales",
                        "categories": "Data!$A$2:$A$3",
                        "values": "Data!$C$2:$C$3",
                    }
                ],
                "pivots": [
                    {
                        "data_range": "Data!A1:C3",
                        "range": "Pivot!A3:E10",
                        "name": "SalesPivot",
                        "rows": [{"name": "Region"}],
                        "data": [{"name": "Sales", "subtotal": "Sum"}],
                    }
                ],
                "slicers": [
                    {
                        "sheet": "Data",
                        "name": "Region",
                        "cell": "E15",
                        "table_sheet": "Data",
                        "table_name": "SalesTable",
                    }
                ],
                "pictures": [{"sheet": "Data", "cell": "H2", "name": "Pixel"}],
            },
        ),
        timeout_seconds=120,
    )

    assert result.passed is True, result
    assert output_path.exists()
    assert result.payload["counts"]["pivots"] == 1
    with ZipFile(output_path) as workbook:
        names = set(workbook.namelist())
    assert "xl/tables/table1.xml" in names
    assert "xl/pivotTables/pivotTable1.xml" in names
    assert "xl/slicerCaches/slicerCache1.xml" in names
    assert "xl/charts/chart1.xml" in names
