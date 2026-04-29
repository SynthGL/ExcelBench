"""Tests for optional external spreadsheet oracle helpers."""

from __future__ import annotations

import shutil
import subprocess
import sys
from importlib.util import module_from_spec, spec_from_file_location
from pathlib import Path
from zipfile import ZipFile

import pytest
from openpyxl import Workbook

from excelbench.harness.external_oracles import (
    ExternalOracleRequest,
    ExternalOracleTool,
    external_oracle_catalog,
    run_external_oracle,
)


def _missing_dotnet_runtime(prefix: str) -> bool:
    if shutil.which("dotnet") is None:
        return True
    completed = subprocess.run(
        ["dotnet", "--list-runtimes"],
        text=True,
        capture_output=True,
        check=False,
    )
    return completed.returncode != 0 or f"Microsoft.NETCore.App {prefix}" not in completed.stdout


def test_catalog_lists_planned_external_oracles() -> None:
    catalog = external_oracle_catalog()

    assert set(catalog) == {
        "excelize",
        "libreoffice",
        "apache-poi",
        "exceljs",
        "closedxml",
        "npoi",
    }
    assert "pivots" in catalog["excelize"].capabilities
    assert "open_save_validate" in catalog["libreoffice"].capabilities
    assert "data_validations" in catalog["apache-poi"].capabilities
    assert "data_validations" in catalog["exceljs"].capabilities
    assert "rich_text" in catalog["npoi"].capabilities


def test_repo_catalog_points_excelize_at_go_helper() -> None:
    repo_root = Path(__file__).resolve().parents[1]
    catalog = external_oracle_catalog(repo_root=repo_root)
    tool = catalog["excelize"]

    assert tool.command == ("go", "run", ".")
    assert tool.cwd == repo_root / "tools" / "external-oracles" / "excelize"
    assert catalog["libreoffice"].command[0] == sys.executable
    assert catalog["libreoffice"].command[1].endswith("libreoffice_oracle.py")
    assert catalog["apache-poi"].command[0] == sys.executable
    assert catalog["apache-poi"].command[1].endswith("poi_oracle.py")
    assert catalog["apache-poi"].cwd == repo_root / "tools" / "external-oracles" / "apache-poi"
    assert catalog["apache-poi"].required_paths
    assert catalog["exceljs"].command[0] == "node"
    assert catalog["exceljs"].command[1].endswith("exceljs-oracle.cjs")
    assert catalog["exceljs"].cwd == repo_root / "tools" / "external-oracles" / "exceljs"
    assert catalog["exceljs"].required_paths
    assert catalog["closedxml"].command[0] == "dotnet"
    assert any(part.endswith("closedxml-oracle.csproj") for part in catalog["closedxml"].command)
    assert catalog["npoi"].command[0] == "dotnet"
    assert any(part.endswith("npoi-oracle.csproj") for part in catalog["npoi"].command)


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


def test_libreoffice_helper_rejects_blank_input_path() -> None:
    repo_root = Path(__file__).resolve().parents[1]
    helper_path = repo_root / "tools" / "external-oracles" / "libreoffice" / "libreoffice_oracle.py"
    spec = spec_from_file_location("excelbench_libreoffice_oracle", helper_path)
    assert spec is not None
    assert spec.loader is not None
    module = module_from_spec(spec)
    spec.loader.exec_module(module)

    response, exit_code = module.run_conversion(
        soffice="/usr/bin/false",
        request={"operation": "open_save_validate", "input_path": "   "},
        extension="xlsx",
        filter_name="Calc Office Open XML",
    )

    assert exit_code == 1
    assert response["error"] == "missing_input_path"


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


@pytest.mark.skipif(
    _missing_dotnet_runtime("8."),
    reason=".NET 8 runtime is required for ClosedXML",
)
def test_closedxml_dotnet_helper_writes_pivot_fixture(tmp_path: Path) -> None:
    repo_root = Path(__file__).resolve().parents[1]
    output_path = tmp_path / "closedxml-smoke.xlsx"
    tool = external_oracle_catalog(repo_root=repo_root)["closedxml"]

    result = run_external_oracle(
        tool,
        ExternalOracleRequest(
            fixture_id="closedxml-smoke",
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
                    {"sheet": "Data", "cell": "A4", "value": "West"},
                    {"sheet": "Data", "cell": "B4", "value": "Services"},
                    {"sheet": "Data", "cell": "C4", "value": 140},
                ],
                "tables": [{"sheet": "Data", "range": "A1:C4", "name": "ClosedXmlSales"}],
                "conditional_formats": [
                    {"sheet": "Data", "range": "C2:C4", "type": "3_color_scale"},
                    {"sheet": "Data", "range": "C2:C4", "type": "data_bar"},
                ],
                "rich_text": [
                    {
                        "sheet": "Data",
                        "cell": "E1",
                        "runs": [
                            {"text": "Review ", "bold": True, "font_color": "#C00000"},
                            {"text": "note", "italic": True},
                        ],
                    }
                ],
                "comments": [
                    {
                        "sheet": "Data",
                        "cell": "C2",
                        "text": "ClosedXML comment smoke.",
                        "author": "ExcelBench",
                    }
                ],
                "protection": [{"sheet": "Data", "password": "audit"}],
                "pivots": [
                    {
                        "data_range": "Data!A1:C4",
                        "cell": "Pivot!A3",
                        "name": "ClosedXmlPivot",
                        "rows": [{"name": "Region"}],
                        "columns": [{"name": "Product"}],
                        "data": [{"name": "Sales"}],
                    }
                ],
            },
        ),
        timeout_seconds=180,
    )

    assert result.passed is True, result
    assert output_path.exists()
    assert result.payload["counts"]["pivots"] == 1
    assert result.payload["counts"]["comments"] == 1
    assert result.payload["counts"]["rich_text"] == 1
    assert result.payload["counts"]["protected_sheets"] == 1
    with ZipFile(output_path) as workbook:
        names = set(workbook.namelist())
        sheet_xml = workbook.read("xl/worksheets/sheet1.xml").decode()
        shared_strings_xml = workbook.read("xl/sharedStrings.xml").decode()
    assert "xl/tables/table1.xml" in names
    assert "xl/comments1.xml" in names
    assert "xl/drawings/vmldrawing.vml" in names
    assert any(part.endswith("pivotTables/pivotTable.xml") for part in names)
    assert any(part.endswith("pivotCache/pivotCacheDefinition1.xml") for part in names)
    assert "conditionalFormatting" in sheet_xml
    assert "sheetProtection" in sheet_xml
    assert ":r>" in shared_strings_xml


@pytest.mark.skipif(_missing_dotnet_runtime("8."), reason=".NET 8 runtime is required for NPOI")
def test_npoi_dotnet_helper_writes_formula_comment_fixture(tmp_path: Path) -> None:
    repo_root = Path(__file__).resolve().parents[1]
    output_path = tmp_path / "npoi-smoke.xlsx"
    tool = external_oracle_catalog(repo_root=repo_root)["npoi"]

    result = run_external_oracle(
        tool,
        ExternalOracleRequest(
            fixture_id="npoi-smoke",
            operation="write_fixture",
            output_path=output_path,
            payload={
                "sheets": [{"name": "NPOI"}],
                "cells": [
                    {"sheet": "NPOI", "cell": "A1", "value": "Account"},
                    {"sheet": "NPOI", "cell": "B1", "value": "Amount"},
                    {"sheet": "NPOI", "cell": "A2", "value": "Revenue"},
                    {"sheet": "NPOI", "cell": "B2", "value": 1250},
                    {"sheet": "NPOI", "cell": "A3", "value": "COGS"},
                    {"sheet": "NPOI", "cell": "B3", "value": -400},
                    {"sheet": "NPOI", "cell": "A4", "value": "Gross profit"},
                    {
                        "sheet": "NPOI",
                        "cell": "B4",
                        "type": "formula",
                        "formula": "SUM(B2:B3)",
                    },
                    {"sheet": "NPOI", "cell": "D1", "value": "Merged review header"},
                ],
                "rich_text": [
                    {
                        "sheet": "NPOI",
                        "cell": "D3",
                        "runs": [
                            {"text": "NPOI ", "bold": True},
                            {"text": "rich text", "italic": True},
                        ],
                    }
                ],
                "comments": [
                    {
                        "sheet": "NPOI",
                        "cell": "B4",
                        "text": "Formula result should preserve calc metadata.",
                        "author": "NPOI Oracle",
                    }
                ],
                "merged_ranges": [{"sheet": "NPOI", "range": "D1:F1"}],
                "protection": [{"sheet": "NPOI", "password": "audit"}],
            },
        ),
        timeout_seconds=180,
    )

    assert result.passed is True, result
    assert output_path.exists()
    assert result.payload["counts"]["formulas"] == 1
    assert result.payload["counts"]["comments"] == 1
    assert result.payload["counts"]["rich_text"] == 1
    assert result.payload["counts"]["merged_ranges"] == 1
    assert result.payload["counts"]["protected_sheets"] == 1
    with ZipFile(output_path) as workbook:
        names = set(workbook.namelist())
        sheet_xml = workbook.read("xl/worksheets/sheet1.xml").decode()
        shared_strings_xml = workbook.read("xl/sharedStrings.xml").decode()
    assert "xl/comments1.xml" in names
    assert "xl/drawings/vmlDrawing1.vml" in names
    assert "sheetProtection" in sheet_xml
    assert "mergeCell" in sheet_xml
    assert "<f>SUM(B2:B3)</f>" in sheet_xml
    assert "<r>" in shared_strings_xml


@pytest.mark.skipif(
    not (
        shutil.which("node")
        and (
            Path(__file__).resolve().parents[1]
            / "tools"
            / "external-oracles"
            / "exceljs"
            / "node_modules"
            / "exceljs"
            / "package.json"
        ).exists()
    ),
    reason="Node and npm-installed ExcelJS dependencies are required for ExcelJS",
)
def test_exceljs_node_helper_writes_table_validation_fixture(tmp_path: Path) -> None:
    repo_root = Path(__file__).resolve().parents[1]
    output_path = tmp_path / "exceljs-smoke.xlsx"
    tool = external_oracle_catalog(repo_root=repo_root)["exceljs"]

    result = run_external_oracle(
        tool,
        ExternalOracleRequest(
            fixture_id="exceljs-smoke",
            operation="write_fixture",
            output_path=output_path,
            payload={
                "sheets": [{"name": "ExcelJS", "freeze_panes": {"x_split": 1, "y_split": 1}}],
                "cells": [
                    {"sheet": "ExcelJS", "cell": "A1", "value": "Metric"},
                    {"sheet": "ExcelJS", "cell": "B1", "value": "Value"},
                    {"sheet": "ExcelJS", "cell": "A2", "value": "Revenue"},
                    {"sheet": "ExcelJS", "cell": "B2", "value": 1200},
                    {"sheet": "ExcelJS", "cell": "A3", "value": "COGS"},
                    {"sheet": "ExcelJS", "cell": "B3", "value": -450},
                    {
                        "sheet": "ExcelJS",
                        "cell": "B4",
                        "type": "formula",
                        "formula": "SUM(B2:B3)",
                        "result": 750,
                    },
                ],
                "rich_text": [
                    {
                        "sheet": "ExcelJS",
                        "cell": "D2",
                        "runs": [
                            {"text": "ExcelJS ", "bold": True},
                            {"text": "rich text", "italic": True},
                        ],
                    }
                ],
                "comments": [{"sheet": "ExcelJS", "cell": "B4", "text": "Formula comment."}],
                "hyperlinks": [
                    {
                        "sheet": "ExcelJS",
                        "cell": "D4",
                        "text": "ExcelJS project",
                        "url": "https://github.com/exceljs/exceljs",
                    }
                ],
                "merged_ranges": [{"sheet": "ExcelJS", "range": "D1:F1"}],
                "data_validations": [
                    {
                        "sheet": "ExcelJS",
                        "cell": "C2",
                        "type": "list",
                        "formulae": ['"Open,Closed"'],
                    }
                ],
                "tables": [
                    {
                        "sheet": "ExcelJS",
                        "name": "ExcelJsTable",
                        "ref": "F1:G3",
                        "columns": [{"name": "Item"}, {"name": "Status"}],
                        "rows": [["Revenue", "Open"], ["COGS", "Closed"]],
                    }
                ],
                "images": [{"sheet": "ExcelJS", "range": "D6:E8"}],
                "protection": [{"sheet": "ExcelJS", "password": "audit"}],
            },
        ),
        timeout_seconds=180,
    )

    assert result.passed is True, result
    assert output_path.exists()
    assert result.payload["counts"]["formulas"] == 1
    assert result.payload["counts"]["comments"] == 1
    assert result.payload["counts"]["rich_text"] == 1
    assert result.payload["counts"]["hyperlinks"] == 1
    assert result.payload["counts"]["tables"] == 1
    assert result.payload["counts"]["data_validations"] == 1
    assert result.payload["counts"]["images"] == 1
    with ZipFile(output_path) as workbook:
        names = set(workbook.namelist())
        sheet_xml = workbook.read("xl/worksheets/sheet1.xml").decode()
        shared_strings_xml = workbook.read("xl/sharedStrings.xml").decode()
    assert "xl/tables/table1.xml" in names
    assert "xl/comments1.xml" in names
    assert "xl/drawings/vmlDrawing1.vml" in names
    assert "xl/drawings/drawing1.xml" in names
    assert "xl/media/image1.png" in names
    assert "dataValidations" in sheet_xml
    assert "sheetProtection" in sheet_xml
    assert "mergeCell" in sheet_xml
    assert "<f>SUM(B2:B3)</f>" in sheet_xml
    assert "<r>" in shared_strings_xml


@pytest.mark.skipif(
    not (
        (
            Path("/opt/homebrew/opt/openjdk/bin/java").exists()
            or shutil.which("java")
        )
        and (
            Path(__file__).resolve().parents[1]
            / "tools"
            / "external-oracles"
            / "apache-poi"
            / "build"
            / "classes"
            / "PoiOracle.class"
        ).exists()
    ),
    reason="Java and built Apache POI helper classes are required for Apache POI",
)
def test_apache_poi_helper_writes_table_validation_fixture(tmp_path: Path) -> None:
    repo_root = Path(__file__).resolve().parents[1]
    output_path = tmp_path / "apache-poi-smoke.xlsx"
    tool = external_oracle_catalog(repo_root=repo_root)["apache-poi"]

    result = run_external_oracle(
        tool,
        ExternalOracleRequest(
            fixture_id="apache-poi-smoke",
            operation="write_fixture",
            output_path=output_path,
            payload={},
        ),
        timeout_seconds=180,
    )

    assert result.passed is True, result
    assert output_path.exists()
    assert result.payload["counts"]["formulas"] == 1
    assert result.payload["counts"]["comments"] == 1
    assert result.payload["counts"]["rich_text"] == 1
    assert result.payload["counts"]["hyperlinks"] == 1
    assert result.payload["counts"]["tables"] == 1
    assert result.payload["counts"]["data_validations"] == 1
    assert result.payload["counts"]["images"] == 1
    with ZipFile(output_path) as workbook:
        names = set(workbook.namelist())
        sheet_xml = workbook.read("xl/worksheets/sheet1.xml").decode()
        shared_strings_xml = workbook.read("xl/sharedStrings.xml").decode()
    assert "xl/tables/table1.xml" in names
    assert "xl/comments1.xml" in names
    assert "xl/drawings/vmlDrawing0.vml" in names
    assert "xl/drawings/drawing1.xml" in names
    assert "xl/media/image1.png" in names
    assert "dataValidations" in sheet_xml
    assert "sheetProtection" in sheet_xml
    assert "mergeCell" in sheet_xml
    assert "<f>SUM(B2:B3)</f>" in sheet_xml
    assert "<r>" in shared_strings_xml


def test_libreoffice_helper_renders_pdf_or_skips(tmp_path: Path) -> None:
    repo_root = Path(__file__).resolve().parents[1]
    input_path = tmp_path / "input.xlsx"
    output_path = tmp_path / "output.pdf"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet["A1"] = "Hello"
    sheet["B1"] = 42
    workbook.save(input_path)

    result = run_external_oracle(
        external_oracle_catalog(repo_root=repo_root)["libreoffice"],
        ExternalOracleRequest(
            fixture_id="libreoffice-render-smoke",
            operation="render_validate",
            input_path=input_path,
            output_path=output_path,
            payload={},
        ),
        timeout_seconds=180,
    )

    if result.skipped:
        assert result.notes == "LibreOffice executable not found"
        return
    assert result.passed is True, result
    assert output_path.exists()
    assert result.payload["bytes"] > 0
