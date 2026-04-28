"""Tests for local external-oracle fixture pack definitions."""

from __future__ import annotations

import os
import shutil
from pathlib import Path
from zipfile import ZipFile

import pytest

from excelbench.harness.external_fixture_pack import (
    apache_poi_fixture_specs,
    closedxml_fixture_specs,
    excelize_fixture_specs,
    exceljs_fixture_specs,
    external_fixture_specs,
    generate_external_fixture_pack,
    npoi_fixture_specs,
)
from excelbench.harness.external_wolfxl_validation import (
    _run_readback_probes,
    validate_wolfxl_external_fixture_pack,
)


def test_excelize_fixture_specs_are_stable() -> None:
    specs = excelize_fixture_specs()

    assert [spec.fixture_id for spec in specs] == [
        "excelize_sales_pivot_slicer_chart",
        "excelize_chart_points_formula_cf",
    ]
    assert all(spec.tool == "excelize" for spec in specs)
    assert "xl/pivotTables/pivotTable1.xml" in specs[0].expected_parts
    assert "xl/charts/chart1.xml" in specs[1].expected_parts
    assert any(probe["kind"] == "conditional_formatting" for probe in specs[0].readback_probes)
    assert any(probe["kind"] == "zip_contains" for probe in specs[0].readback_probes)
    assert any(probe["kind"] == "cell_formula" for probe in specs[1].readback_probes)


def test_closedxml_fixture_specs_are_stable() -> None:
    specs = closedxml_fixture_specs()

    assert [spec.fixture_id for spec in specs] == [
        "closedxml_pivot_cf_table",
        "closedxml_rich_comment_protection",
    ]
    assert all(spec.tool == "closedxml" for spec in specs)
    assert "xl/pivotTables/pivotTable.xml" in specs[0].expected_parts
    assert "pivotCache/pivotCacheDefinition1.xml" in specs[0].expected_parts
    assert any(probe["kind"] == "conditional_formatting" for probe in specs[0].readback_probes)
    assert any(probe["kind"] == "table_metadata" for probe in specs[0].readback_probes)
    assert "xl/comments1.xml" in specs[1].expected_parts
    assert "xl/drawings/vmldrawing.vml" in specs[1].expected_parts
    assert any(probe["kind"] == "comment_text" for probe in specs[1].readback_probes)
    assert any(probe["kind"] == "sheet_protection" for probe in specs[1].readback_probes)
    assert any(probe["kind"] == "rich_text_runs" for probe in specs[1].readback_probes)
    assert [spec.fixture_id for spec in external_fixture_specs()] == [
        "excelize_sales_pivot_slicer_chart",
        "excelize_chart_points_formula_cf",
        "closedxml_pivot_cf_table",
        "closedxml_rich_comment_protection",
        "npoi_formula_comment_merge_protection",
        "exceljs_table_validation_image_comment",
        "apache_poi_table_validation_image_comment",
    ]


def test_apache_poi_fixture_specs_are_stable() -> None:
    specs = apache_poi_fixture_specs()

    assert [spec.fixture_id for spec in specs] == ["apache_poi_table_validation_image_comment"]
    assert all(spec.tool == "apache-poi" for spec in specs)
    assert "xl/workbook.xml" in specs[0].expected_parts
    assert "xl/tables/table1.xml" in specs[0].expected_parts
    assert "xl/drawings/vmlDrawing0.vml" in specs[0].expected_parts
    assert any(probe["kind"] == "data_validation" for probe in specs[0].readback_probes)
    assert any(probe["kind"] == "hyperlink_target" for probe in specs[0].readback_probes)
    assert any(probe["kind"] == "merged_range" for probe in specs[0].readback_probes)
    assert any(probe["kind"] == "relationship_target" for probe in specs[0].readback_probes)
    assert any(probe["kind"] == "sheet_protection" for probe in specs[0].readback_probes)
    assert any(probe["kind"] == "rich_text_runs" for probe in specs[0].readback_probes)
    assert any(probe["kind"] == "workbook_protection" for probe in specs[0].readback_probes)


def test_exceljs_fixture_specs_are_stable() -> None:
    specs = exceljs_fixture_specs()

    assert [spec.fixture_id for spec in specs] == ["exceljs_table_validation_image_comment"]
    assert all(spec.tool == "exceljs" for spec in specs)
    assert "xl/tables/table1.xml" in specs[0].expected_parts
    assert "xl/media/image1.png" in specs[0].expected_parts
    assert any(probe["kind"] == "cell_style" for probe in specs[0].readback_probes)
    assert any(probe["kind"] == "cell_formula" for probe in specs[0].readback_probes)
    assert any(probe["kind"] == "sheet_protection" for probe in specs[0].readback_probes)
    assert any(probe["kind"] == "rich_text_runs" for probe in specs[0].readback_probes)


def test_npoi_fixture_specs_are_stable() -> None:
    specs = npoi_fixture_specs()

    assert [spec.fixture_id for spec in specs] == ["npoi_formula_comment_merge_protection"]
    assert all(spec.tool == "npoi" for spec in specs)
    assert "xl/comments1.xml" in specs[0].expected_parts
    assert "xl/drawings/vmlDrawing1.vml" in specs[0].expected_parts
    assert any(probe["kind"] == "comment_text" for probe in specs[0].readback_probes)
    assert any(probe["kind"] == "sheet_protection" for probe in specs[0].readback_probes)
    assert any(probe["kind"] == "rich_text_runs" for probe in specs[0].readback_probes)


def test_workbook_protection_probe_reads_workbook_xml(tmp_path: Path) -> None:
    openpyxl = pytest.importorskip("openpyxl")

    workbook_path = tmp_path / "protected-workbook.xlsx"
    workbook = openpyxl.Workbook()
    workbook.security.lockStructure = True
    workbook.save(workbook_path)
    workbook.close()

    roundtrip = openpyxl.load_workbook(workbook_path)
    try:
        failures = _run_readback_probes(
            roundtrip,
            workbook_path,
            (
                {
                    "kind": "workbook_protection",
                    "expected": {"lockStructure": True},
                },
            ),
        )
    finally:
        roundtrip.close()

    assert failures == ()


@pytest.mark.skipif(shutil.which("go") is None, reason="Go is required for Excelize fixture pack")
def test_generate_external_fixture_pack_without_validators(tmp_path: Path) -> None:
    repo_root = Path(__file__).resolve().parents[1]

    results = generate_external_fixture_pack(
        tmp_path,
        repo_root=repo_root,
        include_validators=False,
    )

    assert all(result.passed for result in results)
    assert (tmp_path / "manifest.json").exists()
    first = tmp_path / "excelize-sales-pivot-slicer-chart.xlsx"
    with ZipFile(first) as workbook:
        names = set(workbook.namelist())
    assert "xl/pivotTables/pivotTable1.xml" in names
    assert "xl/slicerCaches/slicerCache1.xml" in names
    assert "xl/charts/chart1.xml" in names
    if shutil.which("dotnet") is not None:
        closedxml_fixture = tmp_path / "closedxml-pivot-cf-table.xlsx"
        assert closedxml_fixture.exists()
        with ZipFile(closedxml_fixture) as workbook:
            closedxml_names = set(workbook.namelist())
        assert "xl/tables/table1.xml" in closedxml_names
        assert "xl/pivotTables/pivotTable.xml" in closedxml_names
        assert "pivotCache/pivotCacheDefinition1.xml" in closedxml_names
        rich_fixture = tmp_path / "closedxml-rich-comment-protection.xlsx"
        assert rich_fixture.exists()
        with ZipFile(rich_fixture) as workbook:
            rich_names = set(workbook.namelist())
            sheet_xml = workbook.read("xl/worksheets/sheet1.xml").decode()
            shared_strings_xml = workbook.read("xl/sharedStrings.xml").decode()
        assert "xl/comments1.xml" in rich_names
        assert "xl/drawings/vmldrawing.vml" in rich_names
        assert "sheetProtection" in sheet_xml
        assert ":r>" in shared_strings_xml
        npoi_fixture = tmp_path / "npoi-formula-comment-merge-protection.xlsx"
        assert npoi_fixture.exists()
        with ZipFile(npoi_fixture) as workbook:
            npoi_names = set(workbook.namelist())
            npoi_sheet_xml = workbook.read("xl/worksheets/sheet1.xml").decode()
            npoi_shared_strings_xml = workbook.read("xl/sharedStrings.xml").decode()
        assert "xl/comments1.xml" in npoi_names
        assert "xl/drawings/vmlDrawing1.vml" in npoi_names
        assert "sheetProtection" in npoi_sheet_xml
        assert "mergeCell" in npoi_sheet_xml
        assert "<r>" in npoi_shared_strings_xml
    exceljs_deps = (
        repo_root
        / "tools"
        / "external-oracles"
        / "exceljs"
        / "node_modules"
        / "exceljs"
        / "package.json"
    )
    if shutil.which("node") is not None and exceljs_deps.exists():
        exceljs_fixture = tmp_path / "exceljs-table-validation-image-comment.xlsx"
        assert exceljs_fixture.exists()
        with ZipFile(exceljs_fixture) as workbook:
            exceljs_names = set(workbook.namelist())
            exceljs_sheet_xml = workbook.read("xl/worksheets/sheet1.xml").decode()
            exceljs_shared_strings_xml = workbook.read("xl/sharedStrings.xml").decode()
        assert "xl/tables/table1.xml" in exceljs_names
        assert "xl/comments1.xml" in exceljs_names
        assert "xl/drawings/vmlDrawing1.vml" in exceljs_names
        assert "xl/drawings/drawing1.xml" in exceljs_names
        assert "xl/media/image1.png" in exceljs_names
        assert "dataValidations" in exceljs_sheet_xml
        assert "sheetProtection" in exceljs_sheet_xml
        assert "<r>" in exceljs_shared_strings_xml
    apache_poi_classes = (
        repo_root
        / "tools"
        / "external-oracles"
        / "apache-poi"
        / "build"
        / "classes"
        / "PoiOracle.class"
    )
    if (
        (Path("/opt/homebrew/opt/openjdk/bin/java").exists() or shutil.which("java") is not None)
        and apache_poi_classes.exists()
    ):
        apache_poi_fixture = tmp_path / "apache-poi-table-validation-image-comment.xlsx"
        assert apache_poi_fixture.exists()
        with ZipFile(apache_poi_fixture) as workbook:
            apache_poi_names = set(workbook.namelist())
            apache_poi_workbook_xml = workbook.read("xl/workbook.xml").decode()
            apache_poi_sheet_xml = workbook.read("xl/worksheets/sheet1.xml").decode()
            apache_poi_shared_strings_xml = workbook.read("xl/sharedStrings.xml").decode()
        assert "workbookProtection" in apache_poi_workbook_xml
        assert "lockStructure" in apache_poi_workbook_xml
        assert "xl/tables/table1.xml" in apache_poi_names
        assert "xl/comments1.xml" in apache_poi_names
        assert "xl/drawings/vmlDrawing0.vml" in apache_poi_names
        assert "xl/drawings/drawing1.xml" in apache_poi_names
        assert "xl/media/image1.png" in apache_poi_names
        assert "dataValidations" in apache_poi_sheet_xml
        assert "sheetProtection" in apache_poi_sheet_xml
        assert "<r>" in apache_poi_shared_strings_xml


@pytest.mark.skipif(shutil.which("go") is None, reason="Go is required for Excelize fixture pack")
@pytest.mark.skipif(
    os.environ.get("EXCELBENCH_RUN_WOLFXL_EXTERNAL") != "1",
    reason="Set EXCELBENCH_RUN_WOLFXL_EXTERNAL=1 to validate installed WolfXL",
)
def test_validate_wolfxl_external_fixture_pack(tmp_path: Path) -> None:
    pytest.importorskip("wolfxl")
    repo_root = Path(__file__).resolve().parents[1]
    generate_external_fixture_pack(
        tmp_path,
        repo_root=repo_root,
        include_validators=False,
    )

    results = validate_wolfxl_external_fixture_pack(tmp_path)

    assert all(result.passed for result in results)
    assert (tmp_path / "wolfxl-validation.json").exists()
    assert all(not result.missing_parts_after_save for result in results)
    assert all(not result.readback_failures for result in results)
    assert any(result.readback_probes for result in results)
