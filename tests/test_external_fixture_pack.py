"""Tests for local external-oracle fixture pack definitions."""

from __future__ import annotations

import os
import shutil
from pathlib import Path
from zipfile import ZipFile

import pytest

from excelbench.harness.external_fixture_pack import (
    excelize_fixture_specs,
    generate_external_fixture_pack,
)
from excelbench.harness.external_wolfxl_validation import (
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
