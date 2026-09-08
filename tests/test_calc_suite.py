"""Tests for the Formula Recalculation benchmark tier."""

from __future__ import annotations

import json
from collections.abc import Callable
from importlib.util import module_from_spec, spec_from_file_location
from pathlib import Path
from types import ModuleType
from typing import Any, cast

import pytest

from excelbench.harness.calc import (
    SOFFICE_PATH,
    AsposeCellsFossCalcEngine,
    LibreOfficeCalcEngine,
    WolfXLCalcEngine,
    formula_cells,
    values_match,
)

requires_libreoffice = pytest.mark.skipif(
    not SOFFICE_PATH.is_file(), reason="LibreOffice binary not available"
)

FIXTURE = Path("fixtures/calc/financial_model.xlsx")
EXPECTED = Path("fixtures/calc/expected_values.json")


def _load_fixture_builder() -> Callable[[Path, Path], dict[str, Any]]:
    script_path = Path("scripts/build_calc_fixture.py")
    spec = spec_from_file_location("build_calc_fixture", script_path)
    assert spec is not None and spec.loader is not None
    module = module_from_spec(spec)
    assert isinstance(module, ModuleType)
    spec.loader.exec_module(module)
    return cast(Callable[[Path, Path], dict[str, Any]], module.build_fixture)


build_fixture = _load_fixture_builder()


@requires_libreoffice
def test_builder_regenerates_identical_expected_json(tmp_path: Path) -> None:
    fixture = tmp_path / "financial_model.xlsx"
    expected = tmp_path / "expected_values.json"

    first = build_fixture(fixture, expected)
    first_text = expected.read_text()
    second = build_fixture(fixture, expected)

    assert second == first
    assert expected.read_text() == first_text


def test_expected_values_cover_every_formula_cell() -> None:
    expected = json.loads(EXPECTED.read_text())

    assert set(expected["cells"]) == set(formula_cells(FIXTURE))
    assert 80 <= len(expected["cells"]) <= 140
    assert all(
        cell["type"] in {"number", "string", "bool"}
        for cell in expected["cells"].values()
    )


def test_value_comparison_respects_tolerance_and_scalar_types() -> None:
    assert values_match(10.0, 10.0000005)
    assert values_match(1_000_000_000.0, 1_000_000_000.5)
    assert not values_match(10.0, 10.000002)
    assert values_match("exact", "exact")
    assert not values_match("exact", "different")
    assert values_match(True, True)
    assert not values_match(True, 1)


@requires_libreoffice
def test_local_wolfxl_and_libreoffice_engines_calculate(tmp_path: Path) -> None:
    wolfxl_engine = WolfXLCalcEngine()
    if wolfxl_engine.available():
        import wolfxl

        try:
            major, minor = (int(part) for part in wolfxl.__version__.split(".")[:2])
        except (AttributeError, ValueError):
            major, minor = 0, 0
        if (major, minor) >= (2, 1):
            wolfxl_result = wolfxl_engine.calculate(
                FIXTURE, tmp_path / "wolfxl.xlsx"
            )
            assert wolfxl_result["status"] == "passed", wolfxl_result["reason"]

    libreoffice_result = LibreOfficeCalcEngine().calculate(
        FIXTURE, tmp_path / "libreoffice.xlsx"
    )
    assert libreoffice_result["status"] == "passed", libreoffice_result["reason"]


def test_aspose_engine_reports_unsupported_formulas_without_raising(
    tmp_path: Path,
) -> None:
    result = AsposeCellsFossCalcEngine().calculate(FIXTURE, tmp_path / "aspose.xlsx")

    assert result["status"] in {"failed", "unavailable"}
    assert result["reason"]
