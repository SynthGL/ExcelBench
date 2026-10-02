"""Tests for the Formula Recalculation benchmark tier."""

from __future__ import annotations

import json
import sys
from collections.abc import Callable
from importlib.util import module_from_spec, spec_from_file_location
from pathlib import Path
from types import ModuleType
from typing import Any, cast

import pytest

from excelbench.harness import calc
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
    assert all(cell["type"] in {"number", "string", "bool"} for cell in expected["cells"].values())


def test_value_comparison_respects_tolerance_and_scalar_types() -> None:
    assert values_match(10.0, 10.0000005)
    assert values_match(1_000_000_000.0, 1_000_000_000.5)
    assert not values_match(10.0, 10.000002)
    assert values_match("exact", "exact")
    assert not values_match("exact", "different")
    assert values_match(True, True)
    assert not values_match(True, 1)


@pytest.mark.parametrize(
    ("direct_url", "expected"),
    [
        (None, "2.1.0"),
        (
            '{"url": "file:///build/wolfxl-2.1.0.whl", "archive_info": {}}',
            "2.1.0 (local build)",
        ),
        (
            '{"url": "file:///src/wolfxl", "dir_info": {"editable": true}}',
            "2.1.0 (local editable build)",
        ),
        (
            (
                '{"url": "https://example.invalid/wolfxl.git", '
                '"vcs_info": {"vcs": "git", "commit_id": "1a9ad6700abcdef0123"}}'
            ),
            "2.1.0 (vcs 1a9ad6700abc)",
        ),
    ],
)
def test_package_version_names_non_registry_installs(direct_url: str | None, expected: str) -> None:
    """A local build must not be reported as the registry release it shares a version with."""
    version = calc._describe_install("2.1.0", direct_url)

    assert version == expected
    assert "/" not in version  # install paths and URLs never reach published results


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
            wolfxl_result = wolfxl_engine.calculate(FIXTURE, tmp_path / "wolfxl.xlsx")
            assert wolfxl_result["status"] == "passed", wolfxl_result["reason"]

    libreoffice_result = LibreOfficeCalcEngine().calculate(FIXTURE, tmp_path / "libreoffice.xlsx")
    assert libreoffice_result["status"] == "passed", libreoffice_result["reason"]


def test_aspose_engine_reports_unsupported_formulas_without_raising(
    tmp_path: Path,
) -> None:
    result = AsposeCellsFossCalcEngine().calculate(FIXTURE, tmp_path / "aspose.xlsx")

    assert result["status"] in {"failed", "unavailable"}
    assert result["reason"]


class _ApiOnlyEngine:
    """An engine whose API returns correct values but whose save writes none."""

    name = "api-only"

    def available(self) -> bool:
        return True

    def version(self) -> str | None:
        return "1.0"

    def calculate(self, input_path: Path, output_path: Path) -> dict[str, Any]:
        expected = json.loads(EXPECTED.read_text(encoding="utf-8"))["cells"]
        return {
            "status": "passed",
            "values": {reference: cell["value"] for reference, cell in expected.items()},
            "saved_values": {},
            "reason": None,
        }


def test_api_results_score_separately_from_saved_values(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr(calc, "_engines", lambda: (_ApiOnlyEngine(),))

    result = calc.run_calc_suite(FIXTURE, EXPECTED, tmp_path)["engines"]["api-only"]

    assert result["status"] == "passed"
    assert result["matched"] == result["total"]
    assert result["saved_matched"] == 0
    assert "| api-only | 1.0 | passed | 133/133 | 0/133 |" in (tmp_path / "README.md").read_text(
        encoding="utf-8"
    )


def test_commercial_engine_matches_in_process_engine(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """The subprocess protocol reports what the same build reports in process."""
    pytest.importorskip("wolfxl")
    engine = calc.WolfXLCommercialCalcEngine()
    monkeypatch.delenv(calc.WOLFXL_COMMERCIAL_PYTHON_ENV, raising=False)
    assert not engine.available()

    monkeypatch.setenv(calc.WOLFXL_COMMERCIAL_PYTHON_ENV, sys.executable)
    in_process = WolfXLCalcEngine().calculate(FIXTURE, tmp_path / "in-process.xlsx")
    via_subprocess = engine.calculate(FIXTURE, tmp_path / "subprocess.xlsx")

    assert engine.version() == calc._package_version("wolfxl")
    assert via_subprocess["status"] == in_process["status"] == "passed"
    assert via_subprocess["values"] == in_process["values"]
    assert via_subprocess["saved_values"] == in_process["saved_values"]


def test_report_keeps_every_engine_row_in_the_table(tmp_path: Path) -> None:
    """Engine notes print after the table, so a note never splits it."""
    from excelbench.results.calc_renderer import render_calc_results

    engine = {"status": "failed", "matched": 0, "total": 1, "mismatched_cells": []}
    render_calc_results(
        {"oracle": "o", "engines": {"a": {**engine, "reason": "a broke"}, "b": engine}},
        tmp_path,
    )
    lines = (tmp_path / "README.md").read_text().splitlines()
    rows = [i for i, line in enumerate(lines) if line.startswith(("| a |", "| b |"))]
    assert rows == [rows[0], rows[0] + 1]
    assert lines.index("a: a broke") > rows[-1]
