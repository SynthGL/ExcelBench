"""Tests for Tier-4 handlers: sheet_protection, page_setup, chart_anchor.

Builds the Tier-4 structural fixtures into a temp directory via
scripts/build_tier4_fixtures.py, assembles a benchmark directory from the
generated manifest fragment, and runs the full read + write lanes against
OpenpyxlAdapter — proving the runner dispatch, fixtures, and openpyxl
verifier chain end to end. Also checks that adapters lacking the Tier-4
methods yield structured failure results instead of crashes.
"""

from __future__ import annotations

import importlib.util
import shutil
import sys
from pathlib import Path
from types import ModuleType
from typing import Any

from excelbench.generator.generate import load_manifest
from excelbench.harness.adapters.base import ReadOnlyAdapter
from excelbench.harness.adapters.openpyxl_adapter import OpenpyxlAdapter
from excelbench.harness.runner import run_benchmark
from excelbench.models import (
    BorderInfo,
    CellFormat,
    CellValue,
    DiagnosticCategory,
    LibraryInfo,
)

REPO_ROOT = Path(__file__).resolve().parent.parent
BUILDER_PATH = REPO_ROOT / "scripts" / "build_tier4_fixtures.py"

TIER4_FEATURES = ["sheet_protection", "page_setup", "chart_anchor"]
CASE_COUNTS = {"sheet_protection": 3, "page_setup": 3, "chart_anchor": 2}


def _load_builder() -> ModuleType:
    spec = importlib.util.spec_from_file_location("build_tier4_fixtures", BUILDER_PATH)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules["build_tier4_fixtures"] = module
    spec.loader.exec_module(module)
    return module


def _build_bench_dir(tmp_path: Path) -> Path:
    """Build Tier-4 fixtures into tmp and assemble a runnable benchmark dir."""
    builder = _load_builder()
    fixtures_root = tmp_path / "fixtures"
    fragment_path = builder.build_all(fixtures_root)
    assert fragment_path.exists()

    bench_dir = tmp_path / "bench"
    bench_dir.mkdir()
    shutil.copytree(fixtures_root / "tier4", bench_dir / "tier4")
    shutil.copy(fragment_path, bench_dir / "manifest.json")
    return bench_dir


class TestTier4Benchmark:
    def test_openpyxl_read_and_write_scores(self, tmp_path: Path) -> None:
        bench_dir = _build_bench_dir(tmp_path)

        results = run_benchmark(
            bench_dir, adapters=[OpenpyxlAdapter()], features=TIER4_FEATURES
        )

        by_feature = {score.feature: score for score in results.scores}
        assert set(by_feature) == set(TIER4_FEATURES)
        for feature in TIER4_FEATURES:
            score = by_feature[feature]
            failures = [r for r in score.test_results if not r.passed]
            assert score.read_score == 3, (
                f"{feature} read failures: {[(r.test_case_id, r.actual) for r in failures]}"
            )
            assert score.write_score == 3, (
                f"{feature} write failures: {[(r.test_case_id, r.actual) for r in failures]}"
            )

    def test_fragment_loads_as_manifest(self, tmp_path: Path) -> None:
        builder = _load_builder()
        fragment_path = builder.build_all(tmp_path)

        manifest = load_manifest(fragment_path)

        assert [f.feature for f in manifest.files] == TIER4_FEATURES
        for test_file in manifest.files:
            assert test_file.tier == 4
            assert test_file.file_format == "xlsx"
        for test_file in manifest.files:
            assert len(test_file.test_cases) == CASE_COUNTS[test_file.feature]
            if test_file.feature == "chart_anchor":
                assert all(tc.sheet is None for tc in test_file.test_cases)
            else:
                assert all(
                    tc.sheet and tc.sheet.startswith(test_file.feature)
                    for tc in test_file.test_cases
                )
        protection = manifest.files[0]
        assert {tc.id for tc in protection.test_cases} == {
            "prot_basic",
            "prot_password",
            "prot_granular",
        }
        # The write-only password must ride along in the manifest payload.
        password_case = next(
            tc for tc in protection.test_cases if tc.id == "prot_password"
        )
        assert password_case.expected["protection"]["password"] == "bench-secret"


class _Tier3OnlyReader(ReadOnlyAdapter):
    """Read-only adapter that implements every pre-Tier-4 read and nothing else."""

    def __init__(self) -> None:
        self._inner = OpenpyxlAdapter()

    @property
    def info(self) -> LibraryInfo:
        return LibraryInfo(
            name="tier3-only-reader",
            version="0",
            language="python",
            capabilities={"read"},
        )

    def open_workbook(self, path: Path) -> Any:
        return self._inner.open_workbook(path)

    def close_workbook(self, workbook: Any) -> None:
        self._inner.close_workbook(workbook)

    def get_sheet_names(self, workbook: Any) -> list[str]:
        return self._inner.get_sheet_names(workbook)

    def read_cell_value(self, workbook: Any, sheet: str, cell: str) -> CellValue:
        return self._inner.read_cell_value(workbook, sheet, cell)

    def read_cell_format(self, workbook: Any, sheet: str, cell: str) -> CellFormat:
        return self._inner.read_cell_format(workbook, sheet, cell)

    def read_cell_border(self, workbook: Any, sheet: str, cell: str) -> BorderInfo:
        return self._inner.read_cell_border(workbook, sheet, cell)

    def read_row_height(self, workbook: Any, sheet: str, row: int) -> float | None:
        return self._inner.read_row_height(workbook, sheet, row)

    def read_column_width(self, workbook: Any, sheet: str, column: str) -> float | None:
        return self._inner.read_column_width(workbook, sheet, column)

    def read_merged_ranges(self, workbook: Any, sheet: str) -> list[str]:
        return self._inner.read_merged_ranges(workbook, sheet)

    def read_conditional_formats(
        self, workbook: Any, sheet: str
    ) -> list[dict[str, Any]]:
        return self._inner.read_conditional_formats(workbook, sheet)

    def read_data_validations(self, workbook: Any, sheet: str) -> list[dict[str, Any]]:
        return self._inner.read_data_validations(workbook, sheet)

    def read_hyperlinks(self, workbook: Any, sheet: str) -> list[dict[str, Any]]:
        return self._inner.read_hyperlinks(workbook, sheet)

    def read_images(self, workbook: Any, sheet: str) -> list[dict[str, Any]]:
        return self._inner.read_images(workbook, sheet)

    def read_pivot_tables(self, workbook: Any, sheet: str) -> list[dict[str, Any]]:
        return self._inner.read_pivot_tables(workbook, sheet)

    def read_comments(self, workbook: Any, sheet: str) -> list[dict[str, Any]]:
        return self._inner.read_comments(workbook, sheet)

    def read_freeze_panes(self, workbook: Any, sheet: str) -> dict[str, Any]:
        return self._inner.read_freeze_panes(workbook, sheet)


class TestUnsupportedAdapter:
    def test_missing_tier4_methods_yield_structured_failures(
        self, tmp_path: Path
    ) -> None:
        bench_dir = _build_bench_dir(tmp_path)
        adapter = _Tier3OnlyReader()

        results = run_benchmark(bench_dir, adapters=[adapter], features=TIER4_FEATURES)

        assert len(results.scores) == 3
        for score in results.scores:
            read_results = [
                r for r in score.test_results if r.operation.value == "read"
            ]
            assert len(read_results) == CASE_COUNTS[score.feature]
            for result in read_results:
                assert result.passed is False
                assert "does not implement" in str(result.actual.get("error", ""))
                assert result.diagnostics, (
                    f"{score.feature}/{result.test_case_id} lacks diagnostics"
                )
                assert (
                    result.diagnostics[0].category
                    == DiagnosticCategory.UNSUPPORTED_FEATURE
                )
            # Read-only adapter: the write lane is skipped, not scored.
            assert score.write_score is None
