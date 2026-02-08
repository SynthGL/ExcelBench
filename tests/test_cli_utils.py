"""Tests for CLI utility functions: _results_from_json, _write_profile_index, show_summary."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from excelbench.cli import _results_from_json, _write_profile_index, show_summary
from excelbench.models import Importance, OperationType

JSONDict = dict[str, Any]


def _minimal_results_json(
    *,
    profile: str = "xlsx",
    extra_results: list[JSONDict] | None = None,
) -> JSONDict:
    """Build a minimal but valid results.json dict."""
    return {
        "metadata": {
            "benchmark_version": "0.1.0",
            "run_date": "2026-01-01T00:00:00+00:00",
            "excel_version": "16.0",
            "platform": "Darwin-arm64",
            "profile": profile,
        },
        "libraries": {
            "openpyxl": {
                "name": "openpyxl",
                "version": "3.1.0",
                "language": "python",
                "capabilities": ["read", "write"],
            },
        },
        "results": extra_results or [],
    }


# ═════════════════════════════════════════════════
# _results_from_json
# ═════════════════════════════════════════════════


class TestResultsFromJson:
    def test_metadata_parsing(self) -> None:
        data = _minimal_results_json()
        result = _results_from_json(data)
        assert result.metadata.benchmark_version == "0.1.0"
        assert result.metadata.platform == "Darwin-arm64"
        assert result.metadata.profile == "xlsx"

    def test_metadata_default_profile(self) -> None:
        data = _minimal_results_json()
        del data["metadata"]["profile"]
        result = _results_from_json(data)
        assert result.metadata.profile == "xlsx"

    def test_libraries_parsing(self) -> None:
        data = _minimal_results_json()
        result = _results_from_json(data)
        assert "openpyxl" in result.libraries
        lib = result.libraries["openpyxl"]
        assert lib.name == "openpyxl"
        assert lib.version == "3.1.0"
        assert "read" in lib.capabilities
        assert "write" in lib.capabilities

    def test_new_schema_test_cases(self) -> None:
        data = _minimal_results_json(
            extra_results=[
                {
                    "feature": "cell_values",
                    "library": "openpyxl",
                    "scores": {"read": 3, "write": 2},
                    "test_cases": {
                        "tc1": {
                            "read": {
                                "passed": True,
                                "expected": {"type": "string"},
                                "actual": {"type": "string"},
                                "notes": None,
                                "importance": "basic",
                                "label": "simple",
                            },
                            "write": {
                                "passed": False,
                                "expected": {"type": "string"},
                                "actual": {"type": "number"},
                                "notes": "mismatch",
                            },
                        }
                    },
                }
            ]
        )
        result = _results_from_json(data)
        assert len(result.scores) == 1
        score = result.scores[0]
        assert score.feature == "cell_values"
        assert score.read_score == 3
        assert score.write_score == 2
        assert len(score.test_results) == 2

        read_tr = [r for r in score.test_results if r.operation == OperationType.READ]
        assert len(read_tr) == 1
        assert read_tr[0].passed is True
        assert read_tr[0].importance == Importance.BASIC
        assert read_tr[0].label == "simple"

        write_tr = [
            r for r in score.test_results if r.operation == OperationType.WRITE
        ]
        assert len(write_tr) == 1
        assert write_tr[0].passed is False
        assert write_tr[0].notes == "mismatch"

    def test_legacy_schema_test_cases(self) -> None:
        data = _minimal_results_json(
            extra_results=[
                {
                    "feature": "formulas",
                    "library": "openpyxl",
                    "scores": {"read": 2},
                    "test_cases": {
                        "legacy_tc": {
                            "passed": True,
                            "expected": {"formula": "=1+1"},
                            "actual": {"formula": "=1+1"},
                            "notes": None,
                        }
                    },
                }
            ]
        )
        result = _results_from_json(data)
        assert len(result.scores) == 1
        assert len(result.scores[0].test_results) == 1
        tr = result.scores[0].test_results[0]
        assert tr.operation == OperationType.READ
        assert tr.passed is True

    def test_empty_results(self) -> None:
        data = _minimal_results_json()
        result = _results_from_json(data)
        assert result.scores == []

    def test_pivot_table_note_injection(self) -> None:
        data = _minimal_results_json(
            extra_results=[
                {
                    "feature": "pivot_tables",
                    "library": "openpyxl",
                    "scores": {},
                    "test_cases": {},
                }
            ]
        )
        result = _results_from_json(data)
        score = result.scores[0]
        assert score.notes is not None
        assert "macOS" in score.notes

    def test_pivot_table_existing_notes_preserved(self) -> None:
        data = _minimal_results_json(
            extra_results=[
                {
                    "feature": "pivot_tables",
                    "library": "openpyxl",
                    "scores": {},
                    "test_cases": {},
                    "notes": "Custom note",
                }
            ]
        )
        result = _results_from_json(data)
        assert result.scores[0].notes == "Custom note"

    def test_multiple_libraries(self) -> None:
        data = _minimal_results_json()
        data["libraries"]["xlsxwriter"] = {
            "name": "xlsxwriter",
            "version": "3.2.0",
            "language": "python",
            "capabilities": ["write"],
        }
        result = _results_from_json(data)
        assert len(result.libraries) == 2
        assert "xlsxwriter" in result.libraries

    def test_real_results_file(self) -> None:
        """Smoke test: parse the actual results.json from the repo."""
        results_path = Path(__file__).parent.parent / "results" / "results.json"
        if not results_path.exists():
            pytest.skip("No results.json available")
        import json

        with open(results_path) as f:
            data = json.load(f)
        result = _results_from_json(data)
        assert len(result.scores) > 0
        assert len(result.libraries) > 0


# ═════════════════════════════════════════════════
# _write_profile_index
# ═════════════════════════════════════════════════


class TestWriteProfileIndex:
    def test_creates_readme(self, tmp_path: Path) -> None:
        _write_profile_index(tmp_path)
        readme = tmp_path / "README.md"
        assert readme.exists()
        content = readme.read_text()
        assert "xlsx profile" in content
        assert "xls profile" in content

    def test_creates_nested_dir(self, tmp_path: Path) -> None:
        nested = tmp_path / "deep" / "nested"
        _write_profile_index(nested)
        assert (nested / "README.md").exists()


# ═════════════════════════════════════════════════
# show_summary
# ═════════════════════════════════════════════════


class TestShowSummary:
    def test_does_not_crash(self) -> None:
        data = _minimal_results_json(
            extra_results=[
                {
                    "feature": "cell_values",
                    "library": "openpyxl",
                    "scores": {"read": 3, "write": 2},
                    "test_cases": {},
                }
            ]
        )
        result = _results_from_json(data)
        # Just verify it doesn't raise
        show_summary(result)

    def test_empty_results_summary(self) -> None:
        data = _minimal_results_json()
        result = _results_from_json(data)
        show_summary(result)

    def test_read_only_library(self) -> None:
        data = _minimal_results_json()
        data["libraries"] = {
            "calamine": {
                "name": "calamine",
                "version": "0.6.0",
                "language": "python",
                "capabilities": ["read"],
            },
        }
        data["results"] = [
            {
                "feature": "cell_values",
                "library": "calamine",
                "scores": {"read": 3},
                "test_cases": {},
            }
        ]
        result = _results_from_json(data)
        show_summary(result)

    def test_write_only_library(self) -> None:
        data = _minimal_results_json()
        data["libraries"] = {
            "xlsxwriter": {
                "name": "xlsxwriter",
                "version": "3.2.0",
                "language": "python",
                "capabilities": ["write"],
            },
        }
        data["results"] = [
            {
                "feature": "cell_values",
                "library": "xlsxwriter",
                "scores": {"write": 2},
                "test_cases": {},
            }
        ]
        result = _results_from_json(data)
        show_summary(result)
