"""Tests for utility functions in excelbench.results.renderer."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from unittest.mock import patch

from excelbench.models import (
    BenchmarkMetadata,
    BenchmarkResults,
    FeatureScore,
    Importance,
    LibraryInfo,
    OperationType,
    TestResult,
)
from excelbench.results.renderer import (
    _get_git_commit,
    _group_test_cases,
    _load_previous_history_entry,
    _render_fidelity_deltas,
    _render_per_test_table,
    _test_outcome_kind,
    render_markdown,
    score_emoji,
)

# ─────────────────────────────────────────────────
# score_emoji
# ─────────────────────────────────────────────────


def test_score_emoji_none() -> None:
    assert score_emoji(None) == "➖"


def test_score_emoji_3() -> None:
    assert score_emoji(3) == "🟢 3"


def test_score_emoji_2() -> None:
    assert score_emoji(2) == "🟡 2"


def test_score_emoji_1() -> None:
    assert score_emoji(1) == "🟠 1"


def test_score_emoji_0() -> None:
    assert score_emoji(0) == "🔴 0"


def test_score_emoji_negative() -> None:
    assert score_emoji(-1) == "🔴 0"


# ─────────────────────────────────────────────────
# _group_test_cases
# ─────────────────────────────────────────────────


def _make_tr(
    tc_id: str,
    op: OperationType,
    passed: bool = True,
    label: str | None = None,
) -> TestResult:
    return TestResult(
        test_case_id=tc_id,
        operation=op,
        passed=passed,
        expected={"val": 1},
        actual={"val": 1},
        importance=Importance.BASIC,
        label=label,
    )


def test_group_test_cases_groups_by_id() -> None:
    results = [
        _make_tr("bold", OperationType.READ),
        _make_tr("bold", OperationType.WRITE, passed=False),
        _make_tr("italic", OperationType.READ),
    ]
    grouped = _group_test_cases(results)
    assert "bold" in grouped
    assert "read" in grouped["bold"]
    assert "write" in grouped["bold"]
    assert "italic" in grouped
    assert "read" in grouped["italic"]
    assert "write" not in grouped["italic"]


def test_group_test_cases_preserves_passed() -> None:
    results = [_make_tr("tc1", OperationType.READ, passed=False)]
    grouped = _group_test_cases(results)
    assert grouped["tc1"]["read"]["passed"] is False


def test_group_test_cases_includes_outcome_kind() -> None:
    tr = TestResult(
        test_case_id="tc1",
        operation=OperationType.READ,
        passed=False,
        expected={"val": 1},
        actual={"error": "feature not implemented by adapter"},
        notes="Exception: NotImplementedError",
        importance=Importance.BASIC,
    )
    grouped = _group_test_cases([tr])
    assert grouped["tc1"]["read"]["outcome"] == "not_implemented"


def test_group_test_cases_includes_label() -> None:
    results = [_make_tr("tc1", OperationType.READ, label="Bold text")]
    grouped = _group_test_cases(results)
    assert grouped["tc1"]["read"]["label"] == "Bold text"


def test_group_test_cases_empty() -> None:
    assert _group_test_cases([]) == {}


# ─────────────────────────────────────────────────
# _get_git_commit
# ─────────────────────────────────────────────────


def test_get_git_commit_success() -> None:
    with patch("subprocess.run") as mock_run:
        mock_run.return_value.returncode = 0
        mock_run.return_value.stdout = "abc1234\n"
        assert _get_git_commit() == "abc1234"


def test_get_git_commit_failure() -> None:
    with patch("subprocess.run") as mock_run:
        mock_run.return_value.returncode = 128
        mock_run.return_value.stdout = ""
        assert _get_git_commit() is None


def test_get_git_commit_no_git() -> None:
    with patch("subprocess.run", side_effect=FileNotFoundError):
        assert _get_git_commit() is None


# ─────────────────────────────────────────────────
# _render_per_test_table
# ─────────────────────────────────────────────────


def test_render_per_test_write_only_label() -> None:
    """A test with only write results should use write_tr for label/importance."""
    score = FeatureScore(
        feature="borders",
        library="openpyxl",
        test_results=[
            _make_tr("tc1", OperationType.WRITE, label="Thick border"),
        ],
    )
    lines = _render_per_test_table(score)
    text = "\n".join(lines)
    assert "Thick border" in text
    assert "basic" in text  # importance


def test_render_per_test_missing_read_dash() -> None:
    """When has_read is True but a test_case has no read result, show —."""
    score = FeatureScore(
        feature="cell_values",
        library="openpyxl",
        test_results=[
            _make_tr("tc1", OperationType.READ),
            _make_tr("tc1", OperationType.WRITE),
            _make_tr("tc2", OperationType.WRITE),  # no read for tc2
        ],
    )
    lines = _render_per_test_table(score)
    text = "\n".join(lines)
    assert "—" in text  # tc2 read column is a dash


def test_render_per_test_missing_write_dash() -> None:
    """When has_write is True but a test_case has no write result, show —."""
    score = FeatureScore(
        feature="cell_values",
        library="openpyxl",
        test_results=[
            _make_tr("tc1", OperationType.READ),
            _make_tr("tc1", OperationType.WRITE),
            _make_tr("tc2", OperationType.READ),  # no write for tc2
        ],
    )
    lines = _render_per_test_table(score)
    text = "\n".join(lines)
    assert "—" in text  # tc2 write column is a dash


def test_render_per_test_marks_not_implemented_vs_incorrect() -> None:
    score = FeatureScore(
        feature="pivot_tables",
        library="pyumya",
        test_results=[
            TestResult(
                test_case_id="tc_ni",
                operation=OperationType.READ,
                passed=False,
                expected={"x": 1},
                actual={"error": "pivot tables not implemented"},
                notes="Exception: NotImplementedError",
            ),
            TestResult(
                test_case_id="tc_bad",
                operation=OperationType.WRITE,
                passed=False,
                expected={"x": 1},
                actual={"x": 0},
            ),
        ],
    )
    text = "\n".join(_render_per_test_table(score))
    assert "🚫 NI" in text
    assert "❌" in text
    assert "Legend: ✅ pass · 🚫 NI (not implemented) · ❌ incorrect" in text


def test_test_outcome_kind_prefers_not_implemented() -> None:
    tr = TestResult(
        test_case_id="tc",
        operation=OperationType.READ,
        passed=False,
        expected={"x": 1},
        actual={"error": "not implemented yet"},
    )
    assert _test_outcome_kind(tr) == "not_implemented"


def test_test_outcome_kind_incorrect_default() -> None:
    tr = TestResult(
        test_case_id="tc",
        operation=OperationType.READ,
        passed=False,
        expected={"x": 1},
        actual={"x": 2},
    )
    assert _test_outcome_kind(tr) == "incorrect"


# ─────────────────────────────────────────────────
# render_markdown edge cases
# ─────────────────────────────────────────────────


def _make_results(
    *,
    features: list[str] | None = None,
    libs: dict[str, list[str]] | None = None,
    scores: list[FeatureScore] | None = None,
) -> BenchmarkResults:
    """Build minimal BenchmarkResults for renderer tests."""
    if features is None:
        features = ["cell_values"]
    if libs is None:
        libs = {"openpyxl": ["read", "write"]}
    if scores is None:
        scores = []

    libraries = {
        name: LibraryInfo(
            name=name, version="1.0", language="python", capabilities=set(caps)
        )
        for name, caps in libs.items()
    }
    return BenchmarkResults(
        metadata=BenchmarkMetadata(
            benchmark_version="0.1",
            run_date=datetime(2026, 1, 1),
            excel_version="16.0",
            platform="test",
            profile="xlsx",
        ),
        libraries=libraries,
        scores=scores,
    )


def test_render_markdown_missing_score_skips_lib(tmp_path: Path) -> None:
    """When a feature has no score for a library, skip that lib in detail."""
    results = _make_results(
        features=["cell_values", "borders"],
        libs={"openpyxl": ["read", "write"], "xlrd": ["read"]},
        scores=[
            FeatureScore(
                feature="cell_values",
                library="openpyxl",
                read_score=3,
                write_score=2,
            ),
            # xlrd has no score for cell_values — hits continue at line 199
            # openpyxl has no score for borders — hits continue at line 199
        ],
    )
    out = tmp_path / "test.md"
    render_markdown(results, out)
    content = out.read_text()
    assert "cell_values" in content
    assert "openpyxl" in content


def test_render_markdown_write_only_lib_stats(tmp_path: Path) -> None:
    """A write-only lib should show write stats but skip read (None score)."""
    results = _make_results(
        features=["cell_values"],
        libs={"xlsxwriter": ["write"]},
        scores=[
            FeatureScore(
                feature="cell_values",
                library="xlsxwriter",
                write_score=2,
            ),
        ],
    )
    out = tmp_path / "test.md"
    render_markdown(results, out)
    content = out.read_text()
    assert "xlsxwriter" in content
    assert "Write" in content


def test_render_fidelity_deltas_none_previous() -> None:
    results = _make_results(
        scores=[
            FeatureScore(
                feature="cell_values",
                library="openpyxl",
                read_score=3,
                write_score=3,
            )
        ]
    )
    lines = _render_fidelity_deltas(results, None)
    assert "No previous run found for this profile." in "\n".join(lines)


def test_render_fidelity_deltas_regression_and_improvement() -> None:
    results = _make_results(
        scores=[
            FeatureScore(
                feature="cell_values",
                library="openpyxl",
                read_score=2,
                write_score=3,
            )
        ]
    )
    previous = {
        "profile": "xlsx",
        "scores": {
            "openpyxl": {
                "cell_values": {
                    "read": 3,
                    "write": 2,
                }
            }
        },
    }
    text = "\n".join(_render_fidelity_deltas(results, previous))
    assert "### Regressions" in text
    assert "openpyxl · cell_values · read: 3 → 2" in text
    assert "### Improvements" in text
    assert "openpyxl · cell_values · write: 2 → 3" in text


def test_load_previous_history_entry_filters_profile(tmp_path: Path) -> None:
    history = tmp_path / "history.jsonl"
    history.write_text(
        "\n".join(
            [
                '{"profile":"xls","scores":{"a":{}}}',
                '{"profile":"xlsx","scores":{"openpyxl":{"cell_values":{"read":3,"write":3}}}}',
                '{"profile":"xlsx","scores":{"openpyxl":{"cell_values":{"read":2,"write":3}}}}',
            ]
        )
    )
    latest = _load_previous_history_entry(history, "xlsx")
    assert latest is not None
    assert latest["scores"]["openpyxl"]["cell_values"]["read"] == 2
