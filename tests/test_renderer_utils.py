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
    _compute_fidelity_deltas,
    _get_git_commit,
    _group_test_cases,
    _render_fidelity_deltas,
    _render_per_test_table,
    render_markdown,
    render_results,
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


def test_render_markdown_hides_pyumya_and_includes_modify_label(tmp_path: Path) -> None:
    results = _make_results(
        features=["cell_values"],
        libs={"openpyxl": ["read", "write"], "pyumya": ["read", "write"]},
        scores=[
            FeatureScore(
                feature="cell_values",
                library="openpyxl",
                read_score=3,
                write_score=3,
            ),
            FeatureScore(
                feature="cell_values",
                library="pyumya",
                read_score=3,
                write_score=3,
            ),
        ],
    )
    out = tmp_path / "test.md"
    render_markdown(results, out)
    content = out.read_text()
    assert "modify: Rewrite" in content
    assert "pyumya" not in content


def test_render_results_removes_stale_why_failed_when_all_pass(tmp_path: Path) -> None:
    failed = _make_results(
        scores=[
            FeatureScore(
                feature="cell_values",
                library="openpyxl",
                read_score=0,
                test_results=[_make_tr("tc1", OperationType.READ, passed=False)],
            )
        ]
    )
    passed = _make_results(
        scores=[
            FeatureScore(
                feature="cell_values",
                library="openpyxl",
                read_score=3,
                test_results=[_make_tr("tc1", OperationType.READ, passed=True)],
            )
        ]
    )

    render_results(failed, tmp_path)
    assert (tmp_path / "WHY_FAILED.md").exists()

    render_results(passed, tmp_path)

    assert not (tmp_path / "WHY_FAILED.md").exists()


# ─────────────────────────────────────────────────
# fidelity deltas
# ─────────────────────────────────────────────────


def test_compute_fidelity_deltas_detects_changes() -> None:
    previous = {
        "scores": {"openpyxl": {"cell_values": {"read": 3, "write": 3}}},
    }
    current = {
        "scores": {"openpyxl": {"cell_values": {"read": 2, "write": 3}}},
    }
    deltas = _compute_fidelity_deltas(previous, current)
    assert deltas == [
        {
            "library": "openpyxl",
            "feature": "cell_values",
            "mode": "read",
            "previous": 3,
            "current": 2,
            "delta": -1,
        }
    ]


def test_render_fidelity_deltas_needs_two_runs(tmp_path: Path) -> None:
    out_dir = tmp_path / "results"
    out_dir.mkdir(parents=True)
    (out_dir / "history.jsonl").write_text('{"scores": {}}\n')
    _render_fidelity_deltas(out_dir)
    content = (out_dir / "FIDELITY_DELTAS.md").read_text()
    assert "Need at least two runs" in content


def test_render_fidelity_deltas_writes_regression_table(tmp_path: Path) -> None:
    out_dir = tmp_path / "results"
    out_dir.mkdir(parents=True)
    (out_dir / "history.jsonl").write_text(
        "\n".join(
            [
                '{"run_date":"2026-01-01T00:00:00Z","scores":{"openpyxl":{"cell_values":{"read":3,"write":3}}}}',
                '{"run_date":"2026-01-02T00:00:00Z","scores":{"openpyxl":{"cell_values":{"read":2,"write":3}}}}',
            ]
        )
        + "\n"
    )
    _render_fidelity_deltas(out_dir)
    content = (out_dir / "FIDELITY_DELTAS.md").read_text()
    assert "Regressions: **1**" in content
    assert "| openpyxl | cell_values | read | 3 | 2 | -1 |" in content
