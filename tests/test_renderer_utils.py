"""Tests for utility functions in excelbench.results.renderer."""

from __future__ import annotations

from unittest.mock import patch

from excelbench.models import Importance, OperationType, TestResult
from excelbench.results.renderer import (
    _get_git_commit,
    _group_test_cases,
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
