from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest
from typer.testing import CliRunner

from excelbench.cli import _write_cross_language_index, app


def test_write_cross_language_index(tmp_path: Path) -> None:
    _write_cross_language_index(tmp_path, ["apache-poi", "excelize"])
    index = tmp_path / "CONTEXT.md"
    assert index.exists()
    content = index.read_text()
    assert "cross-language context" in content.lower()
    assert "apache-poi" in content
    assert "excelize" in content


def test_cross_language_context_exits_when_no_adapters(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    monkeypatch.setattr(
        "excelbench.harness.adapters.get_all_adapters",
        lambda: [],
    )
    runner = CliRunner()
    result = runner.invoke(app, ["cross-language-context", "--output", str(tmp_path)])
    assert result.exit_code == 1
    assert "No cross-language adapters available" in result.output


def test_cross_language_context_writes_results_and_context(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    class FakeAdapter:
        def __init__(self, name: str) -> None:
            self.name = name

    def fake_get_all_adapters() -> list[FakeAdapter]:
        return [FakeAdapter("apache-poi"), FakeAdapter("excelize")]

    def fake_run_benchmark(
        test_dir: Path,
        adapters: list[FakeAdapter],
        profile: str = "xlsx",
    ) -> str:
        return "fake-results"

    def fake_render_results(results: Any, output_dir: Path) -> None:
        output_dir.mkdir(parents=True, exist_ok=True)
        (output_dir / "README.md").write_text("rendered")
        (output_dir / "results.json").write_text("{}")
        (output_dir / "matrix.csv").write_text("feature,library")

    monkeypatch.setattr("excelbench.harness.adapters.get_all_adapters", fake_get_all_adapters)
    monkeypatch.setattr("excelbench.harness.runner.run_benchmark", fake_run_benchmark)
    monkeypatch.setattr("excelbench.results.render_results", fake_render_results)

    runner = CliRunner()
    result = runner.invoke(app, ["cross-language-context", "--output", str(tmp_path)])
    assert result.exit_code == 0
    assert (tmp_path / "README.md").exists()
    assert (tmp_path / "results.json").exists()
    assert (tmp_path / "matrix.csv").exists()
    assert (tmp_path / "CONTEXT.md").exists()
