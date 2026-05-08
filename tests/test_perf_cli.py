import json
from datetime import UTC, datetime
from pathlib import Path

from openpyxl import Workbook

from excelbench.cli import perf
from excelbench.generator.generate import write_manifest
from excelbench.models import Importance, Manifest
from excelbench.models import TestCase as BenchCase
from excelbench.models import TestFile as BenchFile
from excelbench.perf.renderer import render_perf_markdown
from excelbench.perf.runner import (
    PerfConfig,
    PerfFeatureResult,
    PerfMetadata,
    PerfOpResult,
    PerfResults,
    PerfStats,
)


def _write_cell_values_suite(test_dir: Path) -> None:
    tier_dir = test_dir / "tier1"
    tier_dir.mkdir(parents=True, exist_ok=True)
    filename = "01_cell_values.xlsx"

    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "cell_values"
    ws["B2"] = "Hello"
    wb.save(tier_dir / filename)

    manifest = Manifest(
        generated_at=datetime.now(UTC),
        excel_version="test",
        generator_version="test",
        file_format="xlsx",
        files=[
            BenchFile(
                path=f"tier1/{filename}",
                feature="cell_values",
                tier=1,
                file_format="xlsx",
                test_cases=[
                    BenchCase(
                        id="string_simple",
                        label="String - simple",
                        row=2,
                        expected={"type": "string", "value": "Hello"},
                        importance=Importance.BASIC,
                    )
                ],
            )
        ],
    )
    write_manifest(manifest, test_dir / "manifest.json")


def test_perf_command_writes_outputs(tmp_path: Path) -> None:
    suite = tmp_path / "suite"
    out = tmp_path / "out"
    _write_cell_values_suite(suite)

    # Call the typer command function directly (no timing assertions).
    perf(
        test_dir=suite,
        output_dir=out,
        features=["cell_values"],
        adapters=["openpyxl"],
        warmup=0,
        iters=1,
        breakdown=False,
        profile="xlsx",
    )

    results_path = out / "perf" / "results.json"
    readme_path = out / "perf" / "README.md"
    csv_path = out / "perf" / "matrix.csv"
    history_path = out / "perf" / "history.jsonl"

    assert results_path.exists()
    assert readme_path.exists()
    assert csv_path.exists()
    assert history_path.exists()

    data = json.loads(results_path.read_text())
    assert data["metadata"]["profile"] == "xlsx"
    assert data["metadata"]["config"]["warmup"] == 0
    assert data["metadata"]["config"]["iters"] == 1
    assert data["metadata"]["config"]["iteration_policy"] == "fixed"
    assert "openpyxl" in data["libraries"]
    assert "run_environment" in data["metadata"]
    assert "cpu_model" in data["metadata"]["run_environment"]


def test_perf_markdown_header_matches_all_rendered_cells(tmp_path: Path) -> None:
    stats = PerfStats(min=1.0, p50=1.0, p95=1.0)
    results = PerfResults(
        metadata=PerfMetadata(
            benchmark_version="test",
            run_date=datetime(2026, 1, 1, tzinfo=UTC),
            excel_version="test",
            platform="test",
            profile="xlsx",
            python="test",
            commit=None,
            config=PerfConfig(
                warmup=0,
                iters=1,
                iteration_policy="fixed",
                breakdown=False,
            ),
        ),
        libraries={
            "openpyxl": {
                "name": "openpyxl",
                "version": "test",
                "language": "python",
                "capabilities": ["read", "write"],
            },
            "python-calamine": {
                "name": "python-calamine",
                "version": "test",
                "language": "python",
                "capabilities": ["read"],
            },
        },
        results=[
            PerfFeatureResult(
                feature="cell_values",
                library="openpyxl",
                workload_size="tiny",
                perf={
                    "read": PerfOpResult(wall_ms=stats, cpu_ms=stats),
                    "write": PerfOpResult(
                        wall_ms=PerfStats(min=2.0, p50=2.0, p95=2.0),
                        cpu_ms=stats,
                    ),
                },
            ),
            PerfFeatureResult(
                feature="cell_values",
                library="python-calamine",
                workload_size="tiny",
                perf={
                    "read": PerfOpResult(
                        wall_ms=PerfStats(min=0.5, p50=0.5, p95=0.5),
                        cpu_ms=stats,
                    ),
                    "write": None,
                },
            ),
        ],
    )

    readme = tmp_path / "README.md"
    render_perf_markdown(results, readme)

    markdown = readme.read_text()
    assert (
        "| Feature | openpyxl (R p50 ms) | openpyxl (W p50 ms) | "
        "python-calamine (R p50 ms) |"
    ) in markdown
    assert "| cell_values | 1.00 | 2.00 | 0.50 |" in markdown
    assert markdown.count("**Tier 0") == 1
    assert "Confidence note:" in markdown
    assert "p50/p95" in markdown
