import json
from datetime import UTC, datetime
from pathlib import Path

from openpyxl import Workbook

from excelbench.cli import perf
from excelbench.generator.generate import write_manifest
from excelbench.models import Importance, Manifest
from excelbench.models import TestCase as BenchCase
from excelbench.models import TestFile as BenchFile


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
