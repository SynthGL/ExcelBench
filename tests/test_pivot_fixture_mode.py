from datetime import UTC, datetime
from pathlib import Path

from openpyxl import Workbook

from excelbench.generator.generate import write_manifest
from excelbench.harness.adapters.base import ReadOnlyAdapter
from excelbench.harness.runner import run_benchmark
from excelbench.models import (
    BorderInfo,
    CellFormat,
    CellType,
    CellValue,
    Importance,
    LibraryInfo,
    Manifest,
)
from excelbench.models import (
    TestCase as BenchCase,
)
from excelbench.models import (
    TestFile as BenchFile,
)


class StubPivotAdapter(ReadOnlyAdapter):
    def __init__(self, pivots: list[dict]):
        self._pivots = pivots

    @property
    def info(self) -> LibraryInfo:
        return LibraryInfo(
            name="stub-pivot",
            version="1.0.0",
            language="python",
            capabilities={"read"},
        )

    @property
    def supported_read_extensions(self) -> set[str]:
        return {".xlsx"}

    def open_workbook(self, path: Path):
        return {"path": str(path)}

    def close_workbook(self, workbook) -> None:
        return None

    def get_sheet_names(self, workbook) -> list[str]:
        return ["Pivot"]

    def read_cell_value(self, workbook, sheet: str, cell: str) -> CellValue:
        return CellValue(type=CellType.BLANK)

    def read_cell_format(self, workbook, sheet: str, cell: str) -> CellFormat:
        return CellFormat()

    def read_cell_border(self, workbook, sheet: str, cell: str) -> BorderInfo:
        return BorderInfo()

    def read_row_height(self, workbook, sheet: str, row: int) -> float | None:
        return None

    def read_column_width(self, workbook, sheet: str, column: str) -> float | None:
        return None

    def read_merged_ranges(self, workbook, sheet: str) -> list[str]:
        return []

    def read_conditional_formats(self, workbook, sheet: str) -> list[dict]:
        return []

    def read_data_validations(self, workbook, sheet: str) -> list[dict]:
        return []

    def read_hyperlinks(self, workbook, sheet: str) -> list[dict]:
        return []

    def read_images(self, workbook, sheet: str) -> list[dict]:
        return []

    def read_pivot_tables(self, workbook, sheet: str) -> list[dict]:
        return list(self._pivots)

    def read_comments(self, workbook, sheet: str) -> list[dict]:
        return []

    def read_freeze_panes(self, workbook, sheet: str) -> dict:
        return {}


def _write_pivot_workbook(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Pivot"
    wb.save(path)


def test_pivot_fixture_absent_keeps_explicit_na_note(tmp_path, monkeypatch):
    monkeypatch.setattr("excelbench.harness.runner.platform.system", lambda: "Darwin")

    test_dir = tmp_path / "tests"
    tier2 = test_dir / "tier2"
    tier2.mkdir(parents=True)
    workbook_path = tier2 / "15_pivot_tables.xlsx"
    _write_pivot_workbook(workbook_path)

    manifest = Manifest(
        generated_at=datetime.now(UTC),
        excel_version="test",
        generator_version="test",
        file_format="xlsx",
        files=[
            BenchFile(
                path="tier2/15_pivot_tables.xlsx",
                feature="pivot_tables",
                tier=2,
                file_format="xlsx",
                test_cases=[],
            )
        ],
    )
    write_manifest(manifest, test_dir / "manifest.json")

    results = run_benchmark(test_dir, adapters=[StubPivotAdapter([])], profile="xlsx")
    assert len(results.scores) == 1
    score = results.scores[0]
    assert score.read_score is None
    assert score.write_score is None
    assert "Unsupported on macOS without a Windows-generated pivot fixture" in (score.notes or "")


def test_pivot_fixture_present_executes_read_path(tmp_path, monkeypatch):
    monkeypatch.setattr("excelbench.harness.runner.platform.system", lambda: "Darwin")

    test_dir = tmp_path / "tests"
    tier2 = test_dir / "tier2"
    tier2.mkdir(parents=True)
    workbook_path = tier2 / "15_pivot_tables.xlsx"
    _write_pivot_workbook(workbook_path)

    expected = {
        "pivot": {
            "name": "SalesPivot",
            "source_range": "Data!A1:D6",
            "target_cell": "Pivot!B3",
        }
    }
    manifest = Manifest(
        generated_at=datetime.now(UTC),
        excel_version="test",
        generator_version="test",
        file_format="xlsx",
        files=[
            BenchFile(
                path="tier2/15_pivot_tables.xlsx",
                feature="pivot_tables",
                tier=2,
                file_format="xlsx",
                test_cases=[
                    BenchCase(
                        id="pivot_basic",
                        label="Pivot: basic layout",
                        row=2,
                        expected=expected,
                        sheet="Pivot",
                        importance=Importance.BASIC,
                    )
                ],
            )
        ],
    )
    write_manifest(manifest, test_dir / "manifest.json")

    adapter = StubPivotAdapter(
        [{"name": "SalesPivot", "source_range": "Data!A1:D6", "target_cell": "Pivot!B3"}]
    )
    results = run_benchmark(test_dir, adapters=[adapter], profile="xlsx")

    assert len(results.scores) == 1
    score = results.scores[0]
    assert score.read_score == 3
    assert score.notes is None
