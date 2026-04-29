from __future__ import annotations

import json
from pathlib import Path

from openpyxl import Workbook
from typer.testing import CliRunner

from excelbench.cli import app


def _write_xlsx(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet["A1"] = "Revenue"
    workbook.save(path)
    workbook.close()


def test_diff_workbooks_cli(tmp_path: Path) -> None:
    left = tmp_path / "left.xlsx"
    right = tmp_path / "right.xlsx"
    output = tmp_path / "out"
    _write_xlsx(left)
    _write_xlsx(right)

    result = CliRunner().invoke(
        app,
        [
            "diff-workbooks",
            "--left",
            str(left),
            "--right",
            str(right),
            "--output",
            str(output),
        ],
    )

    assert result.exit_code == 0
    assert (output / "summary.json").exists()


def test_roundtrip_context_cli_with_openpyxl(tmp_path: Path) -> None:
    fixtures = tmp_path / "fixtures"
    fixtures.mkdir()
    _write_xlsx(fixtures / "sample.xlsx")
    output = tmp_path / "roundtrip"

    result = CliRunner().invoke(
        app,
        [
            "roundtrip-context",
            "--tests",
            str(fixtures),
            "--output",
            str(output),
            "--adapter",
            "openpyxl",
            "--cycles",
            "1",
        ],
    )

    assert result.exit_code == 0
    payload = json.loads((output / "roundtrip.json").read_text())
    assert payload["results"]
    assert all(row["passed"] for row in payload["results"])


def test_compatibility_context_cli_skips_unknown_adapter(tmp_path: Path) -> None:
    output = tmp_path / "compat"

    result = CliRunner().invoke(
        app,
        [
            "compatibility-context",
            "--output",
            str(output),
            "--adapter",
            "unknown-adapter",
        ],
    )

    assert result.exit_code == 0
    payload = json.loads((output / "compatibility.json").read_text())
    assert payload["results"][0]["skipped"] is True


def test_macro_context_cli_skips_without_fixtures(tmp_path: Path) -> None:
    output = tmp_path / "macros"

    result = CliRunner().invoke(
        app,
        ["macro-context", "--tests", str(tmp_path / "missing"), "--output", str(output)],
    )

    assert result.exit_code == 0
    payload = json.loads((output / "results.json").read_text())
    assert payload["results"][0]["skipped"] is True
