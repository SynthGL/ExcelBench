#!/usr/bin/env python3
"""Generate throughput/scale performance fixtures.

These fixtures are intended for *performance* benchmarking (excelbench perf), not fidelity.
They use a compact workload spec in `expected.workload` to avoid huge manifests.

Default output is under `test_files/` so it stays gitignored.
"""

from __future__ import annotations

import argparse
from collections.abc import Iterator
from contextlib import contextmanager
from datetime import UTC, datetime
from pathlib import Path

import xlsxwriter
from xlsxwriter.worksheet import Worksheet

from excelbench.generator.generate import write_manifest
from excelbench.models import Importance, Manifest, TestCase, TestFile


@contextmanager
def _xlsx_workbook(path: Path, sheet: str) -> Iterator[tuple[xlsxwriter.Workbook, Worksheet]]:
    """Create an xlsxwriter workbook with a single worksheet, ensuring close on exit."""
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(str(path))
    try:
        ws = wb.add_worksheet(sheet)
        yield wb, ws
    finally:
        wb.close()


def _coord_to_cell(row: int, col: int) -> str:
    letters = ""
    c = col
    while c > 0:
        c, rem = divmod(c - 1, 26)
        letters = chr(65 + rem) + letters
    return f"{letters}{row}"


def _generate_cell_values_grid(
    *,
    path: Path,
    sheet: str,
    rows: int,
    cols: int,
    start: int = 1,
    step: int = 1,
) -> None:
    with _xlsx_workbook(path, sheet) as (_wb, ws):
        value = start
        for r in range(rows):
            for c in range(cols):
                ws.write_number(r, c, value)
                value += step


def _generate_strings_grid(
    *,
    path: Path,
    sheet: str,
    rows: int,
    cols: int,
    prefix: str = "V",
    repeated: bool = False,
    repeated_value: str = "X",
    length: int | None = None,
) -> None:
    with _xlsx_workbook(path, sheet) as (_wb, ws):
        value = 1
        for r in range(rows):
            for c in range(cols):
                if repeated:
                    s = repeated_value
                else:
                    s = f"{prefix}{value}"
                if length is not None and length > 0:
                    if len(s) < length:
                        s = s + ("x" * (length - len(s)))
                    else:
                        s = s[:length]
                ws.write_string(r, c, s)
                value += 1


def _generate_formulas_grid(
    *,
    path: Path,
    sheet: str,
    rows: int,
    cols: int,
    formula: str = "=1+1",
) -> None:
    with _xlsx_workbook(path, sheet) as (_wb, ws):
        for r in range(rows):
            for c in range(cols):
                ws.write_formula(r, c, formula)


def _generate_bg_colors_grid(
    *,
    path: Path,
    sheet: str,
    rows: int,
    cols: int,
    palette: list[str],
) -> None:
    with _xlsx_workbook(path, sheet) as (wb, ws):
        fmts = [wb.add_format({"bg_color": f"#{c}", "pattern": 1}) for c in palette]
        for r in range(rows):
            for c in range(cols):
                fmt = fmts[(r * cols + c) % len(fmts)]
                ws.write_string(r, c, "Color", fmt)


def _generate_number_formats_grid(
    *,
    path: Path,
    sheet: str,
    rows: int,
    cols: int,
    number_format: str,
) -> None:
    with _xlsx_workbook(path, sheet) as (wb, ws):
        fmt = wb.add_format({"num_format": number_format})
        value = 0.5
        for r in range(rows):
            for c in range(cols):
                ws.write_number(r, c, value, fmt)
                value += 1.0


def _generate_alignment_grid(
    *,
    path: Path,
    sheet: str,
    rows: int,
    cols: int,
    h_align: str,
    v_align: str,
    wrap: bool,
) -> None:
    with _xlsx_workbook(path, sheet) as (wb, ws):
        fmt_dict: dict[str, object] = {
            "align": h_align,
            "valign": v_align,
        }
        if wrap:
            fmt_dict["text_wrap"] = True
        fmt = wb.add_format(fmt_dict)
        for r in range(rows):
            for c in range(cols):
                ws.write_string(r, c, "Align", fmt)


def _generate_borders_grid(
    *,
    path: Path,
    sheet: str,
    rows: int,
    cols: int,
    border_style: str,
) -> None:
    with _xlsx_workbook(path, sheet) as (wb, ws):
        # Map a small subset of styles.
        border_map = {"thin": 1, "medium": 2, "thick": 5, "double": 6}
        border_val = border_map.get(border_style, 1)
        fmt = wb.add_format({"border": border_val})
        for r in range(rows):
            for c in range(cols):
                ws.write_string(r, c, "Border", fmt)


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate ExcelBench throughput fixtures")
    parser.add_argument(
        "--output",
        "-o",
        type=Path,
        default=Path("test_files/throughput_xlsx"),
        help="Output directory (default: test_files/throughput_xlsx)",
    )
    parser.add_argument(
        "--include-100k",
        action="store_true",
        help="Also generate a ~100k-cell fixture (can take a while).",
    )
    args = parser.parse_args()

    out = Path(args.output)
    tier_dir = out / "tier0"
    tier_dir.mkdir(parents=True, exist_ok=True)

    files: list[TestFile] = []

    # 10k = 100x100
    scenario = "cell_values_10k"
    sheet = "S1"
    rows, cols = 100, 100
    end_cell = _coord_to_cell(rows, cols)
    rng = f"A1:{end_cell}"
    filename = "00_cell_values_10k.xlsx"
    _generate_cell_values_grid(path=tier_dir / filename, sheet=sheet, rows=rows, cols=cols)
    files.append(
        TestFile(
            path=f"tier0/{filename}",
            feature=scenario,
            tier=0,
            file_format="xlsx",
            test_cases=[
                TestCase(
                    id=scenario,
                    label="Throughput: cell values (10k cells)",
                    row=1,
                    expected={
                        "workload": {
                            "scenario": scenario,
                            "op": "cell_value",
                            "sheet": sheet,
                            "range": rng,
                            "start": 1,
                            "step": 1,
                        }
                    },
                    importance=Importance.BASIC,
                )
            ],
        )
    )

    # Bulk read variant (same file, bulk API if adapter supports it)
    files.append(
        TestFile(
            path=f"tier0/{filename}",
            feature="cell_values_10k_bulk_read",
            tier=0,
            file_format="xlsx",
            test_cases=[
                TestCase(
                    id="cell_values_10k_bulk_read",
                    label="Throughput: cell values bulk read (10k cells)",
                    row=1,
                    expected={
                        "workload": {
                            "scenario": "cell_values_10k_bulk_read",
                            "op": "bulk_sheet_values",
                            "operations": ["read"],
                            "sheet": sheet,
                            "range": rng,
                        }
                    },
                    importance=Importance.BASIC,
                )
            ],
        )
    )

    # Bulk write variant (create -> bulk write -> save)
    files.append(
        TestFile(
            path=f"tier0/{filename}",
            feature="cell_values_10k_bulk_write",
            tier=0,
            file_format="xlsx",
            test_cases=[
                TestCase(
                    id="cell_values_10k_bulk_write",
                    label="Throughput: cell values bulk write (10k cells)",
                    row=1,
                    expected={
                        "workload": {
                            "scenario": "cell_values_10k_bulk_write",
                            "op": "bulk_write_grid",
                            "operations": ["write"],
                            "sheet": sheet,
                            "range": rng,
                            "start": 1,
                            "step": 1,
                        }
                    },
                    importance=Importance.BASIC,
                )
            ],
        )
    )

    # Bulk write sparse variant: fill 1% of cells (still 10k range)
    files.append(
        TestFile(
            path=f"tier0/{filename}",
            feature="cell_values_10k_sparse_1pct_bulk_write",
            tier=0,
            file_format="xlsx",
            test_cases=[
                TestCase(
                    id="cell_values_10k_sparse_1pct_bulk_write",
                    label="Throughput: cell values bulk write (10k range, sparse 1%)",
                    row=1,
                    expected={
                        "workload": {
                            "scenario": "cell_values_10k_sparse_1pct_bulk_write",
                            "op": "bulk_write_grid",
                            "operations": ["write"],
                            "sheet": sheet,
                            "range": rng,
                            "start": 1,
                            "step": 1,
                            "sparse_every": 100,
                        }
                    },
                    importance=Importance.BASIC,
                )
            ],
        )
    )

    # 1k = 40x25 (useful for very slow per-cell readers)
    scenario = "cell_values_1k"
    sheet = "S1"
    rows, cols = 40, 25
    end_cell = _coord_to_cell(rows, cols)
    rng = f"A1:{end_cell}"
    filename = "00_cell_values_1k.xlsx"
    _generate_cell_values_grid(path=tier_dir / filename, sheet=sheet, rows=rows, cols=cols)
    files.append(
        TestFile(
            path=f"tier0/{filename}",
            feature=scenario,
            tier=0,
            file_format="xlsx",
            test_cases=[
                TestCase(
                    id=scenario,
                    label="Throughput: cell values (1k cells)",
                    row=1,
                    expected={
                        "workload": {
                            "scenario": scenario,
                            "op": "cell_value",
                            "sheet": sheet,
                            "range": rng,
                            "start": 1,
                            "step": 1,
                        }
                    },
                    importance=Importance.BASIC,
                )
            ],
        )
    )

    # Bulk read variant (same file, bulk API if adapter supports it)
    files.append(
        TestFile(
            path=f"tier0/{filename}",
            feature="cell_values_1k_bulk_read",
            tier=0,
            file_format="xlsx",
            test_cases=[
                TestCase(
                    id="cell_values_1k_bulk_read",
                    label="Throughput: cell values bulk read (1k cells)",
                    row=1,
                    expected={
                        "workload": {
                            "scenario": "cell_values_1k_bulk_read",
                            "op": "bulk_sheet_values",
                            "operations": ["read"],
                            "sheet": sheet,
                            "range": rng,
                        }
                    },
                    importance=Importance.BASIC,
                )
            ],
        )
    )

    # Bulk write variant (create -> bulk write -> save)
    files.append(
        TestFile(
            path=f"tier0/{filename}",
            feature="cell_values_1k_bulk_write",
            tier=0,
            file_format="xlsx",
            test_cases=[
                TestCase(
                    id="cell_values_1k_bulk_write",
                    label="Throughput: cell values bulk write (1k cells)",
                    row=1,
                    expected={
                        "workload": {
                            "scenario": "cell_values_1k_bulk_write",
                            "op": "bulk_write_grid",
                            "operations": ["write"],
                            "sheet": sheet,
                            "range": rng,
                            "start": 1,
                            "step": 1,
                        }
                    },
                    importance=Importance.BASIC,
                )
            ],
        )
    )

    # 10k formulas = 100x100
    scenario = "formulas_10k"
    sheet = "S1"
    rows, cols = 100, 100
    end_cell = _coord_to_cell(rows, cols)
    rng = f"A1:{end_cell}"
    filename = "00_formulas_10k.xlsx"
    formula = "=1+1"
    _generate_formulas_grid(
        path=tier_dir / filename,
        sheet=sheet,
        rows=rows,
        cols=cols,
        formula=formula,
    )
    files.append(
        TestFile(
            path=f"tier0/{filename}",
            feature=scenario,
            tier=0,
            file_format="xlsx",
            test_cases=[
                TestCase(
                    id=scenario,
                    label="Throughput: formulas (10k cells)",
                    row=1,
                    expected={
                        "workload": {
                            "scenario": scenario,
                            "op": "formula",
                            "sheet": sheet,
                            "range": rng,
                            "formula": formula,
                        }
                    },
                    importance=Importance.BASIC,
                )
            ],
        )
    )

    # Bulk read variant (same file, bulk API if adapter supports it)
    files.append(
        TestFile(
            path=f"tier0/{filename}",
            feature="formulas_10k_bulk_read",
            tier=0,
            file_format="xlsx",
            test_cases=[
                TestCase(
                    id="formulas_10k_bulk_read",
                    label="Throughput: formulas bulk read (10k cells)",
                    row=1,
                    expected={
                        "workload": {
                            "scenario": "formulas_10k_bulk_read",
                            "op": "bulk_sheet_values",
                            "operations": ["read"],
                            "sheet": sheet,
                            "range": rng,
                        }
                    },
                    importance=Importance.BASIC,
                )
            ],
        )
    )

    # 1k formulas = 40x25
    scenario = "formulas_1k"
    sheet = "S1"
    rows, cols = 40, 25
    end_cell = _coord_to_cell(rows, cols)
    rng = f"A1:{end_cell}"
    filename = "00_formulas_1k.xlsx"
    formula = "=1+1"
    _generate_formulas_grid(
        path=tier_dir / filename,
        sheet=sheet,
        rows=rows,
        cols=cols,
        formula=formula,
    )
    files.append(
        TestFile(
            path=f"tier0/{filename}",
            feature=scenario,
            tier=0,
            file_format="xlsx",
            test_cases=[
                TestCase(
                    id=scenario,
                    label="Throughput: formulas (1k cells)",
                    row=1,
                    expected={
                        "workload": {
                            "scenario": scenario,
                            "op": "formula",
                            "sheet": sheet,
                            "range": rng,
                            "formula": formula,
                        }
                    },
                    importance=Importance.BASIC,
                )
            ],
        )
    )

    # Bulk read variant (same file, bulk API if adapter supports it)

    files.append(
        TestFile(
            path=f"tier0/{filename}",
            feature="formulas_1k_bulk_read",
            tier=0,
            file_format="xlsx",
            test_cases=[
                TestCase(
                    id="formulas_1k_bulk_read",
                    label="Throughput: formulas bulk read (1k cells)",
                    row=1,
                    expected={
                        "workload": {
                            "scenario": "formulas_1k_bulk_read",
                            "op": "bulk_sheet_values",
                            "operations": ["read"],
                            "sheet": sheet,
                            "range": rng,
                        }
                    },
                    importance=Importance.BASIC,
                )
            ],
        )
    )

    # 10k cell values, tall (1000x10) — bulk read/write
    sheet = "S1"
    rows, cols = 1000, 10
    end_cell = _coord_to_cell(rows, cols)
    rng = f"A1:{end_cell}"
    filename = "00_cell_values_10k_1000x10.xlsx"
    _generate_cell_values_grid(path=tier_dir / filename, sheet=sheet, rows=rows, cols=cols)
    files.append(
        TestFile(
            path=f"tier0/{filename}",
            feature="cell_values_10k_1000x10_bulk_read",
            tier=0,
            file_format="xlsx",
            test_cases=[
                TestCase(
                    id="cell_values_10k_1000x10_bulk_read",
                    label="Throughput: cell values bulk read (10k cells, 1000x10)",
                    row=1,
                    expected={
                        "workload": {
                            "scenario": "cell_values_10k_1000x10_bulk_read",
                            "op": "bulk_sheet_values",
                            "operations": ["read"],
                            "sheet": sheet,
                            "range": rng,
                        }
                    },
                    importance=Importance.BASIC,
                )
            ],
        )
    )
    files.append(
        TestFile(
            path=f"tier0/{filename}",
            feature="cell_values_10k_1000x10_bulk_write",
            tier=0,
            file_format="xlsx",
            test_cases=[
                TestCase(
                    id="cell_values_10k_1000x10_bulk_write",
                    label="Throughput: cell values bulk write (10k cells, 1000x10)",
                    row=1,
                    expected={
                        "workload": {
                            "scenario": "cell_values_10k_1000x10_bulk_write",
                            "op": "bulk_write_grid",
                            "operations": ["write"],
                            "sheet": sheet,
                            "range": rng,
                            "start": 1,
                            "step": 1,
                        }
                    },
                    importance=Importance.BASIC,
                )
            ],
        )
    )

    # 10k cell values, wide (10x1000) — bulk read/write
    sheet = "S1"
    rows, cols = 10, 1000
    end_cell = _coord_to_cell(rows, cols)
    rng = f"A1:{end_cell}"
    filename = "00_cell_values_10k_10x1000.xlsx"
    _generate_cell_values_grid(path=tier_dir / filename, sheet=sheet, rows=rows, cols=cols)
    files.append(
        TestFile(
            path=f"tier0/{filename}",
            feature="cell_values_10k_10x1000_bulk_read",
            tier=0,
            file_format="xlsx",
            test_cases=[
                TestCase(
                    id="cell_values_10k_10x1000_bulk_read",
                    label="Throughput: cell values bulk read (10k cells, 10x1000)",
                    row=1,
                    expected={
                        "workload": {
                            "scenario": "cell_values_10k_10x1000_bulk_read",
                            "op": "bulk_sheet_values",
                            "operations": ["read"],
                            "sheet": sheet,
                            "range": rng,
                        }
                    },
                    importance=Importance.BASIC,
                )
            ],
        )
    )
    files.append(
        TestFile(
            path=f"tier0/{filename}",
            feature="cell_values_10k_10x1000_bulk_write",
            tier=0,
            file_format="xlsx",
            test_cases=[
                TestCase(
                    id="cell_values_10k_10x1000_bulk_write",
                    label="Throughput: cell values bulk write (10k cells, 10x1000)",
                    row=1,
                    expected={
                        "workload": {
                            "scenario": "cell_values_10k_10x1000_bulk_write",
                            "op": "bulk_write_grid",
                            "operations": ["write"],
                            "sheet": sheet,
                            "range": rng,
                            "start": 1,
                            "step": 1,
                        }
                    },
                    importance=Importance.BASIC,
                )
            ],
        )
    )

    # 1k strings, unique (40x25) — bulk read/write
    sheet = "S1"
    rows, cols = 40, 25
    end_cell = _coord_to_cell(rows, cols)
    rng = f"A1:{end_cell}"
    filename = "00_strings_unique_1k.xlsx"
    _generate_strings_grid(path=tier_dir / filename, sheet=sheet, rows=rows, cols=cols, prefix="V")
    files.append(
        TestFile(
            path=f"tier0/{filename}",
            feature="strings_unique_1k_bulk_read",
            tier=0,
            file_format="xlsx",
            test_cases=[
                TestCase(
                    id="strings_unique_1k_bulk_read",
                    label="Throughput: strings bulk read (1k cells, unique)",
                    row=1,
                    expected={
                        "workload": {
                            "scenario": "strings_unique_1k_bulk_read",
                            "op": "bulk_sheet_values",
                            "operations": ["read"],
                            "sheet": sheet,
                            "range": rng,
                        }
                    },
                    importance=Importance.BASIC,
                )
            ],
        )
    )
    files.append(
        TestFile(
            path=f"tier0/{filename}",
            feature="strings_unique_1k_bulk_write",
            tier=0,
            file_format="xlsx",
            test_cases=[
                TestCase(
                    id="strings_unique_1k_bulk_write",
                    label="Throughput: strings bulk write (1k cells, unique)",
                    row=1,
                    expected={
                        "workload": {
                            "scenario": "strings_unique_1k_bulk_write",
                            "op": "bulk_write_grid",
                            "operations": ["write"],
                            "sheet": sheet,
                            "range": rng,
                            "value_type": "string",
                            "string_prefix": "V",
                            "start": 1,
                            "step": 1,
                        }
                    },
                    importance=Importance.BASIC,
                )
            ],
        )
    )

    # 1k strings, long payload (unique) — bulk read/write
    sheet = "S1"
    rows, cols = 40, 25
    end_cell = _coord_to_cell(rows, cols)
    rng = f"A1:{end_cell}"
    for length in (64, 256):
        filename = f"00_strings_unique_1k_len{length}.xlsx"
        _generate_strings_grid(
            path=tier_dir / filename,
            sheet=sheet,
            rows=rows,
            cols=cols,
            prefix="V",
            length=length,
        )
        files.append(
            TestFile(
                path=f"tier0/{filename}",
                feature=f"strings_unique_1k_len{length}_bulk_read",
                tier=0,
                file_format="xlsx",
                test_cases=[
                    TestCase(
                        id=f"strings_unique_1k_len{length}_bulk_read",
                        label=f"Throughput: strings bulk read (1k cells, unique, len {length})",
                        row=1,
                        expected={
                            "workload": {
                                "scenario": f"strings_unique_1k_len{length}_bulk_read",
                                "op": "bulk_sheet_values",
                                "operations": ["read"],
                                "sheet": sheet,
                                "range": rng,
                            }
                        },
                        importance=Importance.BASIC,
                    )
                ],
            )
        )
        files.append(
            TestFile(
                path=f"tier0/{filename}",
                feature=f"strings_unique_1k_len{length}_bulk_write",
                tier=0,
                file_format="xlsx",
                test_cases=[
                    TestCase(
                        id=f"strings_unique_1k_len{length}_bulk_write",
                        label=f"Throughput: strings bulk write (1k cells, unique, len {length})",
                        row=1,
                        expected={
                            "workload": {
                                "scenario": f"strings_unique_1k_len{length}_bulk_write",
                                "op": "bulk_write_grid",
                                "operations": ["write"],
                                "sheet": sheet,
                                "range": rng,
                                "value_type": "string",
                                "string_prefix": "V",
                                "string_length": length,
                                "start": 1,
                                "step": 1,
                            }
                        },
                        importance=Importance.BASIC,
                    )
                ],
            )
        )

    # 1k strings, long payload (repeated) — bulk read/write
    sheet = "S1"
    rows, cols = 40, 25
    end_cell = _coord_to_cell(rows, cols)
    rng = f"A1:{end_cell}"
    length = 256
    filename = f"00_strings_repeated_1k_len{length}.xlsx"
    _generate_strings_grid(
        path=tier_dir / filename,
        sheet=sheet,
        rows=rows,
        cols=cols,
        repeated=True,
        repeated_value="X",
        length=length,
    )
    files.append(
        TestFile(
            path=f"tier0/{filename}",
            feature=f"strings_repeated_1k_len{length}_bulk_read",
            tier=0,
            file_format="xlsx",
            test_cases=[
                TestCase(
                    id=f"strings_repeated_1k_len{length}_bulk_read",
                    label=f"Throughput: strings bulk read (1k cells, repeated, len {length})",
                    row=1,
                    expected={
                        "workload": {
                            "scenario": f"strings_repeated_1k_len{length}_bulk_read",
                            "op": "bulk_sheet_values",
                            "operations": ["read"],
                            "sheet": sheet,
                            "range": rng,
                        }
                    },
                    importance=Importance.BASIC,
                )
            ],
        )
    )
    files.append(
        TestFile(
            path=f"tier0/{filename}",
            feature=f"strings_repeated_1k_len{length}_bulk_write",
            tier=0,
            file_format="xlsx",
            test_cases=[
                TestCase(
                    id=f"strings_repeated_1k_len{length}_bulk_write",
                    label=f"Throughput: strings bulk write (1k cells, repeated, len {length})",
                    row=1,
                    expected={
                        "workload": {
                            "scenario": f"strings_repeated_1k_len{length}_bulk_write",
                            "op": "bulk_write_grid",
                            "operations": ["write"],
                            "sheet": sheet,
                            "range": rng,
                            "value_type": "string",
                            "string_mode": "repeated",
                            "string_value": "X",
                            "string_length": length,
                            "start": 1,
                            "step": 1,
                        }
                    },
                    importance=Importance.BASIC,
                )
            ],
        )
    )

    # 10k strings, unique (100x100) — bulk read/write
    sheet = "S1"
    rows, cols = 100, 100
    end_cell = _coord_to_cell(rows, cols)
    rng = f"A1:{end_cell}"
    filename = "00_strings_unique_10k.xlsx"
    _generate_strings_grid(path=tier_dir / filename, sheet=sheet, rows=rows, cols=cols, prefix="V")
    files.append(
        TestFile(
            path=f"tier0/{filename}",
            feature="strings_unique_10k_bulk_read",
            tier=0,
            file_format="xlsx",
            test_cases=[
                TestCase(
                    id="strings_unique_10k_bulk_read",
                    label="Throughput: strings bulk read (10k cells, unique)",
                    row=1,
                    expected={
                        "workload": {
                            "scenario": "strings_unique_10k_bulk_read",
                            "op": "bulk_sheet_values",
                            "operations": ["read"],
                            "sheet": sheet,
                            "range": rng,
                        }
                    },
                    importance=Importance.BASIC,
                )
            ],
        )
    )
    files.append(
        TestFile(
            path=f"tier0/{filename}",
            feature="strings_unique_10k_bulk_write",
            tier=0,
            file_format="xlsx",
            test_cases=[
                TestCase(
                    id="strings_unique_10k_bulk_write",
                    label="Throughput: strings bulk write (10k cells, unique)",
                    row=1,
                    expected={
                        "workload": {
                            "scenario": "strings_unique_10k_bulk_write",
                            "op": "bulk_write_grid",
                            "operations": ["write"],
                            "sheet": sheet,
                            "range": rng,
                            "value_type": "string",
                            "string_prefix": "V",
                            "start": 1,
                            "step": 1,
                        }
                    },
                    importance=Importance.BASIC,
                )
            ],
        )
    )

    # 10k strings, repeated (100x100) — bulk read/write
    sheet = "S1"
    rows, cols = 100, 100
    end_cell = _coord_to_cell(rows, cols)
    rng = f"A1:{end_cell}"
    filename = "00_strings_repeated_10k.xlsx"
    _generate_strings_grid(
        path=tier_dir / filename,
        sheet=sheet,
        rows=rows,
        cols=cols,
        repeated=True,
        repeated_value="X",
    )
    files.append(
        TestFile(
            path=f"tier0/{filename}",
            feature="strings_repeated_10k_bulk_read",
            tier=0,
            file_format="xlsx",
            test_cases=[
                TestCase(
                    id="strings_repeated_10k_bulk_read",
                    label="Throughput: strings bulk read (10k cells, repeated)",
                    row=1,
                    expected={
                        "workload": {
                            "scenario": "strings_repeated_10k_bulk_read",
                            "op": "bulk_sheet_values",
                            "operations": ["read"],
                            "sheet": sheet,
                            "range": rng,
                        }
                    },
                    importance=Importance.BASIC,
                )
            ],
        )
    )
    files.append(
        TestFile(
            path=f"tier0/{filename}",
            feature="strings_repeated_10k_bulk_write",
            tier=0,
            file_format="xlsx",
            test_cases=[
                TestCase(
                    id="strings_repeated_10k_bulk_write",
                    label="Throughput: strings bulk write (10k cells, repeated)",
                    row=1,
                    expected={
                        "workload": {
                            "scenario": "strings_repeated_10k_bulk_write",
                            "op": "bulk_write_grid",
                            "operations": ["write"],
                            "sheet": sheet,
                            "range": rng,
                            "value_type": "string",
                            "string_mode": "repeated",
                            "string_value": "X",
                            "start": 1,
                            "step": 1,
                        }
                    },
                    importance=Importance.BASIC,
                )
            ],
        )
    )
    # 1k background fills = 40x25
    scenario = "background_colors_1k"
    sheet = "S1"
    rows, cols = 40, 25
    end_cell = _coord_to_cell(rows, cols)
    rng = f"A1:{end_cell}"
    filename = "00_background_colors_1k.xlsx"
    palette = ["FF0000", "00FF00", "0000FF", "FFFF00"]
    _generate_bg_colors_grid(
        path=tier_dir / filename,
        sheet=sheet,
        rows=rows,
        cols=cols,
        palette=palette,
    )
    files.append(
        TestFile(
            path=f"tier0/{filename}",
            feature=scenario,
            tier=0,
            file_format="xlsx",
            test_cases=[
                TestCase(
                    id=scenario,
                    label="Throughput: background fills (1k cells)",
                    row=1,
                    expected={
                        "workload": {
                            "scenario": scenario,
                            "op": "bg_color",
                            "sheet": sheet,
                            "range": rng,
                            "palette": [f"#{c}" for c in palette],
                        }
                    },
                    importance=Importance.BASIC,
                )
            ],
        )
    )

    # 1k number formats = 40x25
    scenario = "number_formats_1k"
    sheet = "S1"
    rows, cols = 40, 25
    end_cell = _coord_to_cell(rows, cols)
    rng = f"A1:{end_cell}"
    filename = "00_number_formats_1k.xlsx"
    number_format = "0.00%"
    _generate_number_formats_grid(
        path=tier_dir / filename,
        sheet=sheet,
        rows=rows,
        cols=cols,
        number_format=number_format,
    )
    files.append(
        TestFile(
            path=f"tier0/{filename}",
            feature=scenario,
            tier=0,
            file_format="xlsx",
            test_cases=[
                TestCase(
                    id=scenario,
                    label="Throughput: number formats (1k cells)",
                    row=1,
                    expected={
                        "workload": {
                            "scenario": scenario,
                            "op": "number_format",
                            "sheet": sheet,
                            "range": rng,
                            "number_format": number_format,
                        }
                    },
                    importance=Importance.BASIC,
                )
            ],
        )
    )

    # 1k alignment = 40x25
    scenario = "alignment_1k"
    sheet = "S1"
    rows, cols = 40, 25
    end_cell = _coord_to_cell(rows, cols)
    rng = f"A1:{end_cell}"
    filename = "00_alignment_1k.xlsx"
    _generate_alignment_grid(
        path=tier_dir / filename,
        sheet=sheet,
        rows=rows,
        cols=cols,
        h_align="center",
        v_align="top",
        wrap=True,
    )
    files.append(
        TestFile(
            path=f"tier0/{filename}",
            feature=scenario,
            tier=0,
            file_format="xlsx",
            test_cases=[
                TestCase(
                    id=scenario,
                    label="Throughput: alignment (1k cells)",
                    row=1,
                    expected={
                        "workload": {
                            "scenario": scenario,
                            "op": "alignment",
                            "sheet": sheet,
                            "range": rng,
                            "h_align": "center",
                            "v_align": "top",
                            "wrap": True,
                        }
                    },
                    importance=Importance.BASIC,
                )
            ],
        )
    )

    # 200 borders = 20x10
    scenario = "borders_200"
    sheet = "S1"
    rows, cols = 20, 10
    end_cell = _coord_to_cell(rows, cols)
    rng = f"A1:{end_cell}"
    filename = "00_borders_200.xlsx"
    _generate_borders_grid(
        path=tier_dir / filename,
        sheet=sheet,
        rows=rows,
        cols=cols,
        border_style="thin",
    )
    files.append(
        TestFile(
            path=f"tier0/{filename}",
            feature=scenario,
            tier=0,
            file_format="xlsx",
            test_cases=[
                TestCase(
                    id=scenario,
                    label="Throughput: borders (200 cells)",
                    row=1,
                    expected={
                        "workload": {
                            "scenario": scenario,
                            "op": "border",
                            "sheet": sheet,
                            "range": rng,
                            "border_style": "thin",
                            "border_color": "#000000",
                        }
                    },
                    importance=Importance.BASIC,
                )
            ],
        )
    )

    if args.include_100k:
        # ~100k = 316x316 = 99856 cells
        scenario = "cell_values_100k"
        sheet = "S1"
        rows, cols = 316, 316
        end_cell = _coord_to_cell(rows, cols)
        rng = f"A1:{end_cell}"
        filename = "00_cell_values_100k.xlsx"
        _generate_cell_values_grid(path=tier_dir / filename, sheet=sheet, rows=rows, cols=cols)
        files.append(
            TestFile(
                path=f"tier0/{filename}",
                feature=scenario,
                tier=0,
                file_format="xlsx",
                test_cases=[
                    TestCase(
                        id=scenario,
                        label="Throughput: cell values (~100k cells)",
                        row=1,
                        expected={
                            "workload": {
                                "scenario": scenario,
                                "op": "cell_value",
                                "sheet": sheet,
                                "range": rng,
                                "start": 1,
                                "step": 1,
                            }
                        },
                        importance=Importance.BASIC,
                    )
                ],
            )
        )

        # Bulk read/write variants for the ~100k fixture.
        files.append(
            TestFile(
                path=f"tier0/{filename}",
                feature="cell_values_100k_bulk_read",
                tier=0,
                file_format="xlsx",
                test_cases=[
                    TestCase(
                        id="cell_values_100k_bulk_read",
                        label="Throughput: cell values bulk read (~100k cells)",
                        row=1,
                        expected={
                            "workload": {
                                "scenario": "cell_values_100k_bulk_read",
                                "op": "bulk_sheet_values",
                                "operations": ["read"],
                                "sheet": sheet,
                                "range": rng,
                            }
                        },
                        importance=Importance.BASIC,
                    )
                ],
            )
        )
        files.append(
            TestFile(
                path=f"tier0/{filename}",
                feature="cell_values_100k_bulk_write",
                tier=0,
                file_format="xlsx",
                test_cases=[
                    TestCase(
                        id="cell_values_100k_bulk_write",
                        label="Throughput: cell values bulk write (~100k cells)",
                        row=1,
                        expected={
                            "workload": {
                                "scenario": "cell_values_100k_bulk_write",
                                "op": "bulk_write_grid",
                                "operations": ["write"],
                                "sheet": sheet,
                                "range": rng,
                                "start": 1,
                                "step": 1,
                            }
                        },
                        importance=Importance.BASIC,
                    )
                ],
            )
        )

    manifest = Manifest(
        generated_at=datetime.now(UTC),
        excel_version="xlsxwriter-generated",
        generator_version="throughput-0.1.0",
        file_format="xlsx",
        files=files,
    )
    write_manifest(manifest, out / "manifest.json")

    print(f"✓ Wrote {len(files)} throughput fixture(s) to {out}")
    print(f"  Manifest: {out / 'manifest.json'}")


if __name__ == "__main__":
    main()
