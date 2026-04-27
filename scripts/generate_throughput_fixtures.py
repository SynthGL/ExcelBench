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
from datetime import UTC, date, datetime, timedelta
from pathlib import Path
from typing import Any

import xlsxwriter
from xlsxwriter.worksheet import Worksheet

from excelbench.generator.generate import write_manifest
from excelbench.models import Importance, Manifest, TestCase, TestFile

# Data-shape matrix: 10 dtypes × 4 tiers (1M tier gated behind --include-1m).
# See DEC-019 for the rationale and the mixed_realistic ratio choice.
DATA_SHAPE_DTYPES: list[str] = [
    "int",
    "float",
    "string_short",
    "string_long",
    "boolean",
    "date",
    "datetime",
    "formula_simple",
    "formula_cross_sheet",
    "mixed_realistic",
]

DATA_SHAPE_TIERS: list[tuple[str, int, int]] = [
    ("1k", 40, 25),       # 1000 cells
    ("10k", 100, 100),    # 10000 cells
    ("100k", 316, 316),   # ~99856 cells
    ("1m", 1000, 1000),   # 1000000 cells
]


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


def _generate_data_shape_grid(
    *,
    path: Path,
    sheet: str,
    rows: int,
    cols: int,
    dtype: str,
) -> None:
    """Generate one xlsx fixture for a (dtype, tier) pair.

    Per-cell value generation mirrors the runner's `_run_workload_write` so that
    a generated file roundtrips deterministically: the same per-cell-index value
    that gets written by the bench is what's read back at iteration time. Keep
    the two in sync if either side changes.
    """
    with _xlsx_workbook(path, sheet) as (wb, ws):
        if dtype == "int":
            for r in range(rows):
                for c in range(cols):
                    ws.write_number(r, c, r * cols + c + 1)
            return

        if dtype == "float":
            for r in range(rows):
                for c in range(cols):
                    ws.write_number(r, c, (r * cols + c + 1) * 1.5)
            return

        if dtype == "string_short":
            for r in range(rows):
                for c in range(cols):
                    v = r * cols + c + 1
                    s = f"S{v}"
                    if len(s) < 16:
                        s = s + "x" * (16 - len(s))
                    ws.write_string(r, c, s[:16])
            return

        if dtype == "string_long":
            for r in range(rows):
                for c in range(cols):
                    v = r * cols + c + 1
                    s = f"S{v}"
                    if len(s) < 512:
                        s = s + "x" * (512 - len(s))
                    ws.write_string(r, c, s[:512])
            return

        if dtype == "boolean":
            for r in range(rows):
                for c in range(cols):
                    ws.write_boolean(r, c, bool((r * cols + c) % 2))
            return

        if dtype == "date":
            date_fmt = wb.add_format({"num_format": "yyyy-mm-dd"})
            base = date(2020, 1, 1)
            for r in range(rows):
                for c in range(cols):
                    v = r * cols + c
                    ws.write_datetime(r, c, base + timedelta(days=v), date_fmt)
            return

        if dtype == "datetime":
            dt_fmt = wb.add_format({"num_format": "yyyy-mm-dd hh:mm:ss"})
            base = datetime(2020, 1, 1)
            for r in range(rows):
                for c in range(cols):
                    v = r * cols + c
                    ws.write_datetime(r, c, base + timedelta(seconds=v), dt_fmt)
            return

        if dtype == "formula_simple":
            for r in range(rows):
                for c in range(cols):
                    ws.write_formula(r, c, f"=SUM(A{r + 1}:B{r + 1})")
            return

        if dtype == "formula_cross_sheet":
            ws2 = wb.add_worksheet("Sheet2")
            for r in range(rows):
                ws2.write_number(r, 0, r + 1)
            for r in range(rows):
                for c in range(cols):
                    ws.write_formula(r, c, f"=Sheet2!A{r + 1}")
            return

        if dtype == "mixed_realistic":
            # 60% short string, 30% int, 5% date, 3% formula, 2% blank — DEC-019.
            date_fmt = wb.add_format({"num_format": "yyyy-mm-dd"})
            base = date(2020, 1, 1)
            for r in range(rows):
                for c in range(cols):
                    idx = r * cols + c
                    bucket = idx % 100
                    v = idx + 1
                    if bucket < 60:
                        s = f"S{v}"
                        if len(s) < 16:
                            s = s + "x" * (16 - len(s))
                        ws.write_string(r, c, s[:16])
                    elif bucket < 90:
                        ws.write_number(r, c, v)
                    elif bucket < 95:
                        ws.write_datetime(r, c, base + timedelta(days=v), date_fmt)
                    elif bucket < 98:
                        ws.write_formula(r, c, f"=SUM(A{r + 1}:B{r + 1})")
                    else:
                        ws.write_blank(r, c, None)
            return

        raise ValueError(f"Unsupported data-shape dtype: {dtype!r}")


def _data_shape_write_workload(
    *, dtype: str, sheet: str, rng: str, scenario: str
) -> dict[str, Any]:
    """Return the bulk_write_grid workload spec for a (dtype, scenario) pair.

    The spec is interpreted by the runner's `_run_workload_write`. Most dtypes
    map 1:1 to a runner `value_type`; the string sub-variants fold into the
    existing `string` op via `string_length`.
    """
    base: dict[str, Any] = {
        "scenario": scenario,
        "op": "bulk_write_grid",
        "operations": ["write"],
        "sheet": sheet,
        "range": rng,
        "start": 1,
        "step": 1,
    }

    if dtype == "int":
        base["value_type"] = "number"
    elif dtype == "float":
        base["value_type"] = "float"
    elif dtype == "string_short":
        base["value_type"] = "string"
        base["string_length"] = 16
    elif dtype == "string_long":
        base["value_type"] = "string"
        base["string_length"] = 512
    elif dtype in {
        "boolean",
        "date",
        "datetime",
        "formula_simple",
        "formula_cross_sheet",
        "mixed_realistic",
    }:
        base["value_type"] = dtype
    else:
        raise ValueError(f"Unsupported data-shape dtype: {dtype!r}")

    return base


def generate_data_shape_scenarios(
    tier_dir: Path,
    *,
    include_1m: bool = False,
) -> list[TestFile]:
    """Emit (dtype × tier) fixtures and TestFile manifest entries.

    Default emits 10 dtypes × 3 tiers (1k/10k/100k) = 30 fixtures, each with one
    bulk_read and one bulk_write workload entry (60 manifest rows).

    With ``include_1m=True`` adds the 1M tier (10 more fixtures, 20 more rows).
    The 1M tier is gated because xlsxwriter spends ~3s per 100k cells generating
    a fixture; full 1M run is ~30s × 10 dtypes ≈ 5 min on a bench machine.
    """
    files: list[TestFile] = []
    sheet = "S1"

    for dtype in DATA_SHAPE_DTYPES:
        for tier_label, rows, cols in DATA_SHAPE_TIERS:
            if tier_label == "1m" and not include_1m:
                continue

            scenario = f"data_shape_{dtype}_{tier_label}"
            filename = f"00_{scenario}.xlsx"
            end_cell = _coord_to_cell(rows, cols)
            rng = f"A1:{end_cell}"

            _generate_data_shape_grid(
                path=tier_dir / filename,
                sheet=sheet,
                rows=rows,
                cols=cols,
                dtype=dtype,
            )

            read_feature = f"{scenario}_bulk_read"
            files.append(
                TestFile(
                    path=f"tier0/{filename}",
                    feature=read_feature,
                    tier=0,
                    file_format="xlsx",
                    test_cases=[
                        TestCase(
                            id=read_feature,
                            label=f"Data shape: {dtype} bulk read ({tier_label})",
                            row=1,
                            expected={
                                "workload": {
                                    "scenario": read_feature,
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

            write_feature = f"{scenario}_bulk_write"
            files.append(
                TestFile(
                    path=f"tier0/{filename}",
                    feature=write_feature,
                    tier=0,
                    file_format="xlsx",
                    test_cases=[
                        TestCase(
                            id=write_feature,
                            label=f"Data shape: {dtype} bulk write ({tier_label})",
                            row=1,
                            expected={
                                "workload": _data_shape_write_workload(
                                    dtype=dtype,
                                    sheet=sheet,
                                    rng=rng,
                                    scenario=write_feature,
                                ),
                            },
                            importance=Importance.BASIC,
                        )
                    ],
                )
            )

    return files


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
    parser.add_argument(
        "--shape-only",
        action="store_true",
        help=(
            "Skip legacy throughput scenarios; emit only the data-shape "
            "matrix (10 dtypes × 1k/10k/100k tiers, +1M with --include-1m)."
        ),
    )
    parser.add_argument(
        "--include-1m",
        action="store_true",
        help=(
            "Include the 1M-cell tier in the data-shape matrix. Generation "
            "takes ~5 min on a bench machine."
        ),
    )
    args = parser.parse_args()

    out = Path(args.output)
    tier_dir = out / "tier0"
    tier_dir.mkdir(parents=True, exist_ok=True)

    files: list[TestFile] = []

    if args.shape_only:
        files.extend(
            generate_data_shape_scenarios(tier_dir, include_1m=args.include_1m)
        )
        manifest = Manifest(
            generated_at=datetime.now(UTC),
            excel_version="xlsxwriter-generated",
            generator_version="throughput-0.2.0-shape",
            file_format="xlsx",
            files=files,
        )
        write_manifest(manifest, out / "manifest.json")
        print(f"✓ Wrote {len(files)} data-shape fixture(s) to {out}")
        print(f"  Manifest: {out / 'manifest.json'}")
        return

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

    # Bulk read raw variant (same file, bypasses CellValue wrapping)
    files.append(
        TestFile(
            path=f"tier0/{filename}",
            feature="cell_values_10k_bulk_read_raw",
            tier=0,
            file_format="xlsx",
            test_cases=[
                TestCase(
                    id="cell_values_10k_bulk_read_raw",
                    label="Throughput: cell values bulk read raw (10k cells)",
                    row=1,
                    expected={
                        "workload": {
                            "scenario": "cell_values_10k_bulk_read_raw",
                            "op": "bulk_sheet_values_raw",
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

    # Bulk read raw variant (same file, bypasses CellValue wrapping)
    files.append(
        TestFile(
            path=f"tier0/{filename}",
            feature="cell_values_1k_bulk_read_raw",
            tier=0,
            file_format="xlsx",
            test_cases=[
                TestCase(
                    id="cell_values_1k_bulk_read_raw",
                    label="Throughput: cell values bulk read raw (1k cells)",
                    row=1,
                    expected={
                        "workload": {
                            "scenario": "cell_values_1k_bulk_read_raw",
                            "op": "bulk_sheet_values_raw",
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

    # -- Bulk styled write scenarios (batch values + batch formats/borders) --

    # 1k bg_color bulk write
    scenario = "background_colors_1k_bulk_write"
    sheet = "S1"
    rows, cols = 40, 25
    end_cell = _coord_to_cell(rows, cols)
    rng = f"A1:{end_cell}"
    palette = ["#FF0000", "#00FF00", "#0000FF", "#FFFF00"]
    files.append(
        TestFile(
            path="tier0/00_background_colors_1k.xlsx",
            feature=scenario,
            tier=0,
            file_format="xlsx",
            test_cases=[
                TestCase(
                    id=scenario,
                    label="Throughput: bg_color bulk write (1k cells)",
                    row=1,
                    expected={
                        "workload": {
                            "scenario": scenario,
                            "op": "bulk_write_styled_grid",
                            "operations": ["write"],
                            "sheet": sheet,
                            "range": rng,
                            "style_kind": "format",
                            "palette": palette,
                            "start": 1,
                            "step": 1,
                        }
                    },
                    importance=Importance.BASIC,
                )
            ],
        )
    )

    # 1k number_format bulk write (uses format batch API)
    scenario = "number_formats_1k_bulk_write"
    sheet = "S1"
    rows, cols = 40, 25
    end_cell = _coord_to_cell(rows, cols)
    rng = f"A1:{end_cell}"
    files.append(
        TestFile(
            path="tier0/00_number_formats_1k.xlsx",
            feature=scenario,
            tier=0,
            file_format="xlsx",
            test_cases=[
                TestCase(
                    id=scenario,
                    label="Throughput: number_format bulk write (1k cells)",
                    row=1,
                    expected={
                        "workload": {
                            "scenario": scenario,
                            "op": "bulk_write_styled_grid",
                            "operations": ["write"],
                            "sheet": sheet,
                            "range": rng,
                            "style_kind": "format",
                            "palette": ["#FFFFFF"],
                            "start": 1,
                            "step": 1,
                        }
                    },
                    importance=Importance.BASIC,
                )
            ],
        )
    )

    # 200 borders bulk write
    scenario = "borders_200_bulk_write"
    sheet = "S1"
    rows, cols = 20, 10
    end_cell = _coord_to_cell(rows, cols)
    rng = f"A1:{end_cell}"
    files.append(
        TestFile(
            path="tier0/00_borders_200.xlsx",
            feature=scenario,
            tier=0,
            file_format="xlsx",
            test_cases=[
                TestCase(
                    id=scenario,
                    label="Throughput: borders bulk write (200 cells)",
                    row=1,
                    expected={
                        "workload": {
                            "scenario": scenario,
                            "op": "bulk_write_styled_grid",
                            "operations": ["write"],
                            "sheet": sheet,
                            "range": rng,
                            "style_kind": "border",
                            "border_style": "thin",
                            "border_color": "#000000",
                            "start": 1,
                            "step": 1,
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
                feature="cell_values_100k_bulk_read_raw",
                tier=0,
                file_format="xlsx",
                test_cases=[
                    TestCase(
                        id="cell_values_100k_bulk_read_raw",
                        label="Throughput: cell values bulk read raw (~100k cells)",
                        row=1,
                        expected={
                            "workload": {
                                "scenario": "cell_values_100k_bulk_read_raw",
                                "op": "bulk_sheet_values_raw",
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

    # Append data-shape scenarios alongside legacy in default mode.
    # 1M tier still gated behind --include-1m to keep default generation under
    # 30s per the plan budget.
    files.extend(
        generate_data_shape_scenarios(tier_dir, include_1m=args.include_1m)
    )

    manifest = Manifest(
        generated_at=datetime.now(UTC),
        excel_version="xlsxwriter-generated",
        generator_version="throughput-0.2.0",
        file_format="xlsx",
        files=files,
    )
    write_manifest(manifest, out / "manifest.json")

    print(f"✓ Wrote {len(files)} throughput fixture(s) to {out}")
    print(f"  Manifest: {out / 'manifest.json'}")


if __name__ == "__main__":
    main()
