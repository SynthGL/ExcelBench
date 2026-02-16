#!/usr/bin/env python3
"""Memory profiling for Excel adapters at scale.

Runs each adapter in a **separate subprocess** so that ru_maxrss reflects
only that adapter's memory usage.  Also captures tracemalloc peak for
Python-side allocations (does not include Rust/C heap).

Usage:
    uv run python scripts/memory_profile.py [--adapters a1,a2,...] [--scales 1k,10k,100k]
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
import textwrap
from pathlib import Path

FIXTURE_DIR = Path("test_files/throughput_xlsx/tier0")

SCALE_FILES = {
    "1k": "00_cell_values_1k.xlsx",
    "10k": "00_cell_values_10k.xlsx",
    "100k": "00_cell_values_100k.xlsx",
}

# Adapters with read_sheet_values and/or write_sheet_values
READ_ADAPTERS = [
    "wolfxl",
    "calamine-styled",
    "openpyxl",
    "python-calamine",
    "pandas",
]
WRITE_ADAPTERS = [
    "wolfxl",
    "rust_xlsxwriter",
    "openpyxl",
    "xlsxwriter",
    "pandas",
    "tablib",
]

# Worker script run in subprocess — measures one adapter/op/scale combo
WORKER_SCRIPT = textwrap.dedent("""\
import gc
import json
import resource
import sys
import tracemalloc
from pathlib import Path

adapter_name = sys.argv[1]
op = sys.argv[2]  # "read" or "write"
fixture_path = Path(sys.argv[3])

# --- Resolve adapter by name ---
from excelbench.harness.adapters import get_all_adapters

adapter = None
for a in get_all_adapters():
    if a.name == adapter_name:
        adapter = a
        break
if adapter is None:
    print(json.dumps({"error": f"Adapter {adapter_name!r} not found"}))
    sys.exit(1)

# --- Baseline ---
gc.collect()
rss_before = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
tracemalloc.start()

if op == "read":
    if not hasattr(adapter, "read_sheet_values"):
        print(json.dumps({"error": f"{adapter_name} has no read_sheet_values"}))
        sys.exit(1)
    wb = adapter.open_workbook(fixture_path)
    sheets = adapter.get_sheet_names(wb)
    data = adapter.read_sheet_values(wb, sheets[0])
    row_count = len(data)
    col_count = len(data[0]) if data else 0
    adapter.close_workbook(wb)
elif op == "write":
    if not hasattr(adapter, "write_sheet_values"):
        print(json.dumps({"error": f"{adapter_name} has no write_sheet_values"}))
        sys.exit(1)
    # Read the fixture to get the grid values
    from excelbench.harness.adapters.openpyxl_adapter import OpenpyxlAdapter
    ref = OpenpyxlAdapter()
    ref_wb = ref.open_workbook(fixture_path)
    ref_sheets = ref.get_sheet_names(ref_wb)
    grid = ref.read_sheet_values(ref_wb, ref_sheets[0])
    ref.close_workbook(ref_wb)
    # Convert CellValue objects to raw Python values
    raw_grid = []
    for row in grid:
        raw_row = []
        for cell in row:
            raw_row.append(cell.value if hasattr(cell, "value") else cell)
        raw_grid.append(raw_row)
    row_count = len(raw_grid)
    col_count = len(raw_grid[0]) if raw_grid else 0

    # Reset memory tracking after loading ref data
    tracemalloc.stop()
    gc.collect()
    rss_before = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    tracemalloc.start()

    import tempfile
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=True) as tmp:
        out_path = Path(tmp.name)
    wb = adapter.create_workbook()
    adapter.add_sheet(wb, "Sheet1")
    adapter.write_sheet_values(wb, "Sheet1", "A1", raw_grid)
    adapter.save_workbook(wb, out_path)
    out_path.unlink(missing_ok=True)
else:
    print(json.dumps({"error": f"Unknown op {op!r}"}))
    sys.exit(1)

# --- Measure ---
gc.collect()
tm_current, tm_peak = tracemalloc.get_traced_memory()
tracemalloc.stop()

rss_after = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss

# macOS reports ru_maxrss in bytes; Linux in KB
import platform
if platform.system() == "Darwin":
    rss_before_mb = rss_before / (1024 * 1024)
    rss_after_mb = rss_after / (1024 * 1024)
else:
    rss_before_mb = rss_before / 1024
    rss_after_mb = rss_after / 1024

print(json.dumps({
    "adapter": adapter_name,
    "op": op,
    "cells": row_count * col_count,
    "rss_before_mb": round(rss_before_mb, 2),
    "rss_after_mb": round(rss_after_mb, 2),
    "rss_delta_mb": round(rss_after_mb - rss_before_mb, 2),
    "tracemalloc_peak_mb": round(tm_peak / (1024 * 1024), 2),
}))
""")


def run_one(adapter: str, op: str, fixture: Path) -> dict | None:
    """Run a single adapter/op measurement in a subprocess."""
    try:
        result = subprocess.run(
            [sys.executable, "-c", WORKER_SCRIPT, adapter, op, str(fixture)],
            capture_output=True,
            text=True,
            timeout=120,
        )
    except subprocess.TimeoutExpired:
        return {"adapter": adapter, "op": op, "error": "timeout"}

    if result.returncode != 0:
        stderr = result.stderr.strip()
        # Try to extract JSON from stdout even on failure
        try:
            return json.loads(result.stdout.strip())
        except (json.JSONDecodeError, ValueError):
            return {"adapter": adapter, "op": op, "error": stderr[-300:] if stderr else "unknown"}

    try:
        return json.loads(result.stdout.strip())
    except (json.JSONDecodeError, ValueError):
        return {"adapter": adapter, "op": op, "error": f"bad json: {result.stdout[:200]}"}


def main() -> None:
    parser = argparse.ArgumentParser(description="Memory profiling for Excel adapters")
    parser.add_argument(
        "--adapters",
        help="Comma-separated adapter names (default: all available)",
    )
    parser.add_argument(
        "--scales",
        default="1k,10k,100k",
        help="Comma-separated scales (default: 1k,10k,100k)",
    )
    parser.add_argument(
        "--output",
        help="Output JSON file path",
    )
    args = parser.parse_args()

    scales = [s.strip() for s in args.scales.split(",")]
    read_adapters = args.adapters.split(",") if args.adapters else READ_ADAPTERS
    write_adapters = args.adapters.split(",") if args.adapters else WRITE_ADAPTERS

    results: list[dict] = []

    for scale in scales:
        fixture_name = SCALE_FILES.get(scale)
        if not fixture_name:
            print(f"  [skip] Unknown scale: {scale}")
            continue
        fixture_path = FIXTURE_DIR / fixture_name
        if not fixture_path.exists():
            print(f"  [skip] Fixture not found: {fixture_path}")
            continue

        print(f"\n{'=' * 60}")
        print(f"  Scale: {scale} ({fixture_path.name})")
        print(f"{'=' * 60}")

        # Read benchmarks
        print("\n  --- Bulk Read ---")
        for adapter in read_adapters:
            print(f"  {adapter:25s} ... ", end="", flush=True)
            r = run_one(adapter, "read", fixture_path)
            if r and "error" not in r:
                print(
                    f"RSS delta: {r['rss_delta_mb']:+8.2f} MB | "
                    f"tracemalloc peak: {r['tracemalloc_peak_mb']:8.2f} MB | "
                    f"RSS total: {r['rss_after_mb']:8.2f} MB"
                )
                r["scale"] = scale
                results.append(r)
            elif r:
                print(f"ERROR: {r.get('error', 'unknown')[:80]}")
            else:
                print("ERROR: no result")

        # Write benchmarks
        print("\n  --- Bulk Write ---")
        for adapter in write_adapters:
            print(f"  {adapter:25s} ... ", end="", flush=True)
            r = run_one(adapter, "write", fixture_path)
            if r and "error" not in r:
                print(
                    f"RSS delta: {r['rss_delta_mb']:+8.2f} MB | "
                    f"tracemalloc peak: {r['tracemalloc_peak_mb']:8.2f} MB | "
                    f"RSS total: {r['rss_after_mb']:8.2f} MB"
                )
                r["scale"] = scale
                results.append(r)
            elif r:
                print(f"ERROR: {r.get('error', 'unknown')[:80]}")
            else:
                print("ERROR: no result")

    # Summary table
    if results:
        print(f"\n\n{'=' * 80}")
        print("  SUMMARY: Memory Usage by Adapter")
        print(f"{'=' * 80}")
        print(
            f"  {'Adapter':<25s} {'Op':<7s} {'Scale':<6s} {'Cells':>8s} "
            f"{'RSS Delta':>10s} {'TM Peak':>10s} {'RSS Total':>10s}"
        )
        print(f"  {'-' * 25} {'-' * 6} {'-' * 5} {'-' * 8} {'-' * 10} {'-' * 10} {'-' * 10}")
        for r in results:
            print(
                f"  {r['adapter']:<25s} {r['op']:<7s} {r['scale']:<6s} "
                f"{r['cells']:>8d} {r['rss_delta_mb']:>+9.2f}M "
                f"{r['tracemalloc_peak_mb']:>9.2f}M {r['rss_after_mb']:>9.2f}M"
            )

    if args.output:
        out_path = Path(args.output)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        with open(out_path, "w") as f:
            json.dump(results, f, indent=2)
        print(f"\n  Results written to {out_path}")


if __name__ == "__main__":
    main()
