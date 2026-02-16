#!/usr/bin/env python3
"""Large-scale speed benchmark for Excel adapters.

Generates fixtures at 1M and 10M cell scales, then measures bulk read/write
throughput for key adapters. Uses subprocess isolation for accurate memory
measurement.

Usage:
    uv run python scripts/large_scale_benchmark.py [--scales 1m,5m] [--iters 3]
"""

from __future__ import annotations

import argparse
import gc
import json
import os
import resource
import subprocess
import sys
import tempfile
import time
from pathlib import Path

# Scale definitions: name -> (rows, cols, approx_cells)
SCALES = {
    "100k": (316, 316),      # 99,856 cells (~400 KB)
    "1m": (1000, 1000),      # 1,000,000 cells (~5 MB)
    "5m": (2236, 2236),      # 4,999,696 cells (~25 MB)
    "10m": (3162, 3162),      # 9,998,244 cells (~50 MB)
}

READ_ADAPTERS = ["wolfxl", "calamine-styled", "openpyxl", "python-calamine"]
WRITE_ADAPTERS = ["wolfxl", "rust_xlsxwriter", "openpyxl", "xlsxwriter"]

FIXTURE_DIR = Path("test_files/throughput_xlsx/large_scale")


def generate_fixture(rows: int, cols: int, path: Path) -> float:
    """Generate a numeric grid fixture with xlsxwriter. Returns generation time."""
    path.parent.mkdir(parents=True, exist_ok=True)
    import xlsxwriter

    total_cells = rows * cols
    print(f"  Generating {total_cells:,} cells ({rows}x{cols}) ...", end=" ", flush=True)
    t0 = time.perf_counter()
    wb = xlsxwriter.Workbook(str(path), {"constant_memory": True})
    ws = wb.add_worksheet("S1")
    value = 1
    for r in range(rows):
        for c in range(cols):
            ws.write_number(r, c, value)
            value += 1
    wb.close()
    elapsed = time.perf_counter() - t0
    size_mb = path.stat().st_size / (1024 * 1024)
    print(f"done in {elapsed:.1f}s ({size_mb:.1f} MB)")
    return elapsed


def run_read_benchmark(adapter: str, fixture_path: str, iters: int) -> dict | None:
    """Run a bulk-read benchmark in a subprocess."""
    script = f"""
import gc, json, resource, sys, time, platform
from pathlib import Path

adapter_name = "{adapter}"
fixture_path = Path("{fixture_path}")
iters = {iters}

from excelbench.harness.adapters import get_all_adapters
adapter = None
for a in get_all_adapters():
    if a.name == adapter_name:
        adapter = a
        break
if adapter is None:
    print(json.dumps({{"error": f"Adapter {{adapter_name!r}} not found"}}))
    sys.exit(1)

if not hasattr(adapter, "read_sheet_values"):
    print(json.dumps({{"error": f"{{adapter_name}} has no read_sheet_values"}}))
    sys.exit(1)

# Warmup
wb = adapter.open_workbook(fixture_path)
sheets = adapter.get_sheet_names(wb)
data = adapter.read_sheet_values(wb, sheets[0])
adapter.close_workbook(wb)
row_count = len(data)
col_count = len(data[0]) if data else 0
del data
gc.collect()

# Timed iterations
times = []
for i in range(iters):
    gc.collect()
    t0 = time.perf_counter()
    wb = adapter.open_workbook(fixture_path)
    sheets = adapter.get_sheet_names(wb)
    data = adapter.read_sheet_values(wb, sheets[0])
    adapter.close_workbook(wb)
    elapsed = time.perf_counter() - t0
    times.append(elapsed)
    del data

gc.collect()
rss = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
if platform.system() == "Darwin":
    rss_mb = rss / (1024 * 1024)
else:
    rss_mb = rss / 1024

times.sort()
print(json.dumps({{
    "adapter": adapter_name,
    "op": "read",
    "cells": row_count * col_count,
    "times": times,
    "min_s": round(times[0], 4),
    "median_s": round(times[len(times) // 2], 4),
    "max_s": round(times[-1], 4),
    "cells_per_sec": round((row_count * col_count) / times[len(times) // 2]),
    "rss_peak_mb": round(rss_mb, 1),
}}))
"""
    try:
        result = subprocess.run(
            [sys.executable, "-c", script],
            capture_output=True, text=True, timeout=600,
        )
    except subprocess.TimeoutExpired:
        return {"adapter": adapter, "op": "read", "error": "timeout (10min)"}

    if result.returncode != 0:
        stderr = result.stderr.strip()
        try:
            return json.loads(result.stdout.strip())
        except (json.JSONDecodeError, ValueError):
            return {"adapter": adapter, "op": "read", "error": stderr[-300:] if stderr else "unknown"}

    try:
        return json.loads(result.stdout.strip())
    except (json.JSONDecodeError, ValueError):
        return {"adapter": adapter, "op": "read", "error": f"bad json: {result.stdout[:200]}"}


def run_write_benchmark(adapter: str, rows: int, cols: int, iters: int) -> dict | None:
    """Run a bulk-write benchmark in a subprocess."""
    script = f"""
import gc, json, resource, sys, time, platform, tempfile
from pathlib import Path

adapter_name = "{adapter}"
rows = {rows}
cols = {cols}
iters = {iters}

from excelbench.harness.adapters import get_all_adapters
adapter = None
for a in get_all_adapters():
    if a.name == adapter_name:
        adapter = a
        break
if adapter is None:
    print(json.dumps({{"error": f"Adapter {{adapter_name!r}} not found"}}))
    sys.exit(1)

if not hasattr(adapter, "write_sheet_values"):
    print(json.dumps({{"error": f"{{adapter_name}} has no write_sheet_values"}}))
    sys.exit(1)

# Build grid data
grid = []
value = 1
for r in range(rows):
    row = []
    for c in range(cols):
        row.append(value)
        value += 1
    grid.append(row)

total_cells = rows * cols

# Warmup
with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
    out_path = Path(tmp.name)
wb = adapter.create_workbook()
adapter.add_sheet(wb, "Sheet1")
adapter.write_sheet_values(wb, "Sheet1", "A1", grid)
adapter.save_workbook(wb, out_path)
out_path.unlink(missing_ok=True)
gc.collect()

# Timed iterations
times = []
for i in range(iters):
    gc.collect()
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        out_path = Path(tmp.name)
    t0 = time.perf_counter()
    wb = adapter.create_workbook()
    adapter.add_sheet(wb, "Sheet1")
    adapter.write_sheet_values(wb, "Sheet1", "A1", grid)
    adapter.save_workbook(wb, out_path)
    elapsed = time.perf_counter() - t0
    times.append(elapsed)
    file_size = out_path.stat().st_size
    out_path.unlink(missing_ok=True)

gc.collect()
rss = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
if platform.system() == "Darwin":
    rss_mb = rss / (1024 * 1024)
else:
    rss_mb = rss / 1024

times.sort()
print(json.dumps({{
    "adapter": adapter_name,
    "op": "write",
    "cells": total_cells,
    "times": times,
    "min_s": round(times[0], 4),
    "median_s": round(times[len(times) // 2], 4),
    "max_s": round(times[-1], 4),
    "cells_per_sec": round(total_cells / times[len(times) // 2]),
    "file_size_mb": round(file_size / (1024 * 1024), 1),
    "rss_peak_mb": round(rss_mb, 1),
}}))
"""
    try:
        result = subprocess.run(
            [sys.executable, "-c", script],
            capture_output=True, text=True, timeout=600,
        )
    except subprocess.TimeoutExpired:
        return {"adapter": adapter, "op": "write", "error": "timeout (10min)"}

    if result.returncode != 0:
        stderr = result.stderr.strip()
        try:
            return json.loads(result.stdout.strip())
        except (json.JSONDecodeError, ValueError):
            return {"adapter": adapter, "op": "write", "error": stderr[-300:] if stderr else "unknown"}

    try:
        return json.loads(result.stdout.strip())
    except (json.JSONDecodeError, ValueError):
        return {"adapter": adapter, "op": "write", "error": f"bad json: {result.stdout[:200]}"}


def format_throughput(cells_per_sec: int) -> str:
    if cells_per_sec >= 1_000_000:
        return f"{cells_per_sec / 1_000_000:.2f}M/s"
    elif cells_per_sec >= 1_000:
        return f"{cells_per_sec / 1_000:.0f}K/s"
    else:
        return f"{cells_per_sec}/s"


def main() -> None:
    parser = argparse.ArgumentParser(description="Large-scale Excel benchmark")
    parser.add_argument(
        "--scales",
        default="100k,1m",
        help="Comma-separated scales: 100k,1m,5m,10m (default: 100k,1m)",
    )
    parser.add_argument(
        "--iters",
        type=int,
        default=3,
        help="Number of timed iterations (default: 3)",
    )
    parser.add_argument(
        "--output",
        help="Output JSON file path",
    )
    parser.add_argument(
        "--read-only",
        action="store_true",
        help="Only run read benchmarks",
    )
    parser.add_argument(
        "--write-only",
        action="store_true",
        help="Only run write benchmarks",
    )
    args = parser.parse_args()

    scales = [s.strip() for s in args.scales.split(",")]
    do_read = not args.write_only
    do_write = not args.read_only

    all_results: list[dict] = []

    for scale in scales:
        if scale not in SCALES:
            print(f"Unknown scale: {scale} (available: {', '.join(SCALES.keys())})")
            continue

        rows, cols = SCALES[scale]
        total_cells = rows * cols
        fixture_path = FIXTURE_DIR / f"cell_values_{scale}.xlsx"

        print(f"\n{'=' * 70}")
        print(f"  SCALE: {scale.upper()} ({total_cells:,} cells, {rows}x{cols})")
        print(f"{'=' * 70}")

        # Generate fixture if needed (for read benchmarks)
        if do_read and not fixture_path.exists():
            generate_fixture(rows, cols, fixture_path)
        elif do_read:
            size_mb = fixture_path.stat().st_size / (1024 * 1024)
            print(f"  Using existing fixture: {fixture_path} ({size_mb:.1f} MB)")

        # Read benchmarks
        if do_read:
            print(f"\n  --- Bulk Read ({args.iters} iters, 1 warmup) ---")
            print(f"  {'Adapter':<20s} {'Median':>8s} {'Min':>8s} {'Throughput':>12s} {'RSS':>8s}")
            print(f"  {'-'*20} {'-'*8} {'-'*8} {'-'*12} {'-'*8}")
            for adapter in READ_ADAPTERS:
                print(f"  {adapter:<20s} ", end="", flush=True)
                r = run_read_benchmark(adapter, str(fixture_path), args.iters)
                if r and "error" not in r:
                    r["scale"] = scale
                    all_results.append(r)
                    tp = format_throughput(r["cells_per_sec"])
                    print(
                        f"{r['median_s']:>7.3f}s {r['min_s']:>7.3f}s "
                        f"{tp:>12s} {r['rss_peak_mb']:>7.1f}M"
                    )
                elif r:
                    print(f"ERROR: {r.get('error', 'unknown')[:60]}")
                else:
                    print("ERROR: no result")

        # Write benchmarks
        if do_write:
            print(f"\n  --- Bulk Write ({args.iters} iters, 1 warmup) ---")
            print(f"  {'Adapter':<20s} {'Median':>8s} {'Min':>8s} {'Throughput':>12s} {'File':>8s} {'RSS':>8s}")
            print(f"  {'-'*20} {'-'*8} {'-'*8} {'-'*12} {'-'*8} {'-'*8}")
            for adapter in WRITE_ADAPTERS:
                print(f"  {adapter:<20s} ", end="", flush=True)
                r = run_write_benchmark(adapter, rows, cols, args.iters)
                if r and "error" not in r:
                    r["scale"] = scale
                    all_results.append(r)
                    tp = format_throughput(r["cells_per_sec"])
                    print(
                        f"{r['median_s']:>7.3f}s {r['min_s']:>7.3f}s "
                        f"{tp:>12s} {r['file_size_mb']:>7.1f}M {r['rss_peak_mb']:>7.1f}M"
                    )
                elif r:
                    print(f"ERROR: {r.get('error', 'unknown')[:60]}")
                else:
                    print("ERROR: no result")

    # Summary
    if all_results:
        print(f"\n\n{'=' * 70}")
        print("  SUMMARY: Throughput Comparison")
        print(f"{'=' * 70}")

        # Group by scale and op
        for scale in scales:
            for op in ["read", "write"]:
                group = [r for r in all_results if r.get("scale") == scale and r.get("op") == op]
                if not group:
                    continue
                group.sort(key=lambda x: x.get("median_s", float("inf")))
                fastest = group[0]["median_s"]
                total_cells = group[0]["cells"]

                print(f"\n  {op.upper()} @ {scale.upper()} ({total_cells:,} cells)")
                print(f"  {'Rank':<5s} {'Adapter':<20s} {'Median':>8s} {'Throughput':>12s} {'vs fastest':>10s}")
                print(f"  {'-'*5} {'-'*20} {'-'*8} {'-'*12} {'-'*10}")
                for i, r in enumerate(group, 1):
                    tp = format_throughput(r["cells_per_sec"])
                    ratio = r["median_s"] / fastest if fastest > 0 else 0
                    marker = " <-- fastest" if i == 1 else f" {ratio:.1f}x slower"
                    print(f"  {i:<5d} {r['adapter']:<20s} {r['median_s']:>7.3f}s {tp:>12s} {marker}")

    if args.output:
        out_path = Path(args.output)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        with open(out_path, "w") as f:
            json.dump(all_results, f, indent=2)
        print(f"\n  Results written to {out_path}")


if __name__ == "__main__":
    main()
