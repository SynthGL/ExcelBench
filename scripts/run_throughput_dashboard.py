#!/usr/bin/env python3
"""Run a standard throughput/performance dashboard.

This script generates the throughput fixtures into a gitignored folder and runs
several perf batches with consistent adapter sets.

Outputs are written under the provided output root (default: results_dev_perf_dashboard/).
"""

from __future__ import annotations

import argparse
import json
import shlex
import subprocess
from pathlib import Path
from typing import TypedDict


class _Job(TypedDict):
    name: str
    adapters: list[str]
    features: list[str]


def _run(cmd: list[str]) -> None:
    print("+ " + shlex.join(cmd), flush=True)
    subprocess.run(cmd, check=True)


def main() -> None:
    parser = argparse.ArgumentParser(description="Run ExcelBench throughput dashboard")
    parser.add_argument(
        "--tests",
        type=Path,
        default=Path("test_files/throughput_xlsx"),
        help="Throughput fixtures directory (default: test_files/throughput_xlsx)",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("results_dev_perf_dashboard"),
        help="Output root (default: results_dev_perf_dashboard)",
    )
    parser.add_argument("--warmup", type=int, default=0)
    parser.add_argument("--iters", type=int, default=1)
    parser.add_argument(
        "--breakdown",
        action="store_true",
        help="Enable phase breakdown timings (slower).",
    )
    parser.add_argument(
        "--include-slow",
        action="store_true",
        help="Include python-calamine per-cell scenarios (1k only; bulk reads run by default).",
    )
    parser.add_argument(
        "--include-100k",
        action="store_true",
        help="Include ~100k cell fixture generation and run 100k bulk batches.",
    )
    args = parser.parse_args()

    tests_dir = Path(args.tests)
    out_root = Path(args.output)

    # 1) Generate fixtures.
    gen_cmd = [
        "uv",
        "run",
        "python",
        "scripts/generate_throughput_fixtures.py",
        "--output",
        str(tests_dir),
    ]
    if args.include_100k:
        gen_cmd.append("--include-100k")
    _run(gen_cmd)

    # Validate manifest paths and obvious mixups.
    manifest_path = tests_dir / "manifest.json"
    if not manifest_path.exists():
        raise FileNotFoundError(f"Missing manifest: {manifest_path}")
    manifest = json.loads(manifest_path.read_text())
    for f in manifest.get("files", []):
        feat = str(f.get("feature"))
        rel = str(f.get("path"))
        fp = tests_dir / rel
        if not fp.exists():
            raise FileNotFoundError(f"Missing throughput fixture: {fp}")
        if feat.startswith("formulas") and "cell_values" in rel:
            raise ValueError(f"Bad mapping: {feat} points at {rel}")
        if feat.startswith("cell_values") and "formulas" in rel:
            raise ValueError(f"Bad mapping: {feat} points at {rel}")

    # 2) Dashboard batches.
    jobs: list[_Job] = []

    jobs.append(
        {
            "name": "bulk_read_multi",
            "adapters": [
                "openpyxl",
                "openpyxl-readonly",
                "pandas",
                "polars",
                "python-calamine",
                "tablib",
                "wolfxl",
            ],
            "features": [
                "cell_values_1k_bulk_read",
                "cell_values_1k_bulk_read_raw",
                "cell_values_10k_bulk_read",
                "cell_values_10k_bulk_read_raw",
                "cell_values_10k_1000x10_bulk_read",
                "cell_values_10k_10x1000_bulk_read",
                "formulas_1k_bulk_read",
                "formulas_10k_bulk_read",
                "strings_unique_1k_bulk_read",
                "strings_unique_10k_bulk_read",
                "strings_repeated_10k_bulk_read",
                "strings_unique_1k_len64_bulk_read",
                "strings_unique_1k_len256_bulk_read",
                "strings_repeated_1k_len256_bulk_read",
            ],
        }
    )

    jobs.append(
        {
            "name": "bulk_write_multi",
            "adapters": [
                "xlsxwriter",
                "openpyxl",
                "pandas",
                "tablib",
                "wolfxl",
                "rust_xlsxwriter",
            ],
            "features": [
                "cell_values_1k_bulk_write",
                "cell_values_10k_bulk_write",
                "cell_values_10k_1000x10_bulk_write",
                "cell_values_10k_10x1000_bulk_write",
                "cell_values_10k_sparse_1pct_bulk_write",
                "strings_unique_1k_bulk_write",
                "strings_unique_10k_bulk_write",
                "strings_repeated_10k_bulk_write",
                "strings_unique_1k_len64_bulk_write",
                "strings_unique_1k_len256_bulk_write",
                "strings_repeated_1k_len256_bulk_write",
            ],
        }
    )

    jobs.append(
        {
            "name": "bulk_write_styled",
            "adapters": [
                "wolfxl",
                "openpyxl",
                "xlsxwriter",
            ],
            "features": [
                "background_colors_1k_bulk_write",
                "number_formats_1k_bulk_write",
                "borders_200_bulk_write",
            ],
        }
    )

    jobs.append(
        {
            "name": "per_cell_fast",
            "adapters": ["openpyxl", "xlsxwriter", "pylightxl", "pyexcel", "wolfxl"],
            "features": [
                "cell_values_1k",
                "cell_values_10k",
                "formulas_1k",
                "formulas_10k",
                "background_colors_1k",
                "number_formats_1k",
                "alignment_1k",
                "borders_200",
            ],
        }
    )

    if args.include_slow:
        jobs.append(
            {
                "name": "per_cell_slow",
                "adapters": ["python-calamine"],
                "features": [
                    "cell_values_1k",
                    "formulas_1k",
                ],
            }
        )

    if args.include_100k:
        jobs.append(
            {
                "name": "bulk_read_100k",
                "adapters": [
                    "openpyxl",
                    "openpyxl-readonly",
                    "pandas",
                    "polars",
                    "python-calamine",
                    "tablib",
                    "wolfxl",
                ],
                "features": [
                    "cell_values_100k_bulk_read",
                    "cell_values_100k_bulk_read_raw",
                ],
            }
        )
        jobs.append(
            {
                "name": "bulk_write_100k",
                "adapters": [
                    "xlsxwriter",
                    "openpyxl",
                    "pandas",
                    "tablib",
                    "wolfxl",
                    "rust_xlsxwriter",
                ],
                "features": [
                    "cell_values_100k_bulk_write",
                ],
            }
        )

    for job in jobs:
        name = job["name"]
        adapters = job["adapters"]
        features = job["features"]

        job_out = out_root / name
        job_out.mkdir(parents=True, exist_ok=True)

        cmd = [
            "uv",
            "run",
            "excelbench",
            "perf",
            "--tests",
            str(tests_dir),
            "--output",
            str(job_out),
            "--warmup",
            str(args.warmup),
            "--iters",
            str(args.iters),
        ]
        if args.breakdown:
            cmd.append("--breakdown")
        for a in adapters:
            cmd += ["--adapter", str(a)]
        for f in features:
            cmd += ["--feature", str(f)]
        _run(cmd)

    print(f"\n✓ Dashboard complete: {out_root}", flush=True)


if __name__ == "__main__":
    main()
