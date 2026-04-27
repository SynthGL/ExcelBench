"""Internal subprocess entrypoint for ``time -l`` memory measurement.

Invoked by :func:`excelbench.perf.runner._measure_iteration_under_time_l`. Runs
exactly one iteration of (adapter, kind, feature) and prints metrics JSON to
stdout; the parent harness wraps the invocation in ``/usr/bin/time -l`` and
parses peak RSS from stderr.

Not part of the public CLI surface — only the perf runner calls this.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="excelbench-perf-iter")
    parser.add_argument("--library", required=True)
    parser.add_argument("--kind", choices=["read", "write"], required=True)
    parser.add_argument("--manifest", required=True, type=Path)
    parser.add_argument("--feature", required=True)
    parser.add_argument(
        "--memory-mode",
        default="getrusage",
        choices=["getrusage", "tracemalloc", "all"],
        help="In-process memory mode for the iteration. The 'time' mode is "
        "implicit — that's why the parent process is calling us under "
        "/usr/bin/time -l.",
    )
    args = parser.parse_args(argv)

    from excelbench.generator.generate import load_manifest
    from excelbench.harness.adapters import get_all_adapters
    from excelbench.perf.runner import run_one_iteration

    manifest = load_manifest(args.manifest)
    matching = [f for f in manifest.files if f.feature == args.feature]
    if not matching:
        sys.stderr.write(f"feature {args.feature!r} not in manifest\n")
        return 2
    test_file = matching[0]

    adapters_by_name = {a.name: a for a in get_all_adapters()}
    adapter = adapters_by_name.get(args.library)
    if adapter is None:
        sys.stderr.write(f"adapter {args.library!r} not registered\n")
        return 2

    test_dir = args.manifest.parent
    metrics = run_one_iteration(
        adapter=adapter,
        kind=args.kind,
        test_file=test_file,
        test_dir=test_dir,
        memory_mode=args.memory_mode,
    )
    serializable = {
        k: v
        for k, v in metrics.items()
        if k in ("wall_ms", "cpu_ms", "rss_peak_mb", "python_heap_peak_kb")
        and isinstance(v, int | float)
    }
    sys.stdout.write(json.dumps(serializable) + "\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
