"""Memory measurement harness with three coexisting modes.

The perf runner can measure memory three different ways. None of them is *the*
right answer in isolation — each is honest about a different slice of reality:

- ``getrusage`` (default, cheap): peak RSS via ``resource.getrusage(RUSAGE_SELF)``.
  Returns the **process-lifetime peak**, so once the first heavy iteration has
  allocated, subsequent iterations report the sticky max regardless of whether
  they actually allocated less. Iteration-noisy in practice.

- ``tracemalloc``: peak Python heap via ``tracemalloc.get_traced_memory()``.
  Misses Rust / PyO3 / native-extension allocations entirely. Misleading for
  Rust-backed adapters (wolfxl, python-calamine, rust_xlsxwriter); fine for
  pure-Python adapters (openpyxl, xlsxwriter).

- ``time``: spawn ``/usr/bin/time -l <cmd>`` in a fresh subprocess and parse
  ``maximum resident set size`` from stderr. Honest about Rust allocations
  because the measurement is from the OS. Slow because each iteration pays
  Python startup + adapter import cost; that overhead is included in the
  reported peak (a feature, not a bug — it answers "what does running this
  library actually cost in a fresh process").

The composite ``all`` mode runs all three sequentially per iteration so the
quarterly memory deep-dive can compare them directly.
"""

from __future__ import annotations

import json
import re
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Literal

MemoryMode = Literal["getrusage", "time", "tracemalloc", "all"]
VALID_MEMORY_MODES: tuple[MemoryMode, ...] = ("getrusage", "time", "tracemalloc", "all")


@dataclass(frozen=True)
class MemorySample:
    """One memory measurement of an iteration. Any field may be ``None``.

    All RSS fields are reported in **kilobytes** for direct comparability with
    ``/usr/bin/time -l`` stderr output. The runner converts to MB for display.
    """

    rss_via_getrusage_kb: float | None = None
    rss_via_time_kb: float | None = None
    python_heap_peak_kb: float | None = None


def includes_getrusage(mode: MemoryMode) -> bool:
    return mode in ("getrusage", "all")


def includes_tracemalloc(mode: MemoryMode) -> bool:
    return mode in ("tracemalloc", "all")


def includes_time(mode: MemoryMode) -> bool:
    return mode in ("time", "all")


class MemoryProbe:
    """Context manager that captures the in-process memory modes.

    Use ``with MemoryProbe(mode) as probe: ...`` and read ``probe.sample`` after
    exit. The ``time`` mode is **not** captured here — it requires a subprocess
    and is invoked separately by the runner via :func:`run_iteration_under_time_l`.
    """

    def __init__(self, mode: MemoryMode) -> None:
        self.mode: MemoryMode = mode
        self.sample: MemorySample = MemorySample()
        self._tm_started = False

    def __enter__(self) -> MemoryProbe:
        if includes_tracemalloc(self.mode):
            import tracemalloc

            # Avoid double-start in case a parent context already started it.
            if not tracemalloc.is_tracing():
                tracemalloc.start()
                self._tm_started = True
            else:
                # Reset peak so we measure only this scope.
                tracemalloc.reset_peak()
        return self

    def __exit__(self, *_exc_info: object) -> None:
        rss_via_getrusage_kb: float | None = None
        python_heap_peak_kb: float | None = None

        if includes_getrusage(self.mode):
            import resource

            ru = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
            rss_via_getrusage_kb = _ru_maxrss_to_kb(ru)

        if includes_tracemalloc(self.mode):
            import tracemalloc

            _, peak = tracemalloc.get_traced_memory()
            python_heap_peak_kb = float(peak) / 1024.0
            if self._tm_started:
                tracemalloc.stop()

        self.sample = MemorySample(
            rss_via_getrusage_kb=rss_via_getrusage_kb,
            rss_via_time_kb=None,  # always populated separately by subprocess path
            python_heap_peak_kb=python_heap_peak_kb,
        )


def run_iteration_under_time_l(
    cli_args: list[str],
    *,
    cwd: Path | None = None,
    timeout_s: float = 300.0,
) -> tuple[dict[str, float] | None, float | None]:
    """Run ``/usr/bin/time -l <cli_args>`` and return (subprocess stdout JSON, peak RSS KB).

    The subprocess is expected to print one JSON object to stdout containing
    iteration metrics (wall_ms, cpu_ms). Stderr captures ``time -l`` output, from
    which we extract ``maximum resident set size``.

    Returns ``(metrics, rss_kb)``. Either may be ``None`` if parsing failed —
    callers should treat ``None`` as "measurement unavailable" (e.g., on
    Windows, where ``/usr/bin/time`` doesn't exist).
    """
    time_l = _resolve_time_l_path()
    if time_l is None:
        return None, None

    cmd = [time_l, "-l", *cli_args]
    try:
        completed = subprocess.run(
            cmd,
            cwd=str(cwd) if cwd else None,
            capture_output=True,
            text=True,
            timeout=timeout_s,
            check=False,
        )
    except subprocess.TimeoutExpired:
        return None, None
    except (FileNotFoundError, OSError):
        return None, None

    metrics: dict[str, float] | None = None
    if completed.returncode == 0 and completed.stdout.strip():
        try:
            parsed = json.loads(completed.stdout.strip().splitlines()[-1])
            if isinstance(parsed, dict):
                metrics = {k: float(v) for k, v in parsed.items() if isinstance(v, int | float)}
        except (json.JSONDecodeError, ValueError):
            metrics = None

    rss_kb = parse_time_l_stderr(completed.stderr)
    return metrics, rss_kb


def parse_time_l_stderr(text: str) -> float | None:
    """Extract peak RSS in KB from ``/usr/bin/time -l`` stderr output.

    Cross-platform: macOS reports bytes ("maximum resident set size") while
    Linux's GNU time -l reports kilobytes ("Maximum resident set size (kbytes)").
    """
    if not text:
        return None

    # macOS: "  12345678  maximum resident set size"   (value in bytes)
    macos_pattern = re.compile(r"^\s*(\d+)\s+maximum resident set size", re.MULTILINE)
    m = macos_pattern.search(text)
    if m:
        return float(m.group(1)) / 1024.0  # bytes → KB

    # GNU coreutils time -l (rare): "Maximum resident set size (kbytes): 12345"
    gnu_kb_pattern = re.compile(
        r"^\s*Maximum resident set size \(kbytes\):\s*(\d+)", re.MULTILINE
    )
    m = gnu_kb_pattern.search(text)
    if m:
        return float(m.group(1))  # already KB

    return None


def _resolve_time_l_path() -> str | None:
    """Return the absolute path to ``/usr/bin/time`` if available, else ``None``.

    On macOS this is the BSD time supporting ``-l``. On Linux it is GNU time
    (``-l`` is also accepted). On Windows there is no equivalent and we
    return ``None`` so callers can fall back gracefully.
    """
    if sys.platform == "win32":
        return None
    candidate = "/usr/bin/time"
    if Path(candidate).exists():
        return candidate
    return None


def _ru_maxrss_to_kb(ru_maxrss: float) -> float:
    """Convert ``ru_maxrss`` to kilobytes, accounting for platform differences.

    macOS reports bytes, Linux reports kilobytes (per ``getrusage(2)``).
    """
    if sys.platform == "darwin":
        return float(ru_maxrss) / 1024.0
    return float(ru_maxrss)
