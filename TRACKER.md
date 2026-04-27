# ExcelBench — Sprint Tracker

> Single source of truth for the **7-Dimension Extension** initiative. Each row tracks one
> self-contained sprint (one branch, one PR, one row flip). Resume cold by reading this file
> and the most recent `[*INCOMPLETE*]` marker.

**Last updated**: 2026-04-27 (S2 in progress)

## Status Table

| #  | Dimension                          | Status      | Sprint size | Branch                         | PR  | Acceptance commit range |
|----|------------------------------------|-------------|-------------|--------------------------------|-----|-------------------------|
| S1 | Memory honesty + Tracker bootstrap | Shipped     | S (3–5 d)   | `feat/perf-mem-honesty`        | #28 | `50dc104..HEAD@PR#28`   |
| S2 | Data shape (int/str/date/formula)  | In Progress | M (1 wk)    | `feat/perf-data-shape`         | —   | —                       |
| S3 | File shape (wide/tall/sparse)      | Planned     | M (1 wk)    | `feat/perf-file-shape`         | —   | —                       |
| S4 | High-cost operations               | Planned     | M (1 wk)    | `feat/perf-operations`         | —   | —                       |
| S5 | Workbook complexity perf           | Planned     | M (1 wk)    | `feat/perf-complexity`         | —   | —                       |
| S6 | Cold-start / warm path             | Planned     | S (3–5 d)   | `feat/perf-cold-start`         | —   | —                       |
| S7 | Round-trip fidelity (LibreOffice)  | Planned     | L (~2 wk)   | `feat/fidelity-roundtrip`      | —   | —                       |

**Status legend**: `Planned` → `In Progress` → `Shipped` (or `Blocked` with reason).

## How to Flip a Row

When a sprint lands:

1. Update the row's **Status** to `Shipped`.
2. Fill in the **PR** column (`#NN`).
3. Fill in **Acceptance commit range** (`abc1234..def5678`).
4. Bump the **Last updated** line at the top of this file.
5. Append a sprint acceptance entry (template below) to the **Acceptance Notes** section.
6. Add the corresponding `DEC-NNN` entry to `decisions.md` if not already done.

If a sprint stalls, switch its status to `Blocked` and add a one-line reason in the row.

## Sprint Acceptance Template

Use this template when appending to **Acceptance Notes** below.

```markdown
### S<N> — <Dimension> (YYYY-MM-DD)

**Branch**: `feat/...`  ·  **PR**: #NN  ·  **Commit range**: `abc1234..def5678`

**What shipped**:
- <one-line bullet per major piece>

**Verification**:
- `uv run pytest tests/` ✓
- `uv run ruff check src/ tests/` ✓
- `uv run mypy src/` ✓
- `excelbench <new-subcommand> ...` ✓ (16 adapters, no crashes)
- Dashboard regenerated, results.json + history.jsonl appended.

**Decisions**: DEC-NNN logged in `decisions.md`.

**Deferred / out-of-scope**:
- <items intentionally left for follow-up>
```

## Acceptance Notes

<!-- Newest first. Append entries here as sprints ship. -->

### S1 — Memory honesty + Tracker bootstrap (2026-04-27)

**Branch**: `feat/perf-mem-honesty`  ·  **PR**: [#28](https://github.com/SynthGL/ExcelBench/pull/28)  ·  **Commit range**: `50dc104..HEAD` (final range fills in on merge)

**What shipped**:
- `TRACKER.md` (this file) — 7-row sprint table, row-flip protocol, acceptance template.
- `src/excelbench/perf/memory.py` — three-mode memory harness (`getrusage` / `tracemalloc` /
  `time` via `/usr/bin/time -l` subprocess + `all` composite). `MemoryProbe` context manager
  for in-process modes; `parse_time_l_stderr` cross-platform parser (macOS BSD time + GNU
  time `-l`).
- `PerfOpResult` extended with `rss_kb_via_time` and `python_heap_peak_kb` fields (existing
  `rss_peak_mb` preserved — backwards-compatible).
- `src/excelbench/perf/_iter_subprocess.py` — internal subprocess entrypoint that runs one
  iteration per invocation; wrapped by parent under `/usr/bin/time -l`.
- `excelbench perf --memory-mode={getrusage,tracemalloc,time,all}` CLI flag.
- HTML dashboard renders dual `RSS (MB) — getrusage / time -l` cells with a tooltip
  explaining divergence whenever any entry has a `time -l` measurement.
- DEC-018 documents why three modes coexist and what each is honest about.

**Verification** (run on macOS 25.2, Python 3.13):
- `uv run pytest tests/` ✓ 1140 passed, 32 skipped, 6 xfailed
- `uv run ruff check src/excelbench/perf/ src/excelbench/cli.py src/excelbench/results/html_dashboard.py` ✓
- `uv run mypy src/excelbench/perf/` ✓ no issues
- `excelbench perf --memory-mode=all --feature cell_values --adapter wolfxl --adapter openpyxl --warmup 1 --iters 2`:
  - All three fields populated as expected.
  - Python-heap honesty signal landed: openpyxl uses 16× (read) and 227× (write) more
    Python heap than wolfxl on the same workload, confirming wolfxl pushes allocations into Rust.
  - `time -l/getrusage` ratio ~0.97x on small fixtures (subprocess startup dominates);
    expected to diverge meaningfully once Sprint 2 lands ≥1M-cell fixtures.

**Decisions**: DEC-018 logged in `decisions.md`.

**Deferred / out-of-scope**:
- Tracemalloc reset semantics across nested probes — current code uses `reset_peak()` when a
  probe re-enters an already-traced context. Should be revisited if any caller starts
  tracemalloc outside the probe.
- `time -l` subprocess support on Windows — skipped silently (no `/usr/bin/time`).
  Sprint 6 (cold-start) will set the precedent for cross-platform subprocess handling.
- Visualizing the time-l/getrusage divergence as a dedicated chart — single dual-cell
  with tooltip is sufficient until S2 ships larger fixtures that make the gap visible.

## Reference

- Plan: see the wolfxl session that produced this tracker (multi-sprint roadmap).
- Architecture: [`architecture.md`](architecture.md)
- Decisions: [`decisions.md`](decisions.md)
- Key seams: `src/excelbench/perf/runner.py`, `src/excelbench/harness/adapters/base.py`,
  `src/excelbench/results/html_dashboard.py`.
