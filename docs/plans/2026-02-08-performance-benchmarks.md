# Performance Benchmark Track (Speed + Memory)

Created: 02/08/2026 05:07 AM PST (via pst-timestamp)
Status: Draft

## Context

ExcelBench currently measures feature fidelity (correctness). This document adds a parallel track:
measure speed (and best-effort memory) for each library across the same feature/action surface that
drives the fidelity matrix.

This track does not replace fidelity. It adds a second axis:

- "Can this library do X?" (fidelity)
- "How fast is it at X?" (performance)

## Goals

- Report per-library performance for each feature row in the existing matrix.
- Make performance methodology reproducible and explainable.
- Avoid contaminating timings with oracle verification overhead.

## Non-goals (initially)

- Enforce hard perf thresholds in CI (too noisy without dedicated runners).
- Claim cross-machine comparability; we compare within a machine first.
- Perfect microbenchmarks for every possible workflow.

## Definitions

- Feature: one manifest entry / one fixture file (e.g., `borders`, `number_formats`).
- Operation: `read` or `write`.
- Iteration: one full end-to-end execution for a (library, feature, operation) tuple.
- Hot timing: repeated iterations in one process.
- Cold timing: best-effort fresh-process timing (optional; not implemented in v1).

## What We Measure

For each (library, feature):

- read.wall_ms (min/p50/p95): open workbook -> exercise feature reads -> close workbook
- write.wall_ms (min/p50/p95): create workbook -> exercise feature writes -> save workbook
- read.cpu_ms / write.cpu_ms (min/p50/p95): process CPU time
- read.rss_peak_mb / write.rss_peak_mb: best-effort peak RSS via `resource.getrusage`

Optional (recommended) breakdown timings:

- Read phases: open, sheets, exercise, close
- Write phases: create, add_sheets, exercise, save

## Methodology Choices

### 1) Use canonical fixtures by default

- Default input: `fixtures/excel/` (tracked, stable)
- Allow overrides: `--tests test_files` (local dev)

### 2) Exclude oracle verification

Performance mode measures only the library under test.

- Read: no oracle involved.
- Write: stop timing immediately after `adapter.save_workbook(...)` returns.

Correctness remains the fidelity benchmark's job.

### 3) Multiple iterations + warmup

Default:

- warmup=3 (discard)
- iters=25 (record)

Report min/p50/p95.

Practical note: for very fast features (single-digit ms), timer resolution and fixed overhead can
dominate. In that case increase `iters`.

### 4) Memory is best-effort

`ru_maxrss` units differ by OS. We normalize to MB.

## Outputs

Keep perf outputs separate from fidelity outputs:

- Perf JSON: `<output>/perf/results.json`
- Perf Markdown: `<output>/perf/README.md`
- Perf CSV: `<output>/perf/matrix.csv`
- Perf history: `<output>/perf/history.jsonl` (append-only compact summary)

## Throughput/Scale Scenarios

The feature matrix fixtures are intentionally small and correctness-focused, so perf numbers can be
dominated by fixed overhead (imports, open/decode, save/zip, etc.). To measure per-action throughput,
we add a second fixture set that represents large workloads.

Design:

- Throughput fixtures live outside canonical Excel-generated fixtures.
- Test cases use a compact workload spec: `expected.workload`.
- The perf runner detects a "single-workload" manifest entry and runs loops over a cell range.
- Results include `op_count` + `op_unit` so we can compute throughput from timings.

Generation:

```bash
uv run python scripts/generate_throughput_fixtures.py
uv run excelbench perf --tests test_files/throughput_xlsx --output results_dev_perf_throughput --warmup 1 --iters 5 --breakdown
```

Docs:

- `fixtures/throughput_xlsx/README.md`

## CLI

New command:

```bash
uv run excelbench perf --tests fixtures/excel --output results_dev_perf --warmup 3 --iters 25 --breakdown
```

Key flags:

- `--tests/-t`: fixture root (must include `manifest.json`)
- `--output/-o`: output root (writes into `<output>/perf/`)
- `--feature/-f`: filter features (repeatable)
- `--adapter/-a`: filter adapters by name (repeatable)
- `--warmup`, `--iters`, `--breakdown`

## Implementation Mapping (repo)

- Runner: `src/excelbench/perf/runner.py`
- Renderer: `src/excelbench/perf/renderer.py`
- CLI: `src/excelbench/cli.py` (`excelbench perf`)

The perf runner reuses the fidelity harness' per-feature read/write exercise code paths, but skips
comparison/scoring.

Decision Log

- Chose a separate `excelbench perf` track over mixing into `excelbench benchmark` because it
  avoids timing oracle verification and keeps schemas stable while iterating.
