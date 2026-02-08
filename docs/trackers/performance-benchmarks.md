# Performance Benchmarks - Implementation Tracker

Created: 02/08/2026 05:07 AM PST (via pst-timestamp)

Quick links:
- Plan: `docs/plans/2026-02-08-performance-benchmarks.md`
- Perf runner: `src/excelbench/perf/runner.py`
- Perf renderer: `src/excelbench/perf/renderer.py`
- CLI: `src/excelbench/cli.py`

Status legend:
- [ ] not started
- [~] in progress
- [x] done
- [-] rejected

## Working State (keep updated)

Current focus:
- [ ] PB-M3: baseline runs

Current blockers:
- None

## Milestones

- [x] PB-M0: Spec locked (initial)
- [x] PB-M1: Perf runner exists
- [x] PB-M2: Perf renderer exists
- [x] PB-M3: First baseline run captured + logged
- [x] PB-M4: History / trend tracking (history.jsonl)
- [ ] PB-M5: Optional CI hook (non-blocking)

## Tasks (IDs)

P0 - Foundations
- [x] PB-001: Define perf metrics and units (wall/cpu/rss)
- [x] PB-002: Define JSON schema + versioning strategy
- [x] PB-003: Define iteration/warmup defaults and stats (min/p50/p95)

P1 - Implementation
- [x] PB-010: Add `excelbench perf` CLI command
- [x] PB-011: Implement perf runner (read totals)
- [x] PB-012: Implement perf runner (write totals; exclude verification)
- [x] PB-013: Implement optional phase breakdown
- [x] PB-014: Add perf renderer (markdown)
- [x] PB-015: Add perf renderer (CSV)
- [x] PB-016: Add perf history append (`history.jsonl`)

P2 - Validation / quality
- [x] PB-020: Add perf smoke test (schema + basic run; no timing asserts)
- [ ] PB-021: Validate ru_maxrss normalization on macOS vs Linux

P3 - Optional extensions
- [~] PB-030: Add scale fixtures (10k/100k rows) for throughput scenarios
- [ ] PB-031: Render perf profiles separately from feature matrix

## Session Log (append-only)

Template:

```text
### MM/DD/YYYY HH:MM AM/PM PST/PDT (via pst-timestamp)
- Worked on: PB-XXX, PB-YYY
- Decisions: ...
- Notes: ...
- Blockers: ...
- Next: ...
```

### 02/08/2026 05:07 AM PST (via pst-timestamp)
- Worked on: PB-010..PB-016
- Notes: perf runner+renderer+CLI implemented; baseline run still needs logging.
- Next: PB-M3 (run baseline and fill `docs/trackers/performance-benchmark-runs.md`).

### 02/08/2026 05:08 AM PST (via pst-timestamp)
- Worked on: PB-M3
- Notes: captured Tier 1 baseline run in `results_dev_perf_baseline_tier1/perf/` and logged environment/command.
- Next: run a full feature baseline (Tier 1+2) with smaller iters to sanity-check advanced features.

### 02/08/2026 01:50 PM PST (via pst-timestamp)
- Worked on: PB-M3 (additional baseline)
- Notes: captured full xlsx baseline in `results_dev_perf_baseline_full_xlsx/perf/` and logged it.
- Next: decide whether to add an xls perf baseline (`--profile xls` + `fixtures/excel_xls`).

### 02/08/2026 02:00 PM PST (via pst-timestamp)
- Worked on: PB-M3 (xls baseline)
- Notes: captured xls baseline in `results_dev_perf_baseline_full_xls/perf/` and logged it.
- Next: start PB-030 (throughput/scale fixture set).

### 02/08/2026 02:08 PM PST (via pst-timestamp)
- Worked on: PB-030
- Notes: added throughput generator (`scripts/generate_throughput_fixtures.py`), added workload support in perf runner, generated + ran a first throughput baseline (openpyxl only).
- Next: expand throughput scenarios beyond cell values (formulas/styles) and run multi-adapter baselines with longer timeouts.

### 02/08/2026 02:16 PM PST (via pst-timestamp)
- Worked on: PB-030 (formulas workload)
- Notes: added `formulas_10k` throughput fixture + workload op support; reran throughput baseline for openpyxl.
- Next: add at least one style-heavy workload (e.g., background fill) at smaller N (1k) to avoid huge files.

### 02/08/2026 02:18 PM PST (via pst-timestamp)
- Worked on: PB-030 (multi-adapter throughput run)
- Notes: ran a first multi-adapter throughput baseline (warmup=0 iters=1) for openpyxl/xlsxwriter/pylightxl/pyexcel/python-calamine.
- Next: consider adding a lower-N workload (1k) for very slow per-cell APIs to keep runs reasonable.

### 02/08/2026 02:21 PM PST (via pst-timestamp)
- Worked on: PB-030 (style workload)
- Notes: added a first style-heavy workload (`background_colors_1k`) and verified it runs end-to-end in perf.
- Next: add `number_formats_1k` and `alignment_1k` (and keep borders small).

### 02/08/2026 02:23 PM PST (via pst-timestamp)
- Worked on: PB-030 (more workloads)
- Notes: added `number_formats_1k` + `alignment_1k` workloads; regenerated throughput suite now has 5 scenarios.
- Next: consider a `borders_200` workload (borders are expensive and bloat XLSX styles quickly).
