# ExcelBench Methodology

## Purpose

ExcelBench measures **feature fidelity** for spreadsheet libraries -- how accurately they can read and write Excel features compared to native Excel. It also measures **performance** (throughput, memory) as a secondary axis.

The default public presentation is still **Python-first** because most users are choosing between Python libraries. Cross-language comparisons are supported as a secondary ecosystem-context tier.

See `docs/trackers/cross-language-comparison-strategy.md` for the comparison policy.

## Fidelity vs Performance

- **Fidelity** evaluates correctness and completeness of features (formats, formulas, borders, etc.).
- **Performance** measures throughput (cells/s) and memory (RSS, tracemalloc) for read/write workloads.

Both tracks use the same adapter framework but run independently: fidelity uses oracle verification, performance skips it.

## Oracle Strategy

ExcelBench uses a **hybrid write verification** model:
- **Primary oracle:** Excel itself via `xlwings` (highest fidelity).
- **Fallback oracle:** `openpyxl` when Excel is unavailable (CI / headless).

This allows reliable local verification while keeping CI runnable.

## Fixtures Policy

Two fixture paths exist:
- `fixtures/excel/` (tracked): **canonical Excel-generated files** used in CI.
- `fixtures/excel_xls/` (tracked): **canonical .xls fixtures** for legacy format testing.
- `test_files/` (gitignored): local scratch output for development.

Generate canonical fixtures with:
```bash
uv run excelbench generate --output fixtures/excel        # .xlsx
uv run excelbench generate-xls --output fixtures/excel_xls  # .xls
```

## Feature Tiers

Features are organized into tiers reflecting complexity:

| Tier | Features | Count |
|------|----------|-------|
| **Tier 0** (Core) | cell_values, formulas, multiple_sheets | 3 |
| **Tier 1** (Formatting) | alignment, background_colors, borders, dimensions, number_formats, text_formatting | 6 |
| **Tier 2** (Advanced) | comments, conditional_formatting, data_validation, freeze_panes, hyperlinks, images, merged_cells, pivot_tables | 8 |
| **Tier 3** (Workbook Metadata) | named_ranges, tables | 2 |

**Framework coverage:** 19 modeled features total.
**Fresh release snapshot:** 19 tested features, where 18 are scoreable per-library in
the 2026-04-29 wheel-backed WolfXL 2.0 rerun (pivot_tables is tested but N/A across adapters on macOS fixtures).

## Scoring

Pass-rate model per feature per adapter:
- **3** (Green): all tests pass -- complete fidelity
- **2** (Yellow): >= 80% pass -- functional for common cases
- **1** (Orange): >= 50% pass -- basic recognition, significant limitations
- **0** (Red): < 50% pass -- errors, corruption, or data loss
- **N/A**: adapter does not support this operation direction (read/write)

Test cases have importance weights: `basic` (must-pass) vs `edge` (bonus). Both count toward pass rate but edge-case failures are noted in diagnostics.

## Reproducibility

All results are written to JSON as the source of truth:
- `results/xlsx/results.json` -- XLSX fidelity results
- `results/xls/results.json` -- XLS fidelity results
- `results/perf/results.json` -- Performance results

Renderers produce human-readable output (README.md, matrix.csv, heatmap, HTML dashboard) from these JSON files. The rendering step is deterministic and can be re-run independently:

```bash
uv run excelbench report --input results/xlsx/results.json --output results/xlsx
```

## Comparison policy

- **Primary tables** answer the Python replacement question.
- **Secondary tables** can include cross-language libraries such as Apache POI, ClosedXML, Excelize, and ExcelJS.
- Cross-language results should be labeled as ecosystem context unless the report is explicitly cross-language-focused.
- The current rollout plan for the first two cross-language candidates lives in `docs/trackers/apache-poi-excelize-rollout-plan.md`.
