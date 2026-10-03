# Public Reporting Status

This page explains what the current ExcelBench repo artifacts mean and how to cite them accurately.

## Status

- Package version: `0.1.0`
- Repository: `SynthGL/ExcelBench`
- Last local verification pass: 2026-04-29
- Current public snapshot in `results/xlsx/` and `results/DASHBOARD.md`: 2026-02-17
- Newer perf snapshot in `results/perf/`: 2026-04-20
- Fresh WolfXL 2.0 wheel-backed release snapshot: `results-release-2026-04-28/` generated 2026-04-29 UTC
- Current checked-in cross-language context snapshot: `results-cross-language/` generated 2026-04-29 UTC
- Current checked-in cross-language pivot capability artifact: `results-cross-language-pivots/` generated 2026-04-29 UTC
- Current competitor snapshot: `results-2026-10-03/` (fidelity, mutation, calc, cross-language) generated 2026-10-03 UTC with WolfXL 2.0.8
- Previous competitor snapshot: `results-2026-10-02/` (fidelity, mutation, calc, cross-language, perf) generated 2026-10-02 UTC with WolfXL 2.0.5

## How To Cite ExcelBench Safely

1. Treat every results directory as a timestamped snapshot.
2. Cite the date, platform, and workload profile whenever quoting a number.
3. Separate historical public snapshots from release-blocking reruns.
4. If fidelity and perf were generated on different dates, say that explicitly.

## Current Artifact State

| Artifact | Date | Meaning |
|---|---|---|
| `results/xlsx/README.md` | 2026-02-17 | Current checked-in public XLSX fidelity snapshot |
| `results/DASHBOARD.md` | 2026-02-17 | Current checked-in combined dashboard snapshot |
| `results/perf/README.md` | 2026-04-20 | Historical performance snapshot; its WolfXL column measured private backend objects and is superseded by `results-2026-10-02/perf/` |
| `results-release-2026-04-28/README.md` | 2026-04-29 | Fresh wheel-backed WolfXL 2.0 fidelity rerun |
| `results-release-2026-04-28/perf/README.md` | 2026-04-29 | Historical wheel-backed performance rerun; its WolfXL column measured private backend objects and is superseded by `results-2026-10-02/perf/` |
| `results-cross-language/README.md` | 2026-04-29 | Checked-in cross-language context snapshot for Apache POI and Excelize |
| `results-cross-language-pivots/README.md` | 2026-04-29 | Separate pivot capability artifact for cross-language helpers |
| `results-2026-09-08/xlsx/README.md` | 2026-09-08 | Competitor fidelity snapshot: 22 features, 14 Python adapters, WolfXL 2.1.0 + aspose-cells-foss |
| `results-2026-09-08/mutation/README.md` | 2026-09-08 | Template-mutation lane (wall time, RSS, preservation) |
| `results-2026-09-08/calc/README.md` | 2026-09-08 | Formula-recalculation tier vs LibreOffice oracle on a cache-free fixture |
| `results-2026-09-08/cross-language/README.md` | 2026-09-08 | zavora-xlsx 0.1.2 Rust writer cross-language context |
| `results-2026-10-02/xlsx/README.md` | 2026-10-02 | Competitor fidelity snapshot: 21 scored features (`pivot_tables` unscored), 17 adapters including sheetjs, exceljs, libreoffice; WolfXL 2.0.5 |
| `results-2026-10-02/mutation/README.md` | 2026-10-02 | Template-mutation lane: preservation is valid; wall time and RSS are not comparable (host under heavy unrelated load) |
| `results-2026-10-02/calc/README.md` | 2026-10-02 | Formula-recalculation tier vs LibreOffice 26.8.0.3 oracle on the cache-free 133-formula fixture |
| `results-2026-10-02/cross-language/README.md` | 2026-10-02 | Cross-language context: Apache POI 5.5.1, Excelize 2.10.1, zavora-xlsx 0.1.2 |
| `results-2026-10-02/perf/README.md` | 2026-10-02 | Performance snapshot: April's 19 features x 13 Python adapters, warmup 3, 25 iterations, WolfXL 2.0.5 through its public API, Apple M4 Pro |
| `results-2026-10-03/xlsx/README.md` | 2026-10-03 | Competitor fidelity snapshot: 21 scored features (`pivot_tables` unscored), the same 17 adapters as 2026-10-02; WolfXL 2.0.8, python-calamine 0.8.2, pandas 3.0.6, polars 1.44.2, pyexcel 0.7.6, tablib 3.10.0 |
| `results-2026-10-03/mutation/README.md` | 2026-10-03 | Template-mutation lane (content-model preservation, wall time, RSS); see its Scope note for host load |
| `results-2026-10-03/calc/README.md` | 2026-10-03 | Formula-recalculation tier vs LibreOffice 26.8.0.3 oracle; WolfXL Community 2.0.8 and Commercial 2.3.0 |
| `results-2026-10-03/cross-language/README.md` | 2026-10-03 | Cross-language context: Apache POI 5.5.1, Excelize 2.11.0, zavora-xlsx 0.1.2 |

## Safe Claims Right Now

- ExcelBench provides reproducible fidelity scoring across multiple Python spreadsheet libraries.
- The methodology and raw JSON artifacts are available in-repo.
- The checked-in results are dated snapshots, not timeless truths.
- The fresh WolfXL 2.0 wheel-backed release snapshot is available separately from the older historical baseline.
- A separate cross-language context snapshot is available for ecosystem positioning and should be cited as a separate lane from the Python hero table.
- A separate pivot capability artifact is available for cross-language helpers and should be cited as a capability note, not as a scored lane.
- The 2026-09-08 competitor snapshot extends the scored matrix to 22 features (Tier 4 is openpyxl-structural, not Excel-authored) and adds aspose-cells-foss; its fidelity numbers must not be mixed with /18 denominators from earlier snapshots.
- The mutation and calc lanes are separate decision surfaces (modify-integrity and calculation correctness) and must be cited as their own artifacts, never folded into feature-parity claims.

## Claims That Need Fresh Reruns

- Any statement that merges February fidelity and April perf into one "current" result without caveat.
- Any ecosystem ranking that implies a same-day apples-to-apples rerun if the artifacts are from different dates.
- Any statement that mixes the historical baseline and the WolfXL 2.0 release rerun without naming which snapshot is being cited.
- Any statement that treats the cross-language context snapshot as the same decision surface as the Python replacement snapshot.
- Any statement that treats the pivot capability artifact as if it were part of the scored cross-language matrix.

## Verification Commands

```bash
uv run pytest tests/ --cov-fail-under=65
uv run excelbench benchmark --tests fixtures/excel --output results
uv run excelbench perf --tests fixtures/excel --output results
uv run excelbench report --input results/xlsx/results.json --output results/xlsx
```

## Recommended README Policy

- Lead with what ExcelBench measures.
- Link directly to `METHODOLOGY.md`.
- Mark top-level result summaries as dated snapshots.
- Keep WolfXL-specific launch claims in sync with the WolfXL repo's release evidence page.
- Keep the cross-language context snapshot clearly separated from the Python-first comparison in README and launch copy.
- Keep the pivot capability artifact clearly separated from both the Python-first and cross-language scorecards.
