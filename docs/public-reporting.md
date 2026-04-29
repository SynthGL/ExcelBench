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
| `results/perf/README.md` | 2026-04-20 | Newer performance rerun with fresher timing data |
| `results-release-2026-04-28/README.md` | 2026-04-29 | Fresh wheel-backed WolfXL 2.0 fidelity rerun |
| `results-release-2026-04-28/perf/README.md` | 2026-04-29 | Matching wheel-backed performance rerun |
| `results-cross-language/README.md` | 2026-04-29 | Checked-in cross-language context snapshot for Apache POI and Excelize |
| `results-cross-language-pivots/README.md` | 2026-04-29 | Separate pivot capability artifact for cross-language helpers |

## Safe Claims Right Now

- ExcelBench provides reproducible fidelity scoring across multiple Python spreadsheet libraries.
- The methodology and raw JSON artifacts are available in-repo.
- The checked-in results are dated snapshots, not timeless truths.
- The fresh WolfXL 2.0 wheel-backed release snapshot is available separately from the older historical baseline.
- A separate cross-language context snapshot is available for ecosystem positioning and should be cited as a separate lane from the Python hero table.
- A separate pivot capability artifact is available for cross-language helpers and should be cited as a capability note, not as a scored lane.

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
