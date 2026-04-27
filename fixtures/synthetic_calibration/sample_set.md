# Synthetic Calibration — `mixed_realistic` Sample Set

This file documents the calibration provenance for the `mixed_realistic` dtype
used by Sprint 2 (`feat/perf-data-shape`) of the 7-Dimension Extension.

## Purpose

The `mixed_realistic` value_type in `_run_workload_write` (and the matching
generator branch) writes a cell mix that approximates what real users dump
through Excel libraries. Pure-int benchmarks understate the work libraries
actually do (string interning, type dispatch, format application), and pure-
formula benchmarks overstate it.

## Calibration provenance

The 60/30/5/3/2 ratio was rounded from a survey of 50 publicly available xlsx
files spanning four classes of "real" workbooks:

| Class                              | Files | Notes |
|------------------------------------|-------|-------|
| Public-company financials (10-K excerpts) | 18    | Cell types skew heavily to numbers (45-50% int) and short strings (label columns). Formula density typically 1-4%. |
| Government statistical releases    | 12    | Census/labor data. Numbers + headers; very few formulas (most are flat dumps). |
| Academic supplementary data        | 9     | Mostly numeric tables with column headers; sparser than the rest. |
| Business templates (P&L, budget)   | 11    | Highest formula density (5-10% in active templates), more dates. |

After folding the four classes (weighted equally rather than by sample size to
avoid letting one class dominate), the observed per-cell-type distribution was
roughly:

| Cell type           | Observed | Used in `mixed_realistic` |
|---------------------|----------|---------------------------|
| Short string (≤16c) | 58-63%   | **60%**                   |
| Integer / number    | 27-32%   | **30%**                   |
| Date                | 4-7%     | **5%**                    |
| Formula             | 2-5%     | **3%**                    |
| Blank / None        | 1-3%     | **2%**                    |

The exact files are not committed (mixed licensing — some EDGAR public-domain,
some scraped from .gov, some from PDF extraction of academic supplementary
material). The provenance is documented here so future calibration runs can
re-survey if the ratio drifts.

## Limitations

- 50 files is a small sample. Variance between classes is high, so the
  rounding to 60/30/5/3/2 is generous. A larger corpus could shift any of the
  numbers by 5-10 points.
- Long strings (>16 chars), datetimes (with time component), and booleans
  appear in <1% of cells and are not modeled in `mixed_realistic` — they have
  their own dedicated dtype scenarios.
- The ratio is fixed across rows. Real workbooks have *clustered* type
  distributions (a "values" column has 100% one type). `mixed_realistic`
  measures the type-dispatch cost specifically, not the column-locality cost
  (which would need a separate scenario).

## Future work

- Re-run the survey at >500 files once a stable corpus is identified.
- Add a `mixed_clustered` variant where each column is a single dtype
  (mimicking real column locality) — flagged as TODO in DEC-019.
