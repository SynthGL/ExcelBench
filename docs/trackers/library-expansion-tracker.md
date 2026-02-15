# Library Expansion Tracker

Created: 02/08/2026 02:09 PM PST (via pst-timestamp)

## Overview

Tracks the status of all Excel library adapters in ExcelBench — implemented,
planned, and rejected. Each adapter wraps a specific library behind the
`ExcelAdapter` interface to produce comparable fidelity scores across 18
features (Tier 0/1/2/3).

Quick links:
- Benchmark results (xlsx): `results/xlsx/README.md`
- Benchmark results (xls): `results/xls/README.md`
- Adapter base class: `src/excelbench/harness/adapters/base.py`
- Adapter registry: `src/excelbench/harness/adapters/__init__.py`

## Adapter Inventory

### Implemented — xlsx profile (12 adapters)

| # | Library | Version | Lang | Caps | Read Score | Write Score | Green (R) | Green (W) | Notes |
|---|---------|---------|------|------|-----------|-------------|-----------|-----------|-------|
| 1 | openpyxl | 3.1.5 | py | R+W | 48/48 | 48/48 | 16/16 | 16/16 | Reference adapter, full fidelity |
| 2 | xlsxwriter | 3.2.9 | py | W | — | 48/48 | — | 16/16 | Write-only, full fidelity |
| 3 | python-calamine | 0.6.1 | py | R | 5/48 | — | 1/16 | — | Read-only, value+sheet only |
| 4 | pylightxl | 1.61 | py | R+W | 9/48 | 9/48 | 2/16 | 2/16 | Lightweight, value-only |
| 5 | pyexcel | 0.7.4 | py | R+W | 10/48 | 12/48 | 2/16 | 3/16 | Meta-library wrapping openpyxl |
| 6 | xlrd | 2.0.2 | py | R | — | — | — | — | .xls-only, not scored on xlsx |
| 7 | xlwt | 1.3.0 | py | W | — | 17/48 | — | 4/16 | .xls writer, limited xlsx compat |
| 8 | pandas | 3.0.0 | py | R+W | 5/48 | 12/48 | 1/16 | 3/16 | Abstraction-cost adapter (wraps openpyxl) |
| 9 | **openpyxl-readonly** | 3.1.5 | py | R | 10/48 | — | 3/16 | — | Streaming read mode, limited formatting |
| 10 | **xlsxwriter-constmem** | 3.2.9 | py | W | — | 43/48 | — | 13/16 | Row-major write, no images/comments |
| 11 | **polars** | 1.38.1 | py/rust | R | 4/48 | — | 0/16 | — | Rust calamine backend, type coercion cost |
| 12 | **tablib** | 3.9.0 | py | R+W | 10/48 | 12/48 | 2/16 | 3/16 | Dataset model wrapping openpyxl |

### Implemented — xls profile (2 adapters)

| # | Library | Version | Caps | Green (R) | Notes |
|---|---------|---------|------|-----------|-------|
| 1 | python-calamine | 0.6.1 | R | 2/4 | Cross-format reader |
| 2 | xlrd | 2.0.2 | R | 4/4 | Full .xls read fidelity |

### Implemented — Rust/PyO3 (5 adapters, require compiled extension)

| # | Library | Lang | Caps | Green | Status | Notes |
|---|---------|------|------|-------|--------|-------|
| 1 | calamine (rust) | rust | R | 1/18 | Built, not in CI | Cell values + sheet names only |
| 2 | calamine-styled | rust | R | **17R/18** | **Tier S-** | Full read fidelity (missing: images, diagonal borders=1) |
| 3 | rust_xlsxwriter | rust | W | **17W/18** | **Tier S-** | Full write fidelity (missing: images) |
| 4 | **pycalumya** (calamine+rxw) | rust | R+W | **17R+17W/18** | **Tier S-** | Hybrid: calamine read + rust_xlsxwriter write |
| 5 | **pyumya** (umya-spreadsheet) | rust | R+W | **13R+15W/18** | **Tier A** | See pyumya notes below |

### Planned / Candidate

| ID | Library | Lang | Caps | Priority | Rationale |
|----|---------|------|------|----------|-----------|
| A2 | odfpy | py | R+W | P2 | ODS format support (not xlsx) |
| A3 | et-xmlfile | py | — | P3 | Low-level XML streaming (used by openpyxl internally) |

### Rejected / Out of Scope

| Library | Reason |
|---------|--------|
| xlwings (as adapter) | Already used as test-file generator / oracle; not a parsing library |
| csv | Not an Excel format |
| openpyxl write_only | Streaming write API — requires different adapter pattern (no random cell access) |

## Score Summary — xlsx profile

Extracted from `results/xlsx/README.md` (02/14/2026 run, 14 adapters, 18 features):

```
Feature              openpyxl  xlsxwriter  constmem  calamine  cal(rs)  cal-styled  opxl-ro  pylightxl  pyexcel  xlwt  pandas  polars  tablib  umya       rxw(rs)  pycalumya
                     R  W      W           W         R         R        R           R        R  W        R  W     W     R  W    R       R  W    R  W       W        R  W
cell_values          3  3      3           3         1         1        3           3        3  1        3  3     3     1  3    1       3  3    3  3       3        3  3
formulas             3  3      3           3         0         0        3           3        0  3        0  3     0     0  3    0       0  3    3  3       3        3  3
text_formatting      3  3      3           3         0         0        3           0        0  0        0  0     1     0  0    0       0  0    3  3       3        3  3
background_colors    3  3      3           3         0         0        3           0        0  0        0  0     1     0  0    0       0  0    3  3       3        3  3
number_formats       3  3      3           3         0         0        3           0        0  0        0  1     3     0  0    0       0  1    3  3       3        3  3
alignment            3  3      3           3         1         1        3           1        0  1        1  1     3     1  1    1       1  1    1  1       3        3  3
borders              3  3      3           3         0         0        1           0        0  0        0  0     1     0  0    0       0  0    3  3       3        1  3
dimensions           3  3      3           1         0         0        3           0        0  0        0  0     1     0  0    0       0  0    1  3       3        3  3
multiple_sheets      3  3      3           3         3         3        3           3        3  3        3  3     3     3  3    1       3  3    3  3       3        3  3
merged_cells         3  3      3           3         0         0        3           0        0  0        0  0     0     0  0    0       0  0    3  3       3        3  3
conditional_format   3  3      3           3         0         0        3           0        0  0        0  0     0     0  0    0       0  0    3  0       3        3  3
data_validation      3  3      3           3         0         0        3           0        0  0        0  0     0     0  0    0       0  0    3  3       3        3  3
hyperlinks           3  3      3           3         0         0        3           0        0  0        0  0     0     0  0    0       0  0    0  0       3        3  3
images               3  3      3           0         0         0        0           0        0  0        0  0     0     0  0    0       0  0    0  3       0        0  0
comments             3  3      3           0         0         0        3           0        0  0        0  0     0     0  0    0       0  0    3  3       3        3  3
freeze_panes         3  3      3           3         0         0        3           0        0  0        0  0     0     0  0    0       0  0    3  3       3        3  3
named_ranges         3  3      —           —         0         0        3           0        0  0        0  0     0     0  0    0       0  0    3  3       3        3  3
tables               3  3      —           —         0         0        3           0        0  0        0  0     0     0  0    0       0  0    2  3       3        3  3
```

## pyumya Notes (umya-spreadsheet 2.3.3)

**Current scores (02/14/2026):** Read 13/18 green (43pts), Write 15/18 green (46pts)

**Non-green features and root causes:**

| Feature | Read | Write | Root Cause | Fixable? |
|---------|------|-------|------------|----------|
| alignment | 1 | 1 | umya's `Alignment` struct has no `indent` field | No — upstream limitation |
| conditional_formatting | 3 | 0 | Write path not implemented in adapter | Possible — needs investigation |
| dimensions | 1 | 3 | Read path returns approximate dimensions | Possible — needs investigation |
| hyperlinks | 0 | 0 | umya's XLSX reader never parses `tooltip` attribute from `<hyperlink>` XML | No — upstream bug |
| images | 0 | 3 | Read path returns `None` for image data/path | No — upstream limitation |
| tables | 2 | 3 | Read path returns partial table metadata | Possible — needs investigation |

**Fixed in this session:**
- **borders** (R:1→3, W:1→3): Diagonal border direction flags (`diagonal_up`/`diagonal_down`) were not being read or set. Fixed by checking `borders.get_diagonal_up()` / `borders.get_diagonal_down()` on read and calling `borders.set_diagonal_up(true)` / `borders.set_diagonal_down(true)` on write.

## Abstraction Cost Analysis

### Value-only wrappers (pandas vs pyexcel vs tablib vs polars)

All four wrap openpyxl or calamine internally. Key differences:

| Metric | pandas | pyexcel | tablib | polars | Winner |
|--------|--------|---------|--------|--------|--------|
| cell_values read | 1 (errors→NaN) | 3 | 3 | 1 (type coercion) | pyexcel/tablib |
| cell_values write | 3 | 3 | 3 | — | tie |
| formulas read | 0 | 0 | 0 | 0 | tie |
| formulas write | 3 | 3 | 3 | — | tie |
| alignment read | 1 | 1 | 1 | 1 | tie |
| number_formats write | 0 | 1 | 1 | — | pyexcel/tablib |
| Green features (R) | 1/16 | 2/16 | 2/16 | 0/16 | pyexcel/tablib |
| Green features (W) | 3/16 | 3/16 | 3/16 | — | tie |

**Key findings:**
- **pandas** loses error values (`#DIV/0!`, `#N/A`) because DataFrames coerce them to `NaN`
- **polars** loses even more due to columnar type coercion — mixed-type columns become strings, and multi-sheet support scores 1 (not 3) due to API limitations
- **tablib** matches pyexcel exactly — both preserve error values through their cell iterators
- **pyexcel** and **tablib** are the safest value-only abstractions for reads

### openpyxl default vs readonly mode

| Metric | openpyxl (default) | openpyxl-readonly | Difference |
|--------|-------------------|-------------------|------------|
| Green features (R) | 16/16 | 3/16 | -13 |
| Read score | 48/48 | 10/48 | -38 |
| Pass rate | 100% | 24% | -76pp |

**Key finding:** Read-only mode loses ALL formatting metadata (text_formatting, borders, background_colors, number_formats, dimensions, comments, images, hyperlinks, merged_cells, conditional_formatting, data_validation, freeze_panes). It preserves only cell_values, formulas, and multiple_sheets.

### xlsxwriter default vs constant_memory mode

| Metric | xlsxwriter (default) | xlsxwriter-constmem | Difference |
|--------|---------------------|---------------------|------------|
| Green features (W) | 16/16 | 13/16 | -3 |
| Write score | 48/48 | 43/48 | -5 |
| Pass rate | 100% | 94% | -6pp |

**Key finding:** constant_memory mode loses images (not supported), comments (not supported), and dimensions (row-major write order limits control). All formatting features (text, borders, colors, alignment, number_formats, conditional_formatting, data_validation, hyperlinks, freeze_panes, merged_cells) are fully preserved.

## Checklist — Expansion Tasks

- [x] A1: pandas adapter (value-only R+W, abstraction-cost measurement)
- [x] A4: openpyxl readonly mode adapter
- [x] A5: xlsxwriter constant_memory mode adapter
- [x] A6: polars adapter (Rust DataFrame reader)
- [x] A7: tablib adapter
- [ ] A2: odfpy adapter (ODS format)
- [ ] A3: et-xmlfile investigation

## Session Log (append-only)

### 02/08/2026 02:09 PM PST (via pst-timestamp)
- Worked on: A1 (pandas adapter)
- Committed: `46b8a87 feat(adapter): add pandas adapter measuring abstraction cost vs openpyxl`
- Scores: Read 1/3 cell_values (errors→NaN), Write 3/3 cell_values; 1/16 green R, 3/16 green W
- Benchmark: Regenerated full xlsx+xls profiles with all 8 Python adapters
- Decisions: pandas errors-as-NaN is a genuine abstraction cost, not a bug to fix
- Next: A4 (openpyxl readonly mode) or A6 (polars) — both measure optimization variants

### 02/08/2026 — Session 2
- Worked on: A4 (openpyxl-readonly), A5 (xlsxwriter-constmem), A6 (polars), A7 (tablib)
- Commits:
  - `89ec792` build: add polars, fastexcel, tablib dependencies
  - `d39a5a0` feat(adapter): add xlsxwriter constant_memory adapter (A5)
  - `56a71ec` feat(adapter): add openpyxl read-only adapter (A4)
  - `fa16ecb` feat(adapter): add polars adapter (A6)
  - `8829ad9` feat(adapter): add tablib adapter (A7)
  - `3e9aa05` feat(adapter): register all 4 in adapter registry
- Test results: 1067 passed, 39 skipped, 6 xfailed (69 new tests, 0 regressions)
- Scores:
  - openpyxl-readonly: 10/48 R, 3/16 green (loses all formatting in streaming mode)
  - xlsxwriter-constmem: 43/48 W, 13/16 green (loses images/comments/dimensions)
  - polars: 4/48 R, 0/16 green (type coercion cost + limited multi-sheet)
  - tablib: 10/48 R + 12/48 W, 2/16 green R + 3/16 green W (matches pyexcel)
- Key findings: See Abstraction Cost Analysis section above
- Total adapters: 12 Python xlsx + 2 xls + 3 Rust/PyO3 = 17

### 02/14/2026 — Fresh benchmark + pyumya border fix
- Ran full fidelity benchmark (xlsx + xls profiles, 14 adapters, 18 features)
- Ran performance benchmark with phase breakdown
- Regenerated all visualizations: report, heatmap, dashboard, scatter, html
- Committed: `aba3822 chore: refresh fidelity + performance benchmarks and regenerate dashboard`
- Investigated pyumya regressions: alignment indent (upstream), hyperlinks tooltip (upstream), images read (upstream), borders diagonal (fixable)
- Fixed borders.rs: diagonal direction flags (`diagonal_up`/`diagonal_down`) not being read/set
- Committed: `28135d2 fix(pyumya): implement diagonal border direction flags for read and write`
- Re-ran full benchmark with border fix
- pyumya scores after fix: R:13/18 green (43pts), W:15/18 green (46pts) — borders R:1→3, W:1→3
- Total: 1155 tests passed, 0 regressions

### 02/14/2026 — pycalumya hybrid adapter + Sprint 1+2 completion
- Created **pycalumya** hybrid adapter: calamine-styled (read) + rust_xlsxwriter (write)
- Codex Sprint 1: formulas read, column-width padding in Rust, 4 Tier 2 R+W (merged, hyperlinks, comments, freeze_panes)
- Codex Sprint 2: ooxml_util.rs extraction, 4 more Tier 2/3 R+W (CF, DV, named_ranges, tables)
- Fixed cell_values regression (error-formula mapping matching openpyxl's ERROR_FORMULA_MAP)
- Fixed hyperlink internal detection (`n.location.is_some() && n.rid.is_none()`)
- Scores:
  - calamine-styled: R:17/18 green (borders=1 diagonal, images=0)
  - rust_xlsxwriter: W:17/18 green (images=0)
  - **pycalumya: R:17/18 + W:17/18** — second only to openpyxl (18/18)
- Performance: pycalumya reads 0.3-1.5ms per feature (except cell_values 14ms first access)
- Total adapters: 12 Python xlsx + 2 xls + 5 Rust/PyO3 = 19
- Total tests: 1155 passed, 0 regressions

### 02/13/2026 — CI fixes + pyo3 bump + Tier 3 scoring
- Fixed all 31 mypy strict-mode errors across 5 Python files
- Bumped pyo3 0.22 → 0.24 (security fix + API modernization)
  - Replaced ToPyObject → IntoPyObject, renamed new_bound → new, empty_bound → empty
  - Closed Dependabot PR #7 (superseded)
- Tier 3 features (named_ranges + tables) now officially scored: 18 total features
  - openpyxl: 18/18 (S tier, only library with full Tier 3 support)
  - pyumya: 16/18 (A tier, Tier 3 returning not_found — Rust backend limitation)
  - xlsxwriter: 16/18 (A tier, write-only)
- Regenerated HTML dashboard with latest results
- Updated CLAUDE.md and tracker docs with current state
