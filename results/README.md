# ExcelBench Results

*Generated: 2026-04-20 07:06 UTC*
*Profile: xlsx*
*Excel Version: 16.105.3*
*Platform: Darwin-arm64*

## Overview

> Condensed view — shows the **best score** across read/write for each library. See [Full Results Matrix](#full-results-matrix) for the complete R/W breakdown.

**Tier 0 — Basic Values**

| Feature | openpyxl | opxl-readonly | pandas | polars | pyexcel | pylightxl | calamine | rust_xlsxwriter | tablib | wolfxl | xlsxwriter | xlsx-constmem | xlwt |
|---------|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|
| Cell Values | 🟢 | 🟢 | 🟢 | 🟠 | 🟢 | 🟢 | 🟠 | 🟢 | 🟢 | 🟢 | 🟢 | 🟢 | 🟢 |
| Formulas | 🟢 | 🟢 | 🟢 | 🔴 | 🟢 | 🟢 | 🔴 | 🟢 | 🟢 | 🟢 | 🟢 | 🟢 | 🔴 |
| Sheets | 🟢 | 🟢 | 🟢 | 🟠 | 🟢 | 🟢 | 🟢 | 🟢 | 🟢 | 🟢 | 🟢 | 🟢 | 🟢 |

**Tier 1 — Formatting**

| Feature | openpyxl | opxl-readonly | pandas | polars | pyexcel | pylightxl | calamine | rust_xlsxwriter | tablib | wolfxl | xlsxwriter | xlsx-constmem | xlwt |
|---------|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|
| Alignment | 🟢 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🟠 | 🔴 | 🟢 | 🟠 | 🟠 | 🟢 |
| Bg Colors | 🟢 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🟢 | 🔴 | 🟢 | 🟢 | 🟢 | 🟠 |
| Borders | 🟢 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🟢 | 🔴 | 🟢 | 🟢 | 🟢 | 🟠 |
| Dimensions | 🟢 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🟢 | 🔴 | 🟢 | 🟢 | 🟠 | 🟠 |
| Num Fmt | 🟢 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🟢 | 🔴 | 🟢 | 🟢 | 🟢 | 🟢 |
| Text Fmt | 🟢 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🟢 | 🔴 | 🟢 | 🟢 | 🟢 | 🟠 |

**Tier 2 — Advanced**

| Feature | openpyxl | opxl-readonly | pandas | polars | pyexcel | pylightxl | calamine | rust_xlsxwriter | tablib | wolfxl | xlsxwriter | xlsx-constmem | xlwt |
|---------|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|
| Comments | 🟢 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🟢 | 🔴 | 🟢 | 🟢 | 🔴 | 🔴 |
| Cond Fmt | 🟢 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🟢 | 🔴 | 🟢 | 🟢 | 🟢 | 🔴 |
| Validation | 🟢 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🟢 | 🔴 | 🟢 | 🟢 | 🟢 | 🔴 |
| Freeze | 🟢 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🟢 | 🔴 | 🟢 | 🟢 | 🟢 | 🔴 |
| Hyperlinks | 🟢 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🟢 | 🔴 | 🟢 | 🟢 | 🟢 | 🔴 |
| Images | 🟢 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🟢 | 🔴 | 🔴 |
| Merged | 🟢 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🟢 | 🔴 | 🟢 | 🟢 | 🟢 | 🔴 |

**Tier 3 — Workbook Metadata**

| Feature | openpyxl | opxl-readonly | pandas | polars | pyexcel | pylightxl | calamine | rust_xlsxwriter | tablib | wolfxl | xlsxwriter | xlsx-constmem | xlwt |
|---------|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|
| Named Ranges | 🟢 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🟢 | 🔴 | 🟢 | 🔴 | 🔴 | 🔴 |
| Tables | 🟢 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🟢 | 🔴 | 🟢 | 🔴 | 🔴 | 🔴 |

## Library Tiers

> Libraries ranked by their best capability (max of read/write green features).

| Tier | Library | Caps | Modify | Green Features | Summary |
|:----:|---------|:----:|:------:|:--------------:|---------|
| **S** | openpyxl | R+W | Rewrite | 18/18 | Reference adapter — full read + write fidelity |
| **S-** | wolfxl | R+W | Patch | 17/18 | Hybrid Rust+Python — near-full fidelity at 3-9x throughput |
| **S-** | rust_xlsxwriter | W | No | 16/18 | Rust write backend — used internally by WolfXL |
| **A** | xlsxwriter | W | No | 15/18 | Best write-only option — full formatting support |
| **B** | xlsxwriter-constmem | W | No | 12/18 | Memory-optimized write — loses images, comments, row height |
| **B** | xlwt | W | No | 4/18 | Legacy .xls writer — basic formatting subset |
| **C** | openpyxl-readonly | R | No | 3/18 | Streaming read — loses all formatting metadata |
| **C** | pandas | R+W | Rebuild | 3/18 | DataFrame abstraction — errors coerced to NaN on read |
| **C** | pyexcel | R+W | Rebuild | 3/18 | Meta-library wrapping openpyxl — preserves error values |
| **C** | tablib | R+W | Rebuild | 3/18 | Dataset wrapper — matches pyexcel on fidelity |
| **C** | pylightxl | R+W | Rebuild | 2/18 | Lightweight — basic values, no formatting API |
| **C** | python-calamine | R | No | 1/18 | Fast Rust-backed reader — cell values + sheet names only |
| **D** | polars | R | No | 0/18 | Rust DataFrame reader — columnar type coercion drops fidelity |

## Score Legend

| Score | Meaning |
|-------|---------|
| 🟢 3 | Complete — all basic and edge cases pass |
| 🟡 2 | Functional — all basic pass, one or more edge cases fail |
| 🟠 1 | Minimal — at least one basic case passes, but not all basic cases |
| 🔴 0 | Unsupported — errors or data loss |
| ➖ | Not applicable |

## Full Results Matrix

**Tier 0 — Basic Values**

| Feature | openpyxl (R) | openpyxl (W) | openpyxl-readonly (R) | pandas (R) | pandas (W) | polars (R) | pyexcel (R) | pyexcel (W) | pylightxl (R) | pylightxl (W) | python-calamine (R) | rust_xlsxwriter (W) | tablib (R) | tablib (W) | wolfxl (R) | wolfxl (W) | xlrd (R) | xlsxwriter (W) | xlsxwriter-constmem (W) | xlwt (W) |
|---------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|
| [cell_values](#cell_values-details) | 🟢 3 | 🟢 3 | 🟢 3 | 🟠 1 | 🟢 3 | 🟠 1 | 🟢 3 | 🟢 3 | 🟢 3 | 🟠 1 | 🟠 1 | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 | ➖ | 🟢 3 | 🟢 3 | 🟢 3 |
| [formulas](#formulas-details) | 🟢 3 | 🟢 3 | 🟢 3 | 🔴 0 | 🟢 3 | 🔴 0 | 🔴 0 | 🟢 3 | 🔴 0 | 🟢 3 | 🔴 0 | 🟢 3 | 🔴 0 | 🟢 3 | 🟢 3 | 🟢 3 | ➖ | 🟢 3 | 🟢 3 | 🔴 0 |
| [multiple_sheets](#multiple_sheets-details) | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 | 🟠 1 | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 | ➖ | 🟢 3 | 🟢 3 | 🟢 3 |

**Tier 1 — Formatting**

| Feature | openpyxl (R) | openpyxl (W) | openpyxl-readonly (R) | pandas (R) | pandas (W) | polars (R) | pyexcel (R) | pyexcel (W) | pylightxl (R) | pylightxl (W) | python-calamine (R) | rust_xlsxwriter (W) | tablib (R) | tablib (W) | wolfxl (R) | wolfxl (W) | xlrd (R) | xlsxwriter (W) | xlsxwriter-constmem (W) | xlwt (W) |
|---------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|
| [alignment](#alignment-details) | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🟠 1 | 🔴 0 | 🔴 0 | 🟢 3 | 🟠 1 | ➖ | 🟠 1 | 🟠 1 | 🟢 3 |
| [background_colors](#background_colors-details) | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🟢 3 | 🔴 0 | 🔴 0 | 🟢 3 | 🟢 3 | ➖ | 🟢 3 | 🟢 3 | 🟠 1 |
| [borders](#borders-details) | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🟢 3 | 🔴 0 | 🔴 0 | 🟢 3 | 🟢 3 | ➖ | 🟢 3 | 🟢 3 | 🟠 1 |
| [dimensions](#dimensions-details) | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🟢 3 | 🔴 0 | 🔴 0 | 🟢 3 | 🟢 3 | ➖ | 🟢 3 | 🟠 1 | 🟠 1 |
| [number_formats](#number_formats-details) | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🟢 3 | 🔴 0 | 🔴 0 | 🟢 3 | 🟢 3 | ➖ | 🟢 3 | 🟢 3 | 🟢 3 |
| [text_formatting](#text_formatting-details) | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🟢 3 | 🔴 0 | 🔴 0 | 🟢 3 | 🟢 3 | ➖ | 🟢 3 | 🟢 3 | 🟠 1 |

**Tier 2 — Advanced**

| Feature | openpyxl (R) | openpyxl (W) | openpyxl-readonly (R) | pandas (R) | pandas (W) | polars (R) | pyexcel (R) | pyexcel (W) | pylightxl (R) | pylightxl (W) | python-calamine (R) | rust_xlsxwriter (W) | tablib (R) | tablib (W) | wolfxl (R) | wolfxl (W) | xlrd (R) | xlsxwriter (W) | xlsxwriter-constmem (W) | xlwt (W) |
|---------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|
| [comments](#comments-details) | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🟢 3 | 🔴 0 | 🔴 0 | 🟢 3 | 🟢 3 | ➖ | 🟢 3 | 🔴 0 | 🔴 0 |
| [conditional_formatting](#conditional_formatting-details) | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🟢 3 | 🔴 0 | 🔴 0 | 🟢 3 | 🟢 3 | ➖ | 🟢 3 | 🟢 3 | 🔴 0 |
| [data_validation](#data_validation-details) | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🟢 3 | 🔴 0 | 🔴 0 | 🟢 3 | 🟢 3 | ➖ | 🟢 3 | 🟢 3 | 🔴 0 |
| [freeze_panes](#freeze_panes-details) | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🟢 3 | 🔴 0 | 🔴 0 | 🟢 3 | 🟢 3 | ➖ | 🟢 3 | 🟢 3 | 🔴 0 |
| [hyperlinks](#hyperlinks-details) | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🟢 3 | 🔴 0 | 🔴 0 | 🟢 3 | 🟢 3 | ➖ | 🟢 3 | 🟢 3 | 🔴 0 |
| [images](#images-details) | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | ➖ | 🟢 3 | 🔴 0 | 🔴 0 |
| [merged_cells](#merged_cells-details) | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🟢 3 | 🔴 0 | 🔴 0 | 🟢 3 | 🟢 3 | ➖ | 🟢 3 | 🟢 3 | 🔴 0 |
| [pivot_tables](#pivot_tables-details) | ➖ | ➖ | ➖ | ➖ | ➖ | ➖ | ➖ | ➖ | ➖ | ➖ | ➖ | ➖ | ➖ | ➖ | ➖ | ➖ | ➖ | ➖ | ➖ | ➖ |

**Tier 3 — Workbook Metadata**

| Feature | openpyxl (R) | openpyxl (W) | openpyxl-readonly (R) | pandas (R) | pandas (W) | polars (R) | pyexcel (R) | pyexcel (W) | pylightxl (R) | pylightxl (W) | python-calamine (R) | rust_xlsxwriter (W) | tablib (R) | tablib (W) | wolfxl (R) | wolfxl (W) | xlrd (R) | xlsxwriter (W) | xlsxwriter-constmem (W) | xlwt (W) |
|---------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|
| [named_ranges](#named_ranges-details) | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🟢 3 | 🔴 0 | 🔴 0 | 🟢 3 | 🟢 3 | ➖ | 🔴 0 | 🔴 0 | 🔴 0 |
| [tables](#tables-details) | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🟢 3 | 🔴 0 | 🔴 0 | 🟢 3 | 🟢 3 | ➖ | 🔴 0 | 🔴 0 | 🔴 0 |

## Notes

- **alignment**: Known limitation: pylightxl alignment write is a no-op because the library does not support formatting writes.
- **cell_values**: Known limitation: pylightxl cell-values write has date/boolean/error fidelity limits due to writer encoding behavior.
- **alignment**: Known limitation: python-calamine alignment read is limited because its API does not expose style/alignment metadata.
- **cell_values**: Known limitation: python-calamine can surface formula error cells as blank values in current API responses.
- **cell_values, formulas, ... (19 features)**: Not applicable: xlrd does not support .xlsx input
- **pivot_tables**: Unsupported on macOS without a Windows-generated pivot fixture (fixtures/excel/tier2/15_pivot_tables.xlsx).

## Statistics

| Library | Mode | Tests | Passed | Failed | Pass Rate | Green Features |
|---------|------|-------|--------|--------|-----------|----------------|
| openpyxl | R | 125 | 125 | 0 | 100% | 18/18 |
| openpyxl | W | 125 | 125 | 0 | 100% | 18/18 |
| openpyxl-readonly | R | 125 | 26 | 99 | 21% | 3/18 |
| pandas | R | 125 | 19 | 106 | 15% | 1/18 |
| pandas | W | 125 | 26 | 99 | 21% | 3/18 |
| polars | R | 125 | 17 | 108 | 14% | 0/18 |
| pyexcel | R | 125 | 22 | 103 | 18% | 2/18 |
| pyexcel | W | 125 | 26 | 99 | 21% | 3/18 |
| pylightxl | R | 125 | 22 | 103 | 18% | 2/18 |
| pylightxl | W | 125 | 22 | 103 | 18% | 2/18 |
| python-calamine | R | 125 | 19 | 106 | 15% | 1/18 |
| rust_xlsxwriter | W | 125 | 122 | 3 | 98% | 16/18 |
| tablib | R | 125 | 22 | 103 | 18% | 2/18 |
| tablib | W | 125 | 26 | 99 | 21% | 3/18 |
| wolfxl | R | 125 | 123 | 2 | 98% | 17/18 |
| wolfxl | W | 125 | 122 | 3 | 98% | 16/18 |
| xlsxwriter | W | 125 | 112 | 13 | 90% | 15/18 |
| xlsxwriter-constmem | W | 125 | 105 | 20 | 84% | 12/18 |
| xlwt | W | 125 | 72 | 53 | 58% | 4/18 |

## Libraries Tested

- **openpyxl** v3.1.5 (python) - read, write; modify: Rewrite
- **openpyxl-readonly** v3.1.5 (python) - read; modify: No
- **pandas** v3.0.0 (python) - read, write; modify: Rebuild
- **polars** v1.38.1 (python) - read; modify: No
- **pyexcel** v0.7.4 (python) - read, write; modify: Rebuild
- **pylightxl** v1.61 (python) - read, write; modify: Rebuild
- **python-calamine** v0.6.1 (python) - read; modify: No
- **rust_xlsxwriter** v0.79.4 (rust) - write; modify: No
- **tablib** v3.9.0 (python) - read, write; modify: Rebuild
- **wolfxl** vcal=0.4.0+rxw=0.79.4 (rust) - modify, read, write; modify: Patch
- **xlrd** v2.0.2 (python) - read; modify: No
- **xlsxwriter** v3.2.9 (python) - write; modify: No
- **xlsxwriter-constmem** v3.2.9 (python) - write; modify: No
- **xlwt** v1.3.0 (python) - write; modify: No

## Diagnostics Summary

| Group | Value | Count |
|-------|-------|-------|
| category | data_mismatch | 974 |
| category | internal | 31 |
| category | invalid_input | 61 |
| category | unsupported_feature | 156 |
| severity | error | 1066 |
| severity | warning | 156 |

### Diagnostic Details

| Feature | Library | Test Case | Operation | Category | Severity | Message |
|---------|---------|-----------|-----------|----------|----------|---------|
| cell_values | python-calamine | error_div0 | read | data_mismatch | error | Expected values did not match actual values: expected={'type': 'error', 'value': '#DIV/0!'}, actual={'type': 'blank'} |
| cell_values | python-calamine | error_na | read | data_mismatch | error | Expected values did not match actual values: expected={'type': 'error', 'value': '#N/A'}, actual={'type': 'blank'} |
| cell_values | python-calamine | error_value | read | data_mismatch | error | Expected values did not match actual values: expected={'type': 'error', 'value': '#VALUE!'}, actual={'type': 'blank'} |
| cell_values | pylightxl | date_standard | write | data_mismatch | error | Expected values did not match actual values: expected={'type': 'date', 'value': '2026-02-04'}, actual={'type': 'string', 'value': '2026-02-04'} |
| cell_values | pylightxl | datetime | write | data_mismatch | error | Expected values did not match actual values: expected={'type': 'datetime', 'value': '2026-02-04T10:30:45'}, actual={'type': 'string', 'value': '2026-02-04T10:30:45'} |
| cell_values | pylightxl | boolean_true | write | data_mismatch | error | Expected values did not match actual values: expected={'type': 'boolean', 'value': True}, actual={'type': 'number', 'value': 1} |
| cell_values | pylightxl | boolean_false | write | data_mismatch | error | Expected values did not match actual values: expected={'type': 'boolean', 'value': False}, actual={'type': 'number', 'value': 0} |
| cell_values | pandas | error_div0 | read | data_mismatch | error | Expected values did not match actual values: expected={'type': 'error', 'value': '#DIV/0!'}, actual={'type': 'blank'} |
| cell_values | pandas | error_na | read | data_mismatch | error | Expected values did not match actual values: expected={'type': 'error', 'value': '#N/A'}, actual={'type': 'blank'} |
| cell_values | pandas | error_value | read | data_mismatch | error | Expected values did not match actual values: expected={'type': 'error', 'value': '#VALUE!'}, actual={'type': 'blank'} |
| cell_values | polars | error_div0 | read | data_mismatch | error | Expected values did not match actual values: expected={'type': 'error', 'value': '#DIV/0!'}, actual={'type': 'blank'} |
| cell_values | polars | error_na | read | data_mismatch | error | Expected values did not match actual values: expected={'type': 'error', 'value': '#N/A'}, actual={'type': 'blank'} |
| cell_values | polars | error_value | read | data_mismatch | error | Expected values did not match actual values: expected={'type': 'error', 'value': '#VALUE!'}, actual={'type': 'blank'} |
| formulas | python-calamine | formula_sum | read | internal | error | RuntimeError: Expected formula, got number |
| formulas | python-calamine | formula_cell_ref | read | internal | error | RuntimeError: Expected formula, got blank |
| formulas | python-calamine | formula_concat | read | internal | error | RuntimeError: Expected formula, got string |
| formulas | python-calamine | formula_cross_sheet | read | internal | error | RuntimeError: Expected formula, got number |
| formulas | pylightxl | formula_sum | read | internal | error | RuntimeError: Expected formula, got number |
| formulas | pylightxl | formula_cell_ref | read | internal | error | RuntimeError: Expected formula, got error |
| formulas | pylightxl | formula_concat | read | internal | error | RuntimeError: Expected formula, got string |
| formulas | pylightxl | formula_cross_sheet | read | internal | error | RuntimeError: Expected formula, got number |
| formulas | pyexcel | formula_sum | read | internal | error | RuntimeError: Expected formula, got number |
| formulas | pyexcel | formula_cell_ref | read | internal | error | RuntimeError: Expected formula, got error |
| formulas | pyexcel | formula_concat | read | internal | error | RuntimeError: Expected formula, got string |
| formulas | pyexcel | formula_cross_sheet | read | internal | error | RuntimeError: Expected formula, got number |
| formulas | xlwt | formula_sum | write | internal | error | RuntimeError: Expected formula, got string |
| formulas | xlwt | formula_cell_ref | write | internal | error | RuntimeError: Expected formula, got string |
| formulas | xlwt | formula_concat | write | internal | error | RuntimeError: Expected formula, got string |
| formulas | xlwt | formula_cross_sheet | write | internal | error | RuntimeError: Expected formula, got string |
| formulas | pandas | formula_sum | read | internal | error | RuntimeError: Expected formula, got number |
| formulas | pandas | formula_cell_ref | read | internal | error | RuntimeError: Expected formula, got blank |
| formulas | pandas | formula_concat | read | internal | error | RuntimeError: Expected formula, got string |
| formulas | pandas | formula_cross_sheet | read | internal | error | RuntimeError: Expected formula, got number |
| formulas | polars | formula_sum | read | internal | error | RuntimeError: Expected formula, got number |
| formulas | polars | formula_cell_ref | read | internal | error | RuntimeError: Expected formula, got blank |
| formulas | polars | formula_concat | read | internal | error | RuntimeError: Expected formula, got string |
| formulas | polars | formula_cross_sheet | read | internal | error | RuntimeError: Expected formula, got number |
| formulas | tablib | formula_sum | read | internal | error | RuntimeError: Expected formula, got number |
| formulas | tablib | formula_cell_ref | read | internal | error | RuntimeError: Expected formula, got error |
| formulas | tablib | formula_concat | read | internal | error | RuntimeError: Expected formula, got string |
| formulas | tablib | formula_cross_sheet | read | internal | error | RuntimeError: Expected formula, got number |
| text_formatting | python-calamine | bold | read | data_mismatch | error | Expected values did not match actual values: expected={'bold': True}, actual={} |
| text_formatting | python-calamine | italic | read | data_mismatch | error | Expected values did not match actual values: expected={'italic': True}, actual={} |
| text_formatting | python-calamine | underline_single | read | data_mismatch | error | Expected values did not match actual values: expected={'underline': 'single'}, actual={} |
| text_formatting | python-calamine | underline_double | read | data_mismatch | error | Expected values did not match actual values: expected={'underline': 'double'}, actual={} |
| text_formatting | python-calamine | strikethrough | read | data_mismatch | error | Expected values did not match actual values: expected={'strikethrough': True}, actual={} |
| text_formatting | python-calamine | bold_italic | read | data_mismatch | error | Expected values did not match actual values: expected={'bold': True, 'italic': True}, actual={} |
| text_formatting | python-calamine | font_size_8 | read | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 8}, actual={} |
| text_formatting | python-calamine | font_size_14 | read | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 14}, actual={} |
| text_formatting | python-calamine | font_size_24 | read | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 24}, actual={} |
| text_formatting | python-calamine | font_size_36 | read | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 36}, actual={} |
| text_formatting | python-calamine | font_arial | read | data_mismatch | error | Expected values did not match actual values: expected={'font_name': 'Arial'}, actual={} |
| text_formatting | python-calamine | font_times | read | data_mismatch | error | Expected values did not match actual values: expected={'font_name': 'Times New Roman'}, actual={} |
| text_formatting | python-calamine | font_courier | read | data_mismatch | error | Expected values did not match actual values: expected={'font_name': 'Courier New'}, actual={} |
| text_formatting | python-calamine | color_red | read | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#FF0000'}, actual={} |
| text_formatting | python-calamine | color_blue | read | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#0000FF'}, actual={} |
| text_formatting | python-calamine | color_green | read | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#00FF00'}, actual={} |
| text_formatting | python-calamine | color_custom | read | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#8B4513'}, actual={} |
| text_formatting | python-calamine | combined | read | data_mismatch | error | Expected values did not match actual values: expected={'bold': True, 'font_size': 16, 'font_color': '#FF0000'}, actual={} |
| text_formatting | pylightxl | bold | read | data_mismatch | error | Expected values did not match actual values: expected={'bold': True}, actual={} |
| text_formatting | pylightxl | italic | read | data_mismatch | error | Expected values did not match actual values: expected={'italic': True}, actual={} |
| text_formatting | pylightxl | underline_single | read | data_mismatch | error | Expected values did not match actual values: expected={'underline': 'single'}, actual={} |
| text_formatting | pylightxl | underline_double | read | data_mismatch | error | Expected values did not match actual values: expected={'underline': 'double'}, actual={} |
| text_formatting | pylightxl | strikethrough | read | data_mismatch | error | Expected values did not match actual values: expected={'strikethrough': True}, actual={} |
| text_formatting | pylightxl | bold_italic | read | data_mismatch | error | Expected values did not match actual values: expected={'bold': True, 'italic': True}, actual={} |
| text_formatting | pylightxl | font_size_8 | read | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 8}, actual={} |
| text_formatting | pylightxl | font_size_14 | read | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 14}, actual={} |
| text_formatting | pylightxl | font_size_24 | read | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 24}, actual={} |
| text_formatting | pylightxl | font_size_36 | read | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 36}, actual={} |
| text_formatting | pylightxl | font_arial | read | data_mismatch | error | Expected values did not match actual values: expected={'font_name': 'Arial'}, actual={} |
| text_formatting | pylightxl | font_times | read | data_mismatch | error | Expected values did not match actual values: expected={'font_name': 'Times New Roman'}, actual={} |
| text_formatting | pylightxl | font_courier | read | data_mismatch | error | Expected values did not match actual values: expected={'font_name': 'Courier New'}, actual={} |
| text_formatting | pylightxl | color_red | read | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#FF0000'}, actual={} |
| text_formatting | pylightxl | color_blue | read | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#0000FF'}, actual={} |
| text_formatting | pylightxl | color_green | read | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#00FF00'}, actual={} |
| text_formatting | pylightxl | color_custom | read | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#8B4513'}, actual={} |
| text_formatting | pylightxl | combined | read | data_mismatch | error | Expected values did not match actual values: expected={'bold': True, 'font_size': 16, 'font_color': '#FF0000'}, actual={} |
| text_formatting | pylightxl | bold | write | data_mismatch | error | Expected values did not match actual values: expected={'bold': True}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pylightxl | italic | write | data_mismatch | error | Expected values did not match actual values: expected={'italic': True}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pylightxl | underline_single | write | data_mismatch | error | Expected values did not match actual values: expected={'underline': 'single'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pylightxl | underline_double | write | data_mismatch | error | Expected values did not match actual values: expected={'underline': 'double'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pylightxl | strikethrough | write | data_mismatch | error | Expected values did not match actual values: expected={'strikethrough': True}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pylightxl | bold_italic | write | data_mismatch | error | Expected values did not match actual values: expected={'bold': True, 'italic': True}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pylightxl | font_size_8 | write | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 8}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pylightxl | font_size_14 | write | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 14}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pylightxl | font_size_24 | write | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 24}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pylightxl | font_size_36 | write | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 36}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pylightxl | font_arial | write | data_mismatch | error | Expected values did not match actual values: expected={'font_name': 'Arial'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pylightxl | font_times | write | data_mismatch | error | Expected values did not match actual values: expected={'font_name': 'Times New Roman'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pylightxl | font_courier | write | data_mismatch | error | Expected values did not match actual values: expected={'font_name': 'Courier New'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pylightxl | color_red | write | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#FF0000'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pylightxl | color_blue | write | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#0000FF'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pylightxl | color_green | write | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#00FF00'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pylightxl | color_custom | write | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#8B4513'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pylightxl | combined | write | data_mismatch | error | Expected values did not match actual values: expected={'bold': True, 'font_size': 16, 'font_color': '#FF0000'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pyexcel | bold | read | data_mismatch | error | Expected values did not match actual values: expected={'bold': True}, actual={} |
| text_formatting | pyexcel | italic | read | data_mismatch | error | Expected values did not match actual values: expected={'italic': True}, actual={} |
| text_formatting | pyexcel | underline_single | read | data_mismatch | error | Expected values did not match actual values: expected={'underline': 'single'}, actual={} |
| text_formatting | pyexcel | underline_double | read | data_mismatch | error | Expected values did not match actual values: expected={'underline': 'double'}, actual={} |
| text_formatting | pyexcel | strikethrough | read | data_mismatch | error | Expected values did not match actual values: expected={'strikethrough': True}, actual={} |
| text_formatting | pyexcel | bold_italic | read | data_mismatch | error | Expected values did not match actual values: expected={'bold': True, 'italic': True}, actual={} |
| text_formatting | pyexcel | font_size_8 | read | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 8}, actual={} |
| text_formatting | pyexcel | font_size_14 | read | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 14}, actual={} |
| text_formatting | pyexcel | font_size_24 | read | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 24}, actual={} |
| text_formatting | pyexcel | font_size_36 | read | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 36}, actual={} |
| text_formatting | pyexcel | font_arial | read | data_mismatch | error | Expected values did not match actual values: expected={'font_name': 'Arial'}, actual={} |
| text_formatting | pyexcel | font_times | read | data_mismatch | error | Expected values did not match actual values: expected={'font_name': 'Times New Roman'}, actual={} |
| text_formatting | pyexcel | font_courier | read | data_mismatch | error | Expected values did not match actual values: expected={'font_name': 'Courier New'}, actual={} |
| text_formatting | pyexcel | color_red | read | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#FF0000'}, actual={} |
| text_formatting | pyexcel | color_blue | read | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#0000FF'}, actual={} |
| text_formatting | pyexcel | color_green | read | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#00FF00'}, actual={} |
| text_formatting | pyexcel | color_custom | read | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#8B4513'}, actual={} |
| text_formatting | pyexcel | combined | read | data_mismatch | error | Expected values did not match actual values: expected={'bold': True, 'font_size': 16, 'font_color': '#FF0000'}, actual={} |
| text_formatting | pyexcel | bold | write | data_mismatch | error | Expected values did not match actual values: expected={'bold': True}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pyexcel | italic | write | data_mismatch | error | Expected values did not match actual values: expected={'italic': True}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pyexcel | underline_single | write | data_mismatch | error | Expected values did not match actual values: expected={'underline': 'single'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pyexcel | underline_double | write | data_mismatch | error | Expected values did not match actual values: expected={'underline': 'double'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pyexcel | strikethrough | write | data_mismatch | error | Expected values did not match actual values: expected={'strikethrough': True}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pyexcel | bold_italic | write | data_mismatch | error | Expected values did not match actual values: expected={'bold': True, 'italic': True}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pyexcel | font_size_8 | write | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 8}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pyexcel | font_size_14 | write | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 14}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pyexcel | font_size_24 | write | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 24}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pyexcel | font_size_36 | write | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 36}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pyexcel | font_arial | write | data_mismatch | error | Expected values did not match actual values: expected={'font_name': 'Arial'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pyexcel | font_times | write | data_mismatch | error | Expected values did not match actual values: expected={'font_name': 'Times New Roman'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pyexcel | font_courier | write | data_mismatch | error | Expected values did not match actual values: expected={'font_name': 'Courier New'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pyexcel | color_red | write | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#FF0000'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pyexcel | color_blue | write | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#0000FF'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pyexcel | color_green | write | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#00FF00'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pyexcel | color_custom | write | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#8B4513'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pyexcel | combined | write | data_mismatch | error | Expected values did not match actual values: expected={'bold': True, 'font_size': 16, 'font_color': '#FF0000'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | xlwt | color_green | write | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#00FF00'}, actual={'font_name': 'Arial', 'font_size': 10.0, 'font_color': '#008000'} |
| text_formatting | xlwt | color_custom | write | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#8B4513'}, actual={'font_name': 'Arial', 'font_size': 10.0, 'font_color': '#993300'} |
| text_formatting | pandas | bold | read | data_mismatch | error | Expected values did not match actual values: expected={'bold': True}, actual={} |
| text_formatting | pandas | italic | read | data_mismatch | error | Expected values did not match actual values: expected={'italic': True}, actual={} |
| text_formatting | pandas | underline_single | read | data_mismatch | error | Expected values did not match actual values: expected={'underline': 'single'}, actual={} |
| text_formatting | pandas | underline_double | read | data_mismatch | error | Expected values did not match actual values: expected={'underline': 'double'}, actual={} |
| text_formatting | pandas | strikethrough | read | data_mismatch | error | Expected values did not match actual values: expected={'strikethrough': True}, actual={} |
| text_formatting | pandas | bold_italic | read | data_mismatch | error | Expected values did not match actual values: expected={'bold': True, 'italic': True}, actual={} |
| text_formatting | pandas | font_size_8 | read | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 8}, actual={} |
| text_formatting | pandas | font_size_14 | read | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 14}, actual={} |
| text_formatting | pandas | font_size_24 | read | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 24}, actual={} |
| text_formatting | pandas | font_size_36 | read | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 36}, actual={} |
| text_formatting | pandas | font_arial | read | data_mismatch | error | Expected values did not match actual values: expected={'font_name': 'Arial'}, actual={} |
| text_formatting | pandas | font_times | read | data_mismatch | error | Expected values did not match actual values: expected={'font_name': 'Times New Roman'}, actual={} |
| text_formatting | pandas | font_courier | read | data_mismatch | error | Expected values did not match actual values: expected={'font_name': 'Courier New'}, actual={} |
| text_formatting | pandas | color_red | read | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#FF0000'}, actual={} |
| text_formatting | pandas | color_blue | read | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#0000FF'}, actual={} |
| text_formatting | pandas | color_green | read | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#00FF00'}, actual={} |
| text_formatting | pandas | color_custom | read | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#8B4513'}, actual={} |
| text_formatting | pandas | combined | read | data_mismatch | error | Expected values did not match actual values: expected={'bold': True, 'font_size': 16, 'font_color': '#FF0000'}, actual={} |
| text_formatting | pandas | bold | write | data_mismatch | error | Expected values did not match actual values: expected={'bold': True}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pandas | italic | write | data_mismatch | error | Expected values did not match actual values: expected={'italic': True}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pandas | underline_single | write | data_mismatch | error | Expected values did not match actual values: expected={'underline': 'single'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pandas | underline_double | write | data_mismatch | error | Expected values did not match actual values: expected={'underline': 'double'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pandas | strikethrough | write | data_mismatch | error | Expected values did not match actual values: expected={'strikethrough': True}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pandas | bold_italic | write | data_mismatch | error | Expected values did not match actual values: expected={'bold': True, 'italic': True}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pandas | font_size_8 | write | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 8}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pandas | font_size_14 | write | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 14}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pandas | font_size_24 | write | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 24}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pandas | font_size_36 | write | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 36}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pandas | font_arial | write | data_mismatch | error | Expected values did not match actual values: expected={'font_name': 'Arial'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pandas | font_times | write | data_mismatch | error | Expected values did not match actual values: expected={'font_name': 'Times New Roman'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pandas | font_courier | write | data_mismatch | error | Expected values did not match actual values: expected={'font_name': 'Courier New'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pandas | color_red | write | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#FF0000'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pandas | color_blue | write | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#0000FF'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pandas | color_green | write | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#00FF00'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pandas | color_custom | write | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#8B4513'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | pandas | combined | write | data_mismatch | error | Expected values did not match actual values: expected={'bold': True, 'font_size': 16, 'font_color': '#FF0000'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | openpyxl-readonly | bold | read | data_mismatch | error | Expected values did not match actual values: expected={'bold': True}, actual={} |
| text_formatting | openpyxl-readonly | italic | read | data_mismatch | error | Expected values did not match actual values: expected={'italic': True}, actual={} |
| text_formatting | openpyxl-readonly | underline_single | read | data_mismatch | error | Expected values did not match actual values: expected={'underline': 'single'}, actual={} |
| text_formatting | openpyxl-readonly | underline_double | read | data_mismatch | error | Expected values did not match actual values: expected={'underline': 'double'}, actual={} |
| text_formatting | openpyxl-readonly | strikethrough | read | data_mismatch | error | Expected values did not match actual values: expected={'strikethrough': True}, actual={} |
| text_formatting | openpyxl-readonly | bold_italic | read | data_mismatch | error | Expected values did not match actual values: expected={'bold': True, 'italic': True}, actual={} |
| text_formatting | openpyxl-readonly | font_size_8 | read | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 8}, actual={} |
| text_formatting | openpyxl-readonly | font_size_14 | read | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 14}, actual={} |
| text_formatting | openpyxl-readonly | font_size_24 | read | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 24}, actual={} |
| text_formatting | openpyxl-readonly | font_size_36 | read | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 36}, actual={} |
| text_formatting | openpyxl-readonly | font_arial | read | data_mismatch | error | Expected values did not match actual values: expected={'font_name': 'Arial'}, actual={} |
| text_formatting | openpyxl-readonly | font_times | read | data_mismatch | error | Expected values did not match actual values: expected={'font_name': 'Times New Roman'}, actual={} |
| text_formatting | openpyxl-readonly | font_courier | read | data_mismatch | error | Expected values did not match actual values: expected={'font_name': 'Courier New'}, actual={} |
| text_formatting | openpyxl-readonly | color_red | read | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#FF0000'}, actual={} |
| text_formatting | openpyxl-readonly | color_blue | read | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#0000FF'}, actual={} |
| text_formatting | openpyxl-readonly | color_green | read | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#00FF00'}, actual={} |
| text_formatting | openpyxl-readonly | color_custom | read | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#8B4513'}, actual={} |
| text_formatting | openpyxl-readonly | combined | read | data_mismatch | error | Expected values did not match actual values: expected={'bold': True, 'font_size': 16, 'font_color': '#FF0000'}, actual={} |
| text_formatting | polars | bold | read | data_mismatch | error | Expected values did not match actual values: expected={'bold': True}, actual={} |
| text_formatting | polars | italic | read | data_mismatch | error | Expected values did not match actual values: expected={'italic': True}, actual={} |
| text_formatting | polars | underline_single | read | data_mismatch | error | Expected values did not match actual values: expected={'underline': 'single'}, actual={} |
| text_formatting | polars | underline_double | read | data_mismatch | error | Expected values did not match actual values: expected={'underline': 'double'}, actual={} |
| text_formatting | polars | strikethrough | read | data_mismatch | error | Expected values did not match actual values: expected={'strikethrough': True}, actual={} |
| text_formatting | polars | bold_italic | read | data_mismatch | error | Expected values did not match actual values: expected={'bold': True, 'italic': True}, actual={} |
| text_formatting | polars | font_size_8 | read | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 8}, actual={} |
| text_formatting | polars | font_size_14 | read | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 14}, actual={} |
| text_formatting | polars | font_size_24 | read | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 24}, actual={} |
| text_formatting | polars | font_size_36 | read | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 36}, actual={} |
| text_formatting | polars | font_arial | read | data_mismatch | error | Expected values did not match actual values: expected={'font_name': 'Arial'}, actual={} |
| text_formatting | polars | font_times | read | data_mismatch | error | Expected values did not match actual values: expected={'font_name': 'Times New Roman'}, actual={} |
| text_formatting | polars | font_courier | read | data_mismatch | error | Expected values did not match actual values: expected={'font_name': 'Courier New'}, actual={} |
| text_formatting | polars | color_red | read | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#FF0000'}, actual={} |
| text_formatting | polars | color_blue | read | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#0000FF'}, actual={} |
| text_formatting | polars | color_green | read | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#00FF00'}, actual={} |
| text_formatting | polars | color_custom | read | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#8B4513'}, actual={} |
| text_formatting | polars | combined | read | data_mismatch | error | Expected values did not match actual values: expected={'bold': True, 'font_size': 16, 'font_color': '#FF0000'}, actual={} |
| text_formatting | tablib | bold | read | data_mismatch | error | Expected values did not match actual values: expected={'bold': True}, actual={} |
| text_formatting | tablib | italic | read | data_mismatch | error | Expected values did not match actual values: expected={'italic': True}, actual={} |
| text_formatting | tablib | underline_single | read | data_mismatch | error | Expected values did not match actual values: expected={'underline': 'single'}, actual={} |
| text_formatting | tablib | underline_double | read | data_mismatch | error | Expected values did not match actual values: expected={'underline': 'double'}, actual={} |
| text_formatting | tablib | strikethrough | read | data_mismatch | error | Expected values did not match actual values: expected={'strikethrough': True}, actual={} |
| text_formatting | tablib | bold_italic | read | data_mismatch | error | Expected values did not match actual values: expected={'bold': True, 'italic': True}, actual={} |
| text_formatting | tablib | font_size_8 | read | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 8}, actual={} |
| text_formatting | tablib | font_size_14 | read | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 14}, actual={} |
| text_formatting | tablib | font_size_24 | read | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 24}, actual={} |
| text_formatting | tablib | font_size_36 | read | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 36}, actual={} |
| text_formatting | tablib | font_arial | read | data_mismatch | error | Expected values did not match actual values: expected={'font_name': 'Arial'}, actual={} |
| text_formatting | tablib | font_times | read | data_mismatch | error | Expected values did not match actual values: expected={'font_name': 'Times New Roman'}, actual={} |
| text_formatting | tablib | font_courier | read | data_mismatch | error | Expected values did not match actual values: expected={'font_name': 'Courier New'}, actual={} |
| text_formatting | tablib | color_red | read | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#FF0000'}, actual={} |
| text_formatting | tablib | color_blue | read | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#0000FF'}, actual={} |
| text_formatting | tablib | color_green | read | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#00FF00'}, actual={} |
| text_formatting | tablib | color_custom | read | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#8B4513'}, actual={} |
| text_formatting | tablib | combined | read | data_mismatch | error | Expected values did not match actual values: expected={'bold': True, 'font_size': 16, 'font_color': '#FF0000'}, actual={} |
| text_formatting | tablib | bold | write | data_mismatch | error | Expected values did not match actual values: expected={'bold': True}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | tablib | italic | write | data_mismatch | error | Expected values did not match actual values: expected={'italic': True}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | tablib | underline_single | write | data_mismatch | error | Expected values did not match actual values: expected={'underline': 'single'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | tablib | underline_double | write | data_mismatch | error | Expected values did not match actual values: expected={'underline': 'double'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | tablib | strikethrough | write | data_mismatch | error | Expected values did not match actual values: expected={'strikethrough': True}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | tablib | bold_italic | write | data_mismatch | error | Expected values did not match actual values: expected={'bold': True, 'italic': True}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | tablib | font_size_8 | write | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 8}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | tablib | font_size_14 | write | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 14}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | tablib | font_size_24 | write | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 24}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | tablib | font_size_36 | write | data_mismatch | error | Expected values did not match actual values: expected={'font_size': 36}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | tablib | font_arial | write | data_mismatch | error | Expected values did not match actual values: expected={'font_name': 'Arial'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | tablib | font_times | write | data_mismatch | error | Expected values did not match actual values: expected={'font_name': 'Times New Roman'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | tablib | font_courier | write | data_mismatch | error | Expected values did not match actual values: expected={'font_name': 'Courier New'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | tablib | color_red | write | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#FF0000'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | tablib | color_blue | write | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#0000FF'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | tablib | color_green | write | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#00FF00'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | tablib | color_custom | write | data_mismatch | error | Expected values did not match actual values: expected={'font_color': '#8B4513'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| text_formatting | tablib | combined | write | data_mismatch | error | Expected values did not match actual values: expected={'bold': True, 'font_size': 16, 'font_color': '#FF0000'}, actual={'font_name': 'Calibri', 'font_size': 11.0} |
| background_colors | python-calamine | bg_red | read | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#FF0000'}, actual={} |
| background_colors | python-calamine | bg_blue | read | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#0000FF'}, actual={} |
| background_colors | python-calamine | bg_green | read | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#00FF00'}, actual={} |
| background_colors | python-calamine | bg_custom | read | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#8B4513'}, actual={} |
| background_colors | pylightxl | bg_red | read | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#FF0000'}, actual={} |
| background_colors | pylightxl | bg_blue | read | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#0000FF'}, actual={} |
| background_colors | pylightxl | bg_green | read | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#00FF00'}, actual={} |
| background_colors | pylightxl | bg_custom | read | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#8B4513'}, actual={} |
| background_colors | pylightxl | bg_red | write | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#FF0000'}, actual={} |
| background_colors | pylightxl | bg_blue | write | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#0000FF'}, actual={} |
| background_colors | pylightxl | bg_green | write | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#00FF00'}, actual={} |
| background_colors | pylightxl | bg_custom | write | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#8B4513'}, actual={} |
| background_colors | pyexcel | bg_red | read | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#FF0000'}, actual={} |
| background_colors | pyexcel | bg_blue | read | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#0000FF'}, actual={} |
| background_colors | pyexcel | bg_green | read | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#00FF00'}, actual={} |
| background_colors | pyexcel | bg_custom | read | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#8B4513'}, actual={} |
| background_colors | pyexcel | bg_red | write | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#FF0000'}, actual={} |
| background_colors | pyexcel | bg_blue | write | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#0000FF'}, actual={} |
| background_colors | pyexcel | bg_green | write | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#00FF00'}, actual={} |
| background_colors | pyexcel | bg_custom | write | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#8B4513'}, actual={} |
| background_colors | xlwt | bg_green | write | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#00FF00'}, actual={'bg_color': '#008000'} |
| background_colors | xlwt | bg_custom | write | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#8B4513'}, actual={'bg_color': '#993300'} |
| background_colors | pandas | bg_red | read | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#FF0000'}, actual={} |
| background_colors | pandas | bg_blue | read | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#0000FF'}, actual={} |
| background_colors | pandas | bg_green | read | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#00FF00'}, actual={} |
| background_colors | pandas | bg_custom | read | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#8B4513'}, actual={} |
| background_colors | pandas | bg_red | write | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#FF0000'}, actual={} |
| background_colors | pandas | bg_blue | write | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#0000FF'}, actual={} |
| background_colors | pandas | bg_green | write | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#00FF00'}, actual={} |
| background_colors | pandas | bg_custom | write | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#8B4513'}, actual={} |
| background_colors | openpyxl-readonly | bg_red | read | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#FF0000'}, actual={} |
| background_colors | openpyxl-readonly | bg_blue | read | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#0000FF'}, actual={} |
| background_colors | openpyxl-readonly | bg_green | read | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#00FF00'}, actual={} |
| background_colors | openpyxl-readonly | bg_custom | read | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#8B4513'}, actual={} |
| background_colors | polars | bg_red | read | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#FF0000'}, actual={} |
| background_colors | polars | bg_blue | read | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#0000FF'}, actual={} |
| background_colors | polars | bg_green | read | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#00FF00'}, actual={} |
| background_colors | polars | bg_custom | read | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#8B4513'}, actual={} |
| background_colors | tablib | bg_red | read | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#FF0000'}, actual={} |
| background_colors | tablib | bg_blue | read | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#0000FF'}, actual={} |
| background_colors | tablib | bg_green | read | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#00FF00'}, actual={} |
| background_colors | tablib | bg_custom | read | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#8B4513'}, actual={} |
| background_colors | tablib | bg_red | write | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#FF0000'}, actual={} |
| background_colors | tablib | bg_blue | write | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#0000FF'}, actual={} |
| background_colors | tablib | bg_green | write | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#00FF00'}, actual={} |
| background_colors | tablib | bg_custom | write | data_mismatch | error | Expected values did not match actual values: expected={'bg_color': '#8B4513'}, actual={} |
| number_formats | python-calamine | numfmt_currency | read | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '$#,##0.00'}, actual={} |
| number_formats | python-calamine | numfmt_percent | read | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '0.00%'}, actual={} |
| number_formats | python-calamine | numfmt_date | read | data_mismatch | error | Expected values did not match actual values: expected={'number_format': 'yyyy-mm-dd'}, actual={} |
| number_formats | python-calamine | numfmt_scientific | read | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '0.00E+00'}, actual={} |
| number_formats | python-calamine | numfmt_custom_text | read | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '"USD" 0.00'}, actual={} |
| number_formats | pylightxl | numfmt_currency | read | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '$#,##0.00'}, actual={} |
| number_formats | pylightxl | numfmt_percent | read | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '0.00%'}, actual={} |
| number_formats | pylightxl | numfmt_date | read | data_mismatch | error | Expected values did not match actual values: expected={'number_format': 'yyyy-mm-dd'}, actual={} |
| number_formats | pylightxl | numfmt_scientific | read | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '0.00E+00'}, actual={} |
| number_formats | pylightxl | numfmt_custom_text | read | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '"USD" 0.00'}, actual={} |
| number_formats | pylightxl | numfmt_currency | write | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '$#,##0.00'}, actual={'number_format': 'General'} |
| number_formats | pylightxl | numfmt_percent | write | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '0.00%'}, actual={'number_format': 'General'} |
| number_formats | pylightxl | numfmt_date | write | data_mismatch | error | Expected values did not match actual values: expected={'number_format': 'yyyy-mm-dd'}, actual={'number_format': 'General'} |
| number_formats | pylightxl | numfmt_scientific | write | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '0.00E+00'}, actual={'number_format': 'General'} |
| number_formats | pylightxl | numfmt_custom_text | write | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '"USD" 0.00'}, actual={'number_format': 'General'} |
| number_formats | pyexcel | numfmt_currency | read | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '$#,##0.00'}, actual={} |
| number_formats | pyexcel | numfmt_percent | read | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '0.00%'}, actual={} |
| number_formats | pyexcel | numfmt_date | read | data_mismatch | error | Expected values did not match actual values: expected={'number_format': 'yyyy-mm-dd'}, actual={} |
| number_formats | pyexcel | numfmt_scientific | read | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '0.00E+00'}, actual={} |
| number_formats | pyexcel | numfmt_custom_text | read | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '"USD" 0.00'}, actual={} |
| number_formats | pyexcel | numfmt_currency | write | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '$#,##0.00'}, actual={'number_format': 'General'} |
| number_formats | pyexcel | numfmt_percent | write | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '0.00%'}, actual={'number_format': 'General'} |
| number_formats | pyexcel | numfmt_date | write | data_mismatch | error | Expected values did not match actual values: expected={'number_format': 'yyyy-mm-dd'}, actual={'number_format': 'General'} |
| number_formats | pyexcel | numfmt_scientific | write | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '0.00E+00'}, actual={'number_format': 'General'} |
| number_formats | pyexcel | numfmt_custom_text | write | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '"USD" 0.00'}, actual={'number_format': 'General'} |
| number_formats | pandas | numfmt_currency | read | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '$#,##0.00'}, actual={} |
| number_formats | pandas | numfmt_percent | read | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '0.00%'}, actual={} |
| number_formats | pandas | numfmt_date | read | data_mismatch | error | Expected values did not match actual values: expected={'number_format': 'yyyy-mm-dd'}, actual={} |
| number_formats | pandas | numfmt_scientific | read | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '0.00E+00'}, actual={} |
| number_formats | pandas | numfmt_custom_text | read | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '"USD" 0.00'}, actual={} |
| number_formats | pandas | numfmt_currency | write | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '$#,##0.00'}, actual={'number_format': 'General'} |
| number_formats | pandas | numfmt_percent | write | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '0.00%'}, actual={'number_format': 'General'} |
| number_formats | pandas | numfmt_date | write | data_mismatch | error | Expected values did not match actual values: expected={'number_format': 'yyyy-mm-dd'}, actual={'number_format': 'General'} |
| number_formats | pandas | numfmt_scientific | write | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '0.00E+00'}, actual={'number_format': 'General'} |
| number_formats | pandas | numfmt_custom_text | write | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '"USD" 0.00'}, actual={'number_format': 'General'} |
| number_formats | openpyxl-readonly | numfmt_currency | read | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '$#,##0.00'}, actual={} |
| number_formats | openpyxl-readonly | numfmt_percent | read | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '0.00%'}, actual={} |
| number_formats | openpyxl-readonly | numfmt_date | read | data_mismatch | error | Expected values did not match actual values: expected={'number_format': 'yyyy-mm-dd'}, actual={} |
| number_formats | openpyxl-readonly | numfmt_scientific | read | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '0.00E+00'}, actual={} |
| number_formats | openpyxl-readonly | numfmt_custom_text | read | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '"USD" 0.00'}, actual={} |
| number_formats | polars | numfmt_currency | read | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '$#,##0.00'}, actual={} |
| number_formats | polars | numfmt_percent | read | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '0.00%'}, actual={} |
| number_formats | polars | numfmt_date | read | data_mismatch | error | Expected values did not match actual values: expected={'number_format': 'yyyy-mm-dd'}, actual={} |
| number_formats | polars | numfmt_scientific | read | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '0.00E+00'}, actual={} |
| number_formats | polars | numfmt_custom_text | read | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '"USD" 0.00'}, actual={} |
| number_formats | tablib | numfmt_currency | read | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '$#,##0.00'}, actual={} |
| number_formats | tablib | numfmt_percent | read | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '0.00%'}, actual={} |
| number_formats | tablib | numfmt_date | read | data_mismatch | error | Expected values did not match actual values: expected={'number_format': 'yyyy-mm-dd'}, actual={} |
| number_formats | tablib | numfmt_scientific | read | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '0.00E+00'}, actual={} |
| number_formats | tablib | numfmt_custom_text | read | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '"USD" 0.00'}, actual={} |
| number_formats | tablib | numfmt_currency | write | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '$#,##0.00'}, actual={'number_format': 'General'} |
| number_formats | tablib | numfmt_percent | write | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '0.00%'}, actual={'number_format': 'General'} |
| number_formats | tablib | numfmt_date | write | data_mismatch | error | Expected values did not match actual values: expected={'number_format': 'yyyy-mm-dd'}, actual={'number_format': 'General'} |
| number_formats | tablib | numfmt_scientific | write | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '0.00E+00'}, actual={'number_format': 'General'} |
| number_formats | tablib | numfmt_custom_text | write | data_mismatch | error | Expected values did not match actual values: expected={'number_format': '"USD" 0.00'}, actual={'number_format': 'General'} |
| alignment | xlsxwriter | v_bottom | write | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'bottom'}, actual={} |
| alignment | python-calamine | h_left | read | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'left'}, actual={} |
| alignment | python-calamine | h_center | read | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'center'}, actual={} |
| alignment | python-calamine | h_right | read | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'right'}, actual={} |
| alignment | python-calamine | v_top | read | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'top'}, actual={} |
| alignment | python-calamine | v_center | read | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'center'}, actual={} |
| alignment | python-calamine | v_bottom | read | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'bottom'}, actual={} |
| alignment | python-calamine | wrap_text | read | data_mismatch | error | Expected values did not match actual values: expected={'wrap': True}, actual={} |
| alignment | python-calamine | rotation_45 | read | data_mismatch | error | Expected values did not match actual values: expected={'rotation': 45}, actual={} |
| alignment | python-calamine | indent_2 | read | data_mismatch | error | Expected values did not match actual values: expected={'indent': 2}, actual={} |
| alignment | rust_xlsxwriter | v_bottom | write | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'bottom'}, actual={} |
| alignment | wolfxl | v_bottom | write | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'bottom'}, actual={} |
| alignment | pylightxl | h_left | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| alignment | pylightxl | h_center | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| alignment | pylightxl | h_right | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| alignment | pylightxl | v_top | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| alignment | pylightxl | v_center | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| alignment | pylightxl | v_bottom | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| alignment | pylightxl | wrap_text | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| alignment | pylightxl | rotation_45 | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| alignment | pylightxl | indent_2 | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| alignment | pylightxl | h_left | write | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'left'}, actual={} |
| alignment | pylightxl | h_center | write | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'center'}, actual={} |
| alignment | pylightxl | h_right | write | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'right'}, actual={} |
| alignment | pylightxl | v_top | write | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'top'}, actual={} |
| alignment | pylightxl | v_center | write | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'center'}, actual={} |
| alignment | pylightxl | v_bottom | write | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'bottom'}, actual={} |
| alignment | pylightxl | wrap_text | write | data_mismatch | error | Expected values did not match actual values: expected={'wrap': True}, actual={} |
| alignment | pylightxl | rotation_45 | write | data_mismatch | error | Expected values did not match actual values: expected={'rotation': 45}, actual={} |
| alignment | pylightxl | indent_2 | write | data_mismatch | error | Expected values did not match actual values: expected={'indent': 2}, actual={} |
| alignment | pyexcel | h_left | read | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'left'}, actual={} |
| alignment | pyexcel | h_center | read | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'center'}, actual={} |
| alignment | pyexcel | h_right | read | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'right'}, actual={} |
| alignment | pyexcel | v_top | read | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'top'}, actual={} |
| alignment | pyexcel | v_center | read | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'center'}, actual={} |
| alignment | pyexcel | v_bottom | read | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'bottom'}, actual={} |
| alignment | pyexcel | wrap_text | read | data_mismatch | error | Expected values did not match actual values: expected={'wrap': True}, actual={} |
| alignment | pyexcel | rotation_45 | read | data_mismatch | error | Expected values did not match actual values: expected={'rotation': 45}, actual={} |
| alignment | pyexcel | indent_2 | read | data_mismatch | error | Expected values did not match actual values: expected={'indent': 2}, actual={} |
| alignment | pyexcel | h_left | write | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'left'}, actual={} |
| alignment | pyexcel | h_center | write | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'center'}, actual={} |
| alignment | pyexcel | h_right | write | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'right'}, actual={} |
| alignment | pyexcel | v_top | write | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'top'}, actual={} |
| alignment | pyexcel | v_center | write | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'center'}, actual={} |
| alignment | pyexcel | v_bottom | write | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'bottom'}, actual={} |
| alignment | pyexcel | wrap_text | write | data_mismatch | error | Expected values did not match actual values: expected={'wrap': True}, actual={} |
| alignment | pyexcel | rotation_45 | write | data_mismatch | error | Expected values did not match actual values: expected={'rotation': 45}, actual={} |
| alignment | pyexcel | indent_2 | write | data_mismatch | error | Expected values did not match actual values: expected={'indent': 2}, actual={} |
| alignment | pandas | h_left | read | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'left'}, actual={} |
| alignment | pandas | h_center | read | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'center'}, actual={} |
| alignment | pandas | h_right | read | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'right'}, actual={} |
| alignment | pandas | v_top | read | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'top'}, actual={} |
| alignment | pandas | v_center | read | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'center'}, actual={} |
| alignment | pandas | v_bottom | read | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'bottom'}, actual={} |
| alignment | pandas | wrap_text | read | data_mismatch | error | Expected values did not match actual values: expected={'wrap': True}, actual={} |
| alignment | pandas | rotation_45 | read | data_mismatch | error | Expected values did not match actual values: expected={'rotation': 45}, actual={} |
| alignment | pandas | indent_2 | read | data_mismatch | error | Expected values did not match actual values: expected={'indent': 2}, actual={} |
| alignment | pandas | h_left | write | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'left'}, actual={} |
| alignment | pandas | h_center | write | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'center'}, actual={} |
| alignment | pandas | h_right | write | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'right'}, actual={} |
| alignment | pandas | v_top | write | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'top'}, actual={} |
| alignment | pandas | v_center | write | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'center'}, actual={} |
| alignment | pandas | v_bottom | write | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'bottom'}, actual={} |
| alignment | pandas | wrap_text | write | data_mismatch | error | Expected values did not match actual values: expected={'wrap': True}, actual={} |
| alignment | pandas | rotation_45 | write | data_mismatch | error | Expected values did not match actual values: expected={'rotation': 45}, actual={} |
| alignment | pandas | indent_2 | write | data_mismatch | error | Expected values did not match actual values: expected={'indent': 2}, actual={} |
| alignment | xlsxwriter-constmem | v_bottom | write | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'bottom'}, actual={} |
| alignment | openpyxl-readonly | h_left | read | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'left'}, actual={} |
| alignment | openpyxl-readonly | h_center | read | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'center'}, actual={} |
| alignment | openpyxl-readonly | h_right | read | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'right'}, actual={} |
| alignment | openpyxl-readonly | v_top | read | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'top'}, actual={} |
| alignment | openpyxl-readonly | v_center | read | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'center'}, actual={} |
| alignment | openpyxl-readonly | v_bottom | read | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'bottom'}, actual={} |
| alignment | openpyxl-readonly | wrap_text | read | data_mismatch | error | Expected values did not match actual values: expected={'wrap': True}, actual={} |
| alignment | openpyxl-readonly | rotation_45 | read | data_mismatch | error | Expected values did not match actual values: expected={'rotation': 45}, actual={} |
| alignment | openpyxl-readonly | indent_2 | read | data_mismatch | error | Expected values did not match actual values: expected={'indent': 2}, actual={} |
| alignment | polars | h_left | read | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'left'}, actual={} |
| alignment | polars | h_center | read | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'center'}, actual={} |
| alignment | polars | h_right | read | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'right'}, actual={} |
| alignment | polars | v_top | read | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'top'}, actual={} |
| alignment | polars | v_center | read | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'center'}, actual={} |
| alignment | polars | v_bottom | read | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'bottom'}, actual={} |
| alignment | polars | wrap_text | read | data_mismatch | error | Expected values did not match actual values: expected={'wrap': True}, actual={} |
| alignment | polars | rotation_45 | read | data_mismatch | error | Expected values did not match actual values: expected={'rotation': 45}, actual={} |
| alignment | polars | indent_2 | read | data_mismatch | error | Expected values did not match actual values: expected={'indent': 2}, actual={} |
| alignment | tablib | h_left | read | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'left'}, actual={} |
| alignment | tablib | h_center | read | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'center'}, actual={} |
| alignment | tablib | h_right | read | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'right'}, actual={} |
| alignment | tablib | v_top | read | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'top'}, actual={} |
| alignment | tablib | v_center | read | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'center'}, actual={} |
| alignment | tablib | v_bottom | read | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'bottom'}, actual={} |
| alignment | tablib | wrap_text | read | data_mismatch | error | Expected values did not match actual values: expected={'wrap': True}, actual={} |
| alignment | tablib | rotation_45 | read | data_mismatch | error | Expected values did not match actual values: expected={'rotation': 45}, actual={} |
| alignment | tablib | indent_2 | read | data_mismatch | error | Expected values did not match actual values: expected={'indent': 2}, actual={} |
| alignment | tablib | h_left | write | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'left'}, actual={} |
| alignment | tablib | h_center | write | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'center'}, actual={} |
| alignment | tablib | h_right | write | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'right'}, actual={} |
| alignment | tablib | v_top | write | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'top'}, actual={} |
| alignment | tablib | v_center | write | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'center'}, actual={} |
| alignment | tablib | v_bottom | write | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'bottom'}, actual={} |
| alignment | tablib | wrap_text | write | data_mismatch | error | Expected values did not match actual values: expected={'wrap': True}, actual={} |
| alignment | tablib | rotation_45 | write | data_mismatch | error | Expected values did not match actual values: expected={'rotation': 45}, actual={} |
| alignment | tablib | indent_2 | write | data_mismatch | error | Expected values did not match actual values: expected={'indent': 2}, actual={} |
| borders | python-calamine | thin_all | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#000000'}, actual={} |
| borders | python-calamine | medium_all | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'medium', 'border_color': '#000000'}, actual={} |
| borders | python-calamine | thick_all | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thick', 'border_color': '#000000'}, actual={} |
| borders | python-calamine | double | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'double', 'border_color': '#000000'}, actual={} |
| borders | python-calamine | dashed | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dashed', 'border_color': '#000000'}, actual={} |
| borders | python-calamine | dotted | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dotted', 'border_color': '#000000'}, actual={} |
| borders | python-calamine | dash_dot | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dashDot', 'border_color': '#000000'}, actual={} |
| borders | python-calamine | dash_dot_dot | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dashDotDot', 'border_color': '#000000'}, actual={} |
| borders | python-calamine | top_only | read | data_mismatch | error | Expected values did not match actual values: expected={'border_top': 'thin', 'border_bottom': None, 'border_left': None, 'border_right': None}, actual={} |
| borders | python-calamine | bottom_only | read | data_mismatch | error | Expected values did not match actual values: expected={'border_top': None, 'border_bottom': 'thin', 'border_left': None, 'border_right': None}, actual={} |
| borders | python-calamine | left_only | read | data_mismatch | error | Expected values did not match actual values: expected={'border_top': None, 'border_bottom': None, 'border_left': 'thin', 'border_right': None}, actual={} |
| borders | python-calamine | right_only | read | data_mismatch | error | Expected values did not match actual values: expected={'border_top': None, 'border_bottom': None, 'border_left': None, 'border_right': 'thin'}, actual={} |
| borders | python-calamine | diagonal_up | read | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_up': 'thin'}, actual={} |
| borders | python-calamine | diagonal_down | read | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_down': 'thin'}, actual={} |
| borders | python-calamine | diagonal_both | read | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_up': 'thin', 'border_diagonal_down': 'thin'}, actual={} |
| borders | python-calamine | color_red | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#FF0000'}, actual={} |
| borders | python-calamine | color_blue | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#0000FF'}, actual={} |
| borders | python-calamine | color_custom | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#8B4513'}, actual={} |
| borders | python-calamine | mixed_styles | read | data_mismatch | error | Expected values did not match actual values: expected={'border_top': 'thick', 'border_bottom': 'thin', 'border_left': 'medium', 'border_right': 'dashed'}, actual={} |
| borders | python-calamine | mixed_colors | read | data_mismatch | error | Expected values did not match actual values: expected={'border_top_color': '#FF0000', 'border_bottom_color': '#00FF00', 'border_left_color': '#0000FF', 'border_right_color': '#FFFF00'}, actual={} |
| borders | pylightxl | thin_all | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| borders | pylightxl | medium_all | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| borders | pylightxl | thick_all | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| borders | pylightxl | double | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| borders | pylightxl | dashed | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| borders | pylightxl | dotted | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| borders | pylightxl | dash_dot | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| borders | pylightxl | dash_dot_dot | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| borders | pylightxl | top_only | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| borders | pylightxl | bottom_only | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| borders | pylightxl | left_only | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| borders | pylightxl | right_only | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| borders | pylightxl | diagonal_up | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| borders | pylightxl | diagonal_down | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| borders | pylightxl | diagonal_both | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| borders | pylightxl | color_red | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| borders | pylightxl | color_blue | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| borders | pylightxl | color_custom | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| borders | pylightxl | mixed_styles | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| borders | pylightxl | mixed_colors | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| borders | pylightxl | thin_all | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#000000'}, actual={} |
| borders | pylightxl | medium_all | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'medium', 'border_color': '#000000'}, actual={} |
| borders | pylightxl | thick_all | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thick', 'border_color': '#000000'}, actual={} |
| borders | pylightxl | double | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'double', 'border_color': '#000000'}, actual={} |
| borders | pylightxl | dashed | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dashed', 'border_color': '#000000'}, actual={} |
| borders | pylightxl | dotted | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dotted', 'border_color': '#000000'}, actual={} |
| borders | pylightxl | dash_dot | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dashDot', 'border_color': '#000000'}, actual={} |
| borders | pylightxl | dash_dot_dot | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dashDotDot', 'border_color': '#000000'}, actual={} |
| borders | pylightxl | top_only | write | data_mismatch | error | Expected values did not match actual values: expected={'border_top': 'thin', 'border_bottom': None, 'border_left': None, 'border_right': None}, actual={} |
| borders | pylightxl | bottom_only | write | data_mismatch | error | Expected values did not match actual values: expected={'border_top': None, 'border_bottom': 'thin', 'border_left': None, 'border_right': None}, actual={} |
| borders | pylightxl | left_only | write | data_mismatch | error | Expected values did not match actual values: expected={'border_top': None, 'border_bottom': None, 'border_left': 'thin', 'border_right': None}, actual={} |
| borders | pylightxl | right_only | write | data_mismatch | error | Expected values did not match actual values: expected={'border_top': None, 'border_bottom': None, 'border_left': None, 'border_right': 'thin'}, actual={} |
| borders | pylightxl | diagonal_up | write | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_up': 'thin'}, actual={} |
| borders | pylightxl | diagonal_down | write | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_down': 'thin'}, actual={} |
| borders | pylightxl | diagonal_both | write | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_up': 'thin', 'border_diagonal_down': 'thin'}, actual={} |
| borders | pylightxl | color_red | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#FF0000'}, actual={} |
| borders | pylightxl | color_blue | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#0000FF'}, actual={} |
| borders | pylightxl | color_custom | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#8B4513'}, actual={} |
| borders | pylightxl | mixed_styles | write | data_mismatch | error | Expected values did not match actual values: expected={'border_top': 'thick', 'border_bottom': 'thin', 'border_left': 'medium', 'border_right': 'dashed'}, actual={} |
| borders | pylightxl | mixed_colors | write | data_mismatch | error | Expected values did not match actual values: expected={'border_top_color': '#FF0000', 'border_bottom_color': '#00FF00', 'border_left_color': '#0000FF', 'border_right_color': '#FFFF00'}, actual={} |
| borders | pyexcel | thin_all | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#000000'}, actual={} |
| borders | pyexcel | medium_all | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'medium', 'border_color': '#000000'}, actual={} |
| borders | pyexcel | thick_all | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thick', 'border_color': '#000000'}, actual={} |
| borders | pyexcel | double | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'double', 'border_color': '#000000'}, actual={} |
| borders | pyexcel | dashed | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dashed', 'border_color': '#000000'}, actual={} |
| borders | pyexcel | dotted | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dotted', 'border_color': '#000000'}, actual={} |
| borders | pyexcel | dash_dot | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dashDot', 'border_color': '#000000'}, actual={} |
| borders | pyexcel | dash_dot_dot | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dashDotDot', 'border_color': '#000000'}, actual={} |
| borders | pyexcel | top_only | read | data_mismatch | error | Expected values did not match actual values: expected={'border_top': 'thin', 'border_bottom': None, 'border_left': None, 'border_right': None}, actual={} |
| borders | pyexcel | bottom_only | read | data_mismatch | error | Expected values did not match actual values: expected={'border_top': None, 'border_bottom': 'thin', 'border_left': None, 'border_right': None}, actual={} |
| borders | pyexcel | left_only | read | data_mismatch | error | Expected values did not match actual values: expected={'border_top': None, 'border_bottom': None, 'border_left': 'thin', 'border_right': None}, actual={} |
| borders | pyexcel | right_only | read | data_mismatch | error | Expected values did not match actual values: expected={'border_top': None, 'border_bottom': None, 'border_left': None, 'border_right': 'thin'}, actual={} |
| borders | pyexcel | diagonal_up | read | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_up': 'thin'}, actual={} |
| borders | pyexcel | diagonal_down | read | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_down': 'thin'}, actual={} |
| borders | pyexcel | diagonal_both | read | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_up': 'thin', 'border_diagonal_down': 'thin'}, actual={} |
| borders | pyexcel | color_red | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#FF0000'}, actual={} |
| borders | pyexcel | color_blue | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#0000FF'}, actual={} |
| borders | pyexcel | color_custom | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#8B4513'}, actual={} |
| borders | pyexcel | mixed_styles | read | data_mismatch | error | Expected values did not match actual values: expected={'border_top': 'thick', 'border_bottom': 'thin', 'border_left': 'medium', 'border_right': 'dashed'}, actual={} |
| borders | pyexcel | mixed_colors | read | data_mismatch | error | Expected values did not match actual values: expected={'border_top_color': '#FF0000', 'border_bottom_color': '#00FF00', 'border_left_color': '#0000FF', 'border_right_color': '#FFFF00'}, actual={} |
| borders | pyexcel | thin_all | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#000000'}, actual={} |
| borders | pyexcel | medium_all | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'medium', 'border_color': '#000000'}, actual={} |
| borders | pyexcel | thick_all | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thick', 'border_color': '#000000'}, actual={} |
| borders | pyexcel | double | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'double', 'border_color': '#000000'}, actual={} |
| borders | pyexcel | dashed | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dashed', 'border_color': '#000000'}, actual={} |
| borders | pyexcel | dotted | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dotted', 'border_color': '#000000'}, actual={} |
| borders | pyexcel | dash_dot | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dashDot', 'border_color': '#000000'}, actual={} |
| borders | pyexcel | dash_dot_dot | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dashDotDot', 'border_color': '#000000'}, actual={} |
| borders | pyexcel | top_only | write | data_mismatch | error | Expected values did not match actual values: expected={'border_top': 'thin', 'border_bottom': None, 'border_left': None, 'border_right': None}, actual={} |
| borders | pyexcel | bottom_only | write | data_mismatch | error | Expected values did not match actual values: expected={'border_top': None, 'border_bottom': 'thin', 'border_left': None, 'border_right': None}, actual={} |
| borders | pyexcel | left_only | write | data_mismatch | error | Expected values did not match actual values: expected={'border_top': None, 'border_bottom': None, 'border_left': 'thin', 'border_right': None}, actual={} |
| borders | pyexcel | right_only | write | data_mismatch | error | Expected values did not match actual values: expected={'border_top': None, 'border_bottom': None, 'border_left': None, 'border_right': 'thin'}, actual={} |
| borders | pyexcel | diagonal_up | write | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_up': 'thin'}, actual={} |
| borders | pyexcel | diagonal_down | write | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_down': 'thin'}, actual={} |
| borders | pyexcel | diagonal_both | write | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_up': 'thin', 'border_diagonal_down': 'thin'}, actual={} |
| borders | pyexcel | color_red | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#FF0000'}, actual={} |
| borders | pyexcel | color_blue | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#0000FF'}, actual={} |
| borders | pyexcel | color_custom | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#8B4513'}, actual={} |
| borders | pyexcel | mixed_styles | write | data_mismatch | error | Expected values did not match actual values: expected={'border_top': 'thick', 'border_bottom': 'thin', 'border_left': 'medium', 'border_right': 'dashed'}, actual={} |
| borders | pyexcel | mixed_colors | write | data_mismatch | error | Expected values did not match actual values: expected={'border_top_color': '#FF0000', 'border_bottom_color': '#00FF00', 'border_left_color': '#0000FF', 'border_right_color': '#FFFF00'}, actual={} |
| borders | xlwt | diagonal_up | write | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_up': 'thin'}, actual={'border_diagonal_down': 'thin'} |
| borders | xlwt | diagonal_down | write | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_down': 'thin'}, actual={'border_diagonal_up': 'thin'} |
| borders | xlwt | color_custom | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#8B4513'}, actual={'border_style': 'thin', 'border_color': '#993300'} |
| borders | xlwt | mixed_colors | write | data_mismatch | error | Expected values did not match actual values: expected={'border_top_color': '#FF0000', 'border_bottom_color': '#00FF00', 'border_left_color': '#0000FF', 'border_right_color': '#FFFF00'}, actual={'border_top_color': '#FF0000', 'border_bottom_color': '#008000', 'border_left_color': '#0000FF', 'border_right_color': '#FFFF00', 'border_style': 'thin'} |
| borders | pandas | thin_all | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#000000'}, actual={} |
| borders | pandas | medium_all | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'medium', 'border_color': '#000000'}, actual={} |
| borders | pandas | thick_all | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thick', 'border_color': '#000000'}, actual={} |
| borders | pandas | double | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'double', 'border_color': '#000000'}, actual={} |
| borders | pandas | dashed | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dashed', 'border_color': '#000000'}, actual={} |
| borders | pandas | dotted | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dotted', 'border_color': '#000000'}, actual={} |
| borders | pandas | dash_dot | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dashDot', 'border_color': '#000000'}, actual={} |
| borders | pandas | dash_dot_dot | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dashDotDot', 'border_color': '#000000'}, actual={} |
| borders | pandas | top_only | read | data_mismatch | error | Expected values did not match actual values: expected={'border_top': 'thin', 'border_bottom': None, 'border_left': None, 'border_right': None}, actual={} |
| borders | pandas | bottom_only | read | data_mismatch | error | Expected values did not match actual values: expected={'border_top': None, 'border_bottom': 'thin', 'border_left': None, 'border_right': None}, actual={} |
| borders | pandas | left_only | read | data_mismatch | error | Expected values did not match actual values: expected={'border_top': None, 'border_bottom': None, 'border_left': 'thin', 'border_right': None}, actual={} |
| borders | pandas | right_only | read | data_mismatch | error | Expected values did not match actual values: expected={'border_top': None, 'border_bottom': None, 'border_left': None, 'border_right': 'thin'}, actual={} |
| borders | pandas | diagonal_up | read | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_up': 'thin'}, actual={} |
| borders | pandas | diagonal_down | read | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_down': 'thin'}, actual={} |
| borders | pandas | diagonal_both | read | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_up': 'thin', 'border_diagonal_down': 'thin'}, actual={} |
| borders | pandas | color_red | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#FF0000'}, actual={} |
| borders | pandas | color_blue | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#0000FF'}, actual={} |
| borders | pandas | color_custom | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#8B4513'}, actual={} |
| borders | pandas | mixed_styles | read | data_mismatch | error | Expected values did not match actual values: expected={'border_top': 'thick', 'border_bottom': 'thin', 'border_left': 'medium', 'border_right': 'dashed'}, actual={} |
| borders | pandas | mixed_colors | read | data_mismatch | error | Expected values did not match actual values: expected={'border_top_color': '#FF0000', 'border_bottom_color': '#00FF00', 'border_left_color': '#0000FF', 'border_right_color': '#FFFF00'}, actual={} |
| borders | pandas | thin_all | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#000000'}, actual={} |
| borders | pandas | medium_all | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'medium', 'border_color': '#000000'}, actual={} |
| borders | pandas | thick_all | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thick', 'border_color': '#000000'}, actual={} |
| borders | pandas | double | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'double', 'border_color': '#000000'}, actual={} |
| borders | pandas | dashed | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dashed', 'border_color': '#000000'}, actual={} |
| borders | pandas | dotted | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dotted', 'border_color': '#000000'}, actual={} |
| borders | pandas | dash_dot | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dashDot', 'border_color': '#000000'}, actual={} |
| borders | pandas | dash_dot_dot | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dashDotDot', 'border_color': '#000000'}, actual={} |
| borders | pandas | top_only | write | data_mismatch | error | Expected values did not match actual values: expected={'border_top': 'thin', 'border_bottom': None, 'border_left': None, 'border_right': None}, actual={} |
| borders | pandas | bottom_only | write | data_mismatch | error | Expected values did not match actual values: expected={'border_top': None, 'border_bottom': 'thin', 'border_left': None, 'border_right': None}, actual={} |
| borders | pandas | left_only | write | data_mismatch | error | Expected values did not match actual values: expected={'border_top': None, 'border_bottom': None, 'border_left': 'thin', 'border_right': None}, actual={} |
| borders | pandas | right_only | write | data_mismatch | error | Expected values did not match actual values: expected={'border_top': None, 'border_bottom': None, 'border_left': None, 'border_right': 'thin'}, actual={} |
| borders | pandas | diagonal_up | write | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_up': 'thin'}, actual={} |
| borders | pandas | diagonal_down | write | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_down': 'thin'}, actual={} |
| borders | pandas | diagonal_both | write | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_up': 'thin', 'border_diagonal_down': 'thin'}, actual={} |
| borders | pandas | color_red | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#FF0000'}, actual={} |
| borders | pandas | color_blue | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#0000FF'}, actual={} |
| borders | pandas | color_custom | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#8B4513'}, actual={} |
| borders | pandas | mixed_styles | write | data_mismatch | error | Expected values did not match actual values: expected={'border_top': 'thick', 'border_bottom': 'thin', 'border_left': 'medium', 'border_right': 'dashed'}, actual={} |
| borders | pandas | mixed_colors | write | data_mismatch | error | Expected values did not match actual values: expected={'border_top_color': '#FF0000', 'border_bottom_color': '#00FF00', 'border_left_color': '#0000FF', 'border_right_color': '#FFFF00'}, actual={} |
| borders | openpyxl-readonly | thin_all | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#000000'}, actual={} |
| borders | openpyxl-readonly | medium_all | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'medium', 'border_color': '#000000'}, actual={} |
| borders | openpyxl-readonly | thick_all | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thick', 'border_color': '#000000'}, actual={} |
| borders | openpyxl-readonly | double | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'double', 'border_color': '#000000'}, actual={} |
| borders | openpyxl-readonly | dashed | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dashed', 'border_color': '#000000'}, actual={} |
| borders | openpyxl-readonly | dotted | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dotted', 'border_color': '#000000'}, actual={} |
| borders | openpyxl-readonly | dash_dot | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dashDot', 'border_color': '#000000'}, actual={} |
| borders | openpyxl-readonly | dash_dot_dot | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dashDotDot', 'border_color': '#000000'}, actual={} |
| borders | openpyxl-readonly | top_only | read | data_mismatch | error | Expected values did not match actual values: expected={'border_top': 'thin', 'border_bottom': None, 'border_left': None, 'border_right': None}, actual={} |
| borders | openpyxl-readonly | bottom_only | read | data_mismatch | error | Expected values did not match actual values: expected={'border_top': None, 'border_bottom': 'thin', 'border_left': None, 'border_right': None}, actual={} |
| borders | openpyxl-readonly | left_only | read | data_mismatch | error | Expected values did not match actual values: expected={'border_top': None, 'border_bottom': None, 'border_left': 'thin', 'border_right': None}, actual={} |
| borders | openpyxl-readonly | right_only | read | data_mismatch | error | Expected values did not match actual values: expected={'border_top': None, 'border_bottom': None, 'border_left': None, 'border_right': 'thin'}, actual={} |
| borders | openpyxl-readonly | diagonal_up | read | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_up': 'thin'}, actual={} |
| borders | openpyxl-readonly | diagonal_down | read | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_down': 'thin'}, actual={} |
| borders | openpyxl-readonly | diagonal_both | read | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_up': 'thin', 'border_diagonal_down': 'thin'}, actual={} |
| borders | openpyxl-readonly | color_red | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#FF0000'}, actual={} |
| borders | openpyxl-readonly | color_blue | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#0000FF'}, actual={} |
| borders | openpyxl-readonly | color_custom | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#8B4513'}, actual={} |
| borders | openpyxl-readonly | mixed_styles | read | data_mismatch | error | Expected values did not match actual values: expected={'border_top': 'thick', 'border_bottom': 'thin', 'border_left': 'medium', 'border_right': 'dashed'}, actual={} |
| borders | openpyxl-readonly | mixed_colors | read | data_mismatch | error | Expected values did not match actual values: expected={'border_top_color': '#FF0000', 'border_bottom_color': '#00FF00', 'border_left_color': '#0000FF', 'border_right_color': '#FFFF00'}, actual={} |
| borders | polars | thin_all | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#000000'}, actual={} |
| borders | polars | medium_all | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'medium', 'border_color': '#000000'}, actual={} |
| borders | polars | thick_all | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thick', 'border_color': '#000000'}, actual={} |
| borders | polars | double | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'double', 'border_color': '#000000'}, actual={} |
| borders | polars | dashed | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dashed', 'border_color': '#000000'}, actual={} |
| borders | polars | dotted | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dotted', 'border_color': '#000000'}, actual={} |
| borders | polars | dash_dot | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dashDot', 'border_color': '#000000'}, actual={} |
| borders | polars | dash_dot_dot | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dashDotDot', 'border_color': '#000000'}, actual={} |
| borders | polars | top_only | read | data_mismatch | error | Expected values did not match actual values: expected={'border_top': 'thin', 'border_bottom': None, 'border_left': None, 'border_right': None}, actual={} |
| borders | polars | bottom_only | read | data_mismatch | error | Expected values did not match actual values: expected={'border_top': None, 'border_bottom': 'thin', 'border_left': None, 'border_right': None}, actual={} |
| borders | polars | left_only | read | data_mismatch | error | Expected values did not match actual values: expected={'border_top': None, 'border_bottom': None, 'border_left': 'thin', 'border_right': None}, actual={} |
| borders | polars | right_only | read | data_mismatch | error | Expected values did not match actual values: expected={'border_top': None, 'border_bottom': None, 'border_left': None, 'border_right': 'thin'}, actual={} |
| borders | polars | diagonal_up | read | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_up': 'thin'}, actual={} |
| borders | polars | diagonal_down | read | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_down': 'thin'}, actual={} |
| borders | polars | diagonal_both | read | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_up': 'thin', 'border_diagonal_down': 'thin'}, actual={} |
| borders | polars | color_red | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#FF0000'}, actual={} |
| borders | polars | color_blue | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#0000FF'}, actual={} |
| borders | polars | color_custom | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#8B4513'}, actual={} |
| borders | polars | mixed_styles | read | data_mismatch | error | Expected values did not match actual values: expected={'border_top': 'thick', 'border_bottom': 'thin', 'border_left': 'medium', 'border_right': 'dashed'}, actual={} |
| borders | polars | mixed_colors | read | data_mismatch | error | Expected values did not match actual values: expected={'border_top_color': '#FF0000', 'border_bottom_color': '#00FF00', 'border_left_color': '#0000FF', 'border_right_color': '#FFFF00'}, actual={} |
| borders | tablib | thin_all | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#000000'}, actual={} |
| borders | tablib | medium_all | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'medium', 'border_color': '#000000'}, actual={} |
| borders | tablib | thick_all | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thick', 'border_color': '#000000'}, actual={} |
| borders | tablib | double | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'double', 'border_color': '#000000'}, actual={} |
| borders | tablib | dashed | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dashed', 'border_color': '#000000'}, actual={} |
| borders | tablib | dotted | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dotted', 'border_color': '#000000'}, actual={} |
| borders | tablib | dash_dot | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dashDot', 'border_color': '#000000'}, actual={} |
| borders | tablib | dash_dot_dot | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dashDotDot', 'border_color': '#000000'}, actual={} |
| borders | tablib | top_only | read | data_mismatch | error | Expected values did not match actual values: expected={'border_top': 'thin', 'border_bottom': None, 'border_left': None, 'border_right': None}, actual={} |
| borders | tablib | bottom_only | read | data_mismatch | error | Expected values did not match actual values: expected={'border_top': None, 'border_bottom': 'thin', 'border_left': None, 'border_right': None}, actual={} |
| borders | tablib | left_only | read | data_mismatch | error | Expected values did not match actual values: expected={'border_top': None, 'border_bottom': None, 'border_left': 'thin', 'border_right': None}, actual={} |
| borders | tablib | right_only | read | data_mismatch | error | Expected values did not match actual values: expected={'border_top': None, 'border_bottom': None, 'border_left': None, 'border_right': 'thin'}, actual={} |
| borders | tablib | diagonal_up | read | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_up': 'thin'}, actual={} |
| borders | tablib | diagonal_down | read | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_down': 'thin'}, actual={} |
| borders | tablib | diagonal_both | read | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_up': 'thin', 'border_diagonal_down': 'thin'}, actual={} |
| borders | tablib | color_red | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#FF0000'}, actual={} |
| borders | tablib | color_blue | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#0000FF'}, actual={} |
| borders | tablib | color_custom | read | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#8B4513'}, actual={} |
| borders | tablib | mixed_styles | read | data_mismatch | error | Expected values did not match actual values: expected={'border_top': 'thick', 'border_bottom': 'thin', 'border_left': 'medium', 'border_right': 'dashed'}, actual={} |
| borders | tablib | mixed_colors | read | data_mismatch | error | Expected values did not match actual values: expected={'border_top_color': '#FF0000', 'border_bottom_color': '#00FF00', 'border_left_color': '#0000FF', 'border_right_color': '#FFFF00'}, actual={} |
| borders | tablib | thin_all | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#000000'}, actual={} |
| borders | tablib | medium_all | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'medium', 'border_color': '#000000'}, actual={} |
| borders | tablib | thick_all | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thick', 'border_color': '#000000'}, actual={} |
| borders | tablib | double | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'double', 'border_color': '#000000'}, actual={} |
| borders | tablib | dashed | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dashed', 'border_color': '#000000'}, actual={} |
| borders | tablib | dotted | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dotted', 'border_color': '#000000'}, actual={} |
| borders | tablib | dash_dot | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dashDot', 'border_color': '#000000'}, actual={} |
| borders | tablib | dash_dot_dot | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'dashDotDot', 'border_color': '#000000'}, actual={} |
| borders | tablib | top_only | write | data_mismatch | error | Expected values did not match actual values: expected={'border_top': 'thin', 'border_bottom': None, 'border_left': None, 'border_right': None}, actual={} |
| borders | tablib | bottom_only | write | data_mismatch | error | Expected values did not match actual values: expected={'border_top': None, 'border_bottom': 'thin', 'border_left': None, 'border_right': None}, actual={} |
| borders | tablib | left_only | write | data_mismatch | error | Expected values did not match actual values: expected={'border_top': None, 'border_bottom': None, 'border_left': 'thin', 'border_right': None}, actual={} |
| borders | tablib | right_only | write | data_mismatch | error | Expected values did not match actual values: expected={'border_top': None, 'border_bottom': None, 'border_left': None, 'border_right': 'thin'}, actual={} |
| borders | tablib | diagonal_up | write | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_up': 'thin'}, actual={} |
| borders | tablib | diagonal_down | write | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_down': 'thin'}, actual={} |
| borders | tablib | diagonal_both | write | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_up': 'thin', 'border_diagonal_down': 'thin'}, actual={} |
| borders | tablib | color_red | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#FF0000'}, actual={} |
| borders | tablib | color_blue | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#0000FF'}, actual={} |
| borders | tablib | color_custom | write | data_mismatch | error | Expected values did not match actual values: expected={'border_style': 'thin', 'border_color': '#8B4513'}, actual={} |
| borders | tablib | mixed_styles | write | data_mismatch | error | Expected values did not match actual values: expected={'border_top': 'thick', 'border_bottom': 'thin', 'border_left': 'medium', 'border_right': 'dashed'}, actual={} |
| borders | tablib | mixed_colors | write | data_mismatch | error | Expected values did not match actual values: expected={'border_top_color': '#FF0000', 'border_bottom_color': '#00FF00', 'border_left_color': '#0000FF', 'border_right_color': '#FFFF00'}, actual={} |
| dimensions | python-calamine | row_height_30 | read | data_mismatch | error | Expected values did not match actual values: expected={'row_height': 30}, actual={'row_height': None} |
| dimensions | python-calamine | row_height_45 | read | data_mismatch | error | Expected values did not match actual values: expected={'row_height': 45}, actual={'row_height': None} |
| dimensions | python-calamine | col_width_20 | read | data_mismatch | error | Expected values did not match actual values: expected={'column_width': 20}, actual={'column_width': None} |
| dimensions | python-calamine | col_width_8 | read | data_mismatch | error | Expected values did not match actual values: expected={'column_width': 8}, actual={'column_width': None} |
| dimensions | pylightxl | row_height_30 | read | data_mismatch | error | Expected values did not match actual values: expected={'row_height': 30}, actual={'row_height': None} |
| dimensions | pylightxl | row_height_45 | read | data_mismatch | error | Expected values did not match actual values: expected={'row_height': 45}, actual={'row_height': None} |
| dimensions | pylightxl | col_width_20 | read | data_mismatch | error | Expected values did not match actual values: expected={'column_width': 20}, actual={'column_width': None} |
| dimensions | pylightxl | col_width_8 | read | data_mismatch | error | Expected values did not match actual values: expected={'column_width': 8}, actual={'column_width': None} |
| dimensions | pylightxl | row_height_30 | write | data_mismatch | error | Expected values did not match actual values: expected={'row_height': 30}, actual={'row_height': None} |
| dimensions | pylightxl | row_height_45 | write | data_mismatch | error | Expected values did not match actual values: expected={'row_height': 45}, actual={'row_height': None} |
| dimensions | pylightxl | col_width_20 | write | data_mismatch | error | Expected values did not match actual values: expected={'column_width': 20}, actual={'column_width': 13.0} |
| dimensions | pylightxl | col_width_8 | write | data_mismatch | error | Expected values did not match actual values: expected={'column_width': 8}, actual={'column_width': 13.0} |
| dimensions | pyexcel | row_height_30 | read | data_mismatch | error | Expected values did not match actual values: expected={'row_height': 30}, actual={'row_height': None} |
| dimensions | pyexcel | row_height_45 | read | data_mismatch | error | Expected values did not match actual values: expected={'row_height': 45}, actual={'row_height': None} |
| dimensions | pyexcel | col_width_20 | read | data_mismatch | error | Expected values did not match actual values: expected={'column_width': 20}, actual={'column_width': None} |
| dimensions | pyexcel | col_width_8 | read | data_mismatch | error | Expected values did not match actual values: expected={'column_width': 8}, actual={'column_width': None} |
| dimensions | pyexcel | row_height_30 | write | data_mismatch | error | Expected values did not match actual values: expected={'row_height': 30}, actual={'row_height': None} |
| dimensions | pyexcel | row_height_45 | write | data_mismatch | error | Expected values did not match actual values: expected={'row_height': 45}, actual={'row_height': None} |
| dimensions | pyexcel | col_width_20 | write | data_mismatch | error | Expected values did not match actual values: expected={'column_width': 20}, actual={'column_width': 13.0} |
| dimensions | pyexcel | col_width_8 | write | data_mismatch | error | Expected values did not match actual values: expected={'column_width': 8}, actual={'column_width': 13.0} |
| dimensions | xlwt | row_height_30 | write | data_mismatch | error | Expected values did not match actual values: expected={'row_height': 30}, actual={'row_height': None} |
| dimensions | xlwt | row_height_45 | write | data_mismatch | error | Expected values did not match actual values: expected={'row_height': 45}, actual={'row_height': None} |
| dimensions | pandas | row_height_30 | read | data_mismatch | error | Expected values did not match actual values: expected={'row_height': 30}, actual={'row_height': None} |
| dimensions | pandas | row_height_45 | read | data_mismatch | error | Expected values did not match actual values: expected={'row_height': 45}, actual={'row_height': None} |
| dimensions | pandas | col_width_20 | read | data_mismatch | error | Expected values did not match actual values: expected={'column_width': 20}, actual={'column_width': None} |
| dimensions | pandas | col_width_8 | read | data_mismatch | error | Expected values did not match actual values: expected={'column_width': 8}, actual={'column_width': None} |
| dimensions | pandas | row_height_30 | write | data_mismatch | error | Expected values did not match actual values: expected={'row_height': 30}, actual={'row_height': None} |
| dimensions | pandas | row_height_45 | write | data_mismatch | error | Expected values did not match actual values: expected={'row_height': 45}, actual={'row_height': None} |
| dimensions | pandas | col_width_20 | write | data_mismatch | error | Expected values did not match actual values: expected={'column_width': 20}, actual={'column_width': 13.0} |
| dimensions | pandas | col_width_8 | write | data_mismatch | error | Expected values did not match actual values: expected={'column_width': 8}, actual={'column_width': 13.0} |
| dimensions | xlsxwriter-constmem | row_height_30 | write | data_mismatch | error | Expected values did not match actual values: expected={'row_height': 30}, actual={'row_height': None} |
| dimensions | xlsxwriter-constmem | row_height_45 | write | data_mismatch | error | Expected values did not match actual values: expected={'row_height': 45}, actual={'row_height': None} |
| dimensions | openpyxl-readonly | row_height_30 | read | data_mismatch | error | Expected values did not match actual values: expected={'row_height': 30}, actual={'row_height': None} |
| dimensions | openpyxl-readonly | row_height_45 | read | data_mismatch | error | Expected values did not match actual values: expected={'row_height': 45}, actual={'row_height': None} |
| dimensions | openpyxl-readonly | col_width_20 | read | data_mismatch | error | Expected values did not match actual values: expected={'column_width': 20}, actual={'column_width': None} |
| dimensions | openpyxl-readonly | col_width_8 | read | data_mismatch | error | Expected values did not match actual values: expected={'column_width': 8}, actual={'column_width': None} |
| dimensions | polars | row_height_30 | read | data_mismatch | error | Expected values did not match actual values: expected={'row_height': 30}, actual={'row_height': None} |
| dimensions | polars | row_height_45 | read | data_mismatch | error | Expected values did not match actual values: expected={'row_height': 45}, actual={'row_height': None} |
| dimensions | polars | col_width_20 | read | data_mismatch | error | Expected values did not match actual values: expected={'column_width': 20}, actual={'column_width': None} |
| dimensions | polars | col_width_8 | read | data_mismatch | error | Expected values did not match actual values: expected={'column_width': 8}, actual={'column_width': None} |
| dimensions | tablib | row_height_30 | read | data_mismatch | error | Expected values did not match actual values: expected={'row_height': 30}, actual={'row_height': None} |
| dimensions | tablib | row_height_45 | read | data_mismatch | error | Expected values did not match actual values: expected={'row_height': 45}, actual={'row_height': None} |
| dimensions | tablib | col_width_20 | read | data_mismatch | error | Expected values did not match actual values: expected={'column_width': 20}, actual={'column_width': None} |
| dimensions | tablib | col_width_8 | read | data_mismatch | error | Expected values did not match actual values: expected={'column_width': 8}, actual={'column_width': None} |
| dimensions | tablib | row_height_30 | write | data_mismatch | error | Expected values did not match actual values: expected={'row_height': 30}, actual={'row_height': None} |
| dimensions | tablib | row_height_45 | write | data_mismatch | error | Expected values did not match actual values: expected={'row_height': 45}, actual={'row_height': None} |
| dimensions | tablib | col_width_20 | write | data_mismatch | error | Expected values did not match actual values: expected={'column_width': 20}, actual={'column_width': 13.0} |
| dimensions | tablib | col_width_8 | write | data_mismatch | error | Expected values did not match actual values: expected={'column_width': 8}, actual={'column_width': 13.0} |
| multiple_sheets | polars | value_beta | read | data_mismatch | error | Expected values did not match actual values: expected={'type': 'string', 'value': 'Beta'}, actual={'type': 'blank'} |
| multiple_sheets | polars | value_gamma | read | data_mismatch | error | Expected values did not match actual values: expected={'type': 'string', 'value': 'Gamma'}, actual={'type': 'blank'} |
| merged_cells | python-calamine | merge_horizontal | read | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B2:D2', 'top_left_value': 'Merged'}, actual={'merged_range': None} |
| merged_cells | python-calamine | merge_vertical | read | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B3:B5', 'top_left_value': 'Vertical'}, actual={'merged_range': None} |
| merged_cells | python-calamine | merge_value_off_top_left | read | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B6:D6', 'top_left_value': 'OffTop', 'non_top_left_nonempty': 0}, actual={'merged_range': None} |
| merged_cells | python-calamine | merge_top_left_fill | read | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B7:D7', 'top_left_value': 'Fill', 'top_left_bg_color': '#FF0000'}, actual={'merged_range': None} |
| merged_cells | pylightxl | merge_horizontal | read | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B2:D2', 'top_left_value': 'Merged'}, actual={'merged_range': None} |
| merged_cells | pylightxl | merge_vertical | read | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B3:B5', 'top_left_value': 'Vertical'}, actual={'merged_range': None} |
| merged_cells | pylightxl | merge_value_off_top_left | read | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B6:D6', 'top_left_value': 'OffTop', 'non_top_left_nonempty': 0}, actual={'merged_range': None} |
| merged_cells | pylightxl | merge_top_left_fill | read | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B7:D7', 'top_left_value': 'Fill', 'top_left_bg_color': '#FF0000'}, actual={'merged_range': None} |
| merged_cells | pylightxl | merge_horizontal | write | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B2:D2', 'top_left_value': 'Merged'}, actual={'merged_range': None} |
| merged_cells | pylightxl | merge_vertical | write | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B3:B5', 'top_left_value': 'Vertical'}, actual={'merged_range': None} |
| merged_cells | pylightxl | merge_value_off_top_left | write | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B6:D6', 'top_left_value': 'OffTop', 'non_top_left_nonempty': 0}, actual={'merged_range': None} |
| merged_cells | pylightxl | merge_top_left_fill | write | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B7:D7', 'top_left_value': 'Fill', 'top_left_bg_color': '#FF0000'}, actual={'merged_range': None} |
| merged_cells | pyexcel | merge_horizontal | read | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B2:D2', 'top_left_value': 'Merged'}, actual={'merged_range': None} |
| merged_cells | pyexcel | merge_vertical | read | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B3:B5', 'top_left_value': 'Vertical'}, actual={'merged_range': None} |
| merged_cells | pyexcel | merge_value_off_top_left | read | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B6:D6', 'top_left_value': 'OffTop', 'non_top_left_nonempty': 0}, actual={'merged_range': None} |
| merged_cells | pyexcel | merge_top_left_fill | read | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B7:D7', 'top_left_value': 'Fill', 'top_left_bg_color': '#FF0000'}, actual={'merged_range': None} |
| merged_cells | pyexcel | merge_horizontal | write | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B2:D2', 'top_left_value': 'Merged'}, actual={'merged_range': None} |
| merged_cells | pyexcel | merge_vertical | write | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B3:B5', 'top_left_value': 'Vertical'}, actual={'merged_range': None} |
| merged_cells | pyexcel | merge_value_off_top_left | write | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B6:D6', 'top_left_value': 'OffTop', 'non_top_left_nonempty': 0}, actual={'merged_range': None} |
| merged_cells | pyexcel | merge_top_left_fill | write | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B7:D7', 'top_left_value': 'Fill', 'top_left_bg_color': '#FF0000'}, actual={'merged_range': None} |
| merged_cells | xlwt | merge_horizontal | write | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B2:D2', 'top_left_value': 'Merged'}, actual={'merged_range': 'B2:D2', 'top_left_value': None} |
| merged_cells | xlwt | merge_vertical | write | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B3:B5', 'top_left_value': 'Vertical'}, actual={'merged_range': 'B3:B5', 'top_left_value': None} |
| merged_cells | xlwt | merge_value_off_top_left | write | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B6:D6', 'top_left_value': 'OffTop', 'non_top_left_nonempty': 0}, actual={'merged_range': 'B6:D6', 'top_left_value': None, 'non_top_left_nonempty': 0} |
| merged_cells | xlwt | merge_top_left_fill | write | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B7:D7', 'top_left_value': 'Fill', 'top_left_bg_color': '#FF0000'}, actual={'merged_range': 'B7:D7', 'top_left_value': None, 'top_left_bg_color': None} |
| merged_cells | pandas | merge_horizontal | read | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B2:D2', 'top_left_value': 'Merged'}, actual={'merged_range': None} |
| merged_cells | pandas | merge_vertical | read | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B3:B5', 'top_left_value': 'Vertical'}, actual={'merged_range': None} |
| merged_cells | pandas | merge_value_off_top_left | read | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B6:D6', 'top_left_value': 'OffTop', 'non_top_left_nonempty': 0}, actual={'merged_range': None} |
| merged_cells | pandas | merge_top_left_fill | read | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B7:D7', 'top_left_value': 'Fill', 'top_left_bg_color': '#FF0000'}, actual={'merged_range': None} |
| merged_cells | pandas | merge_horizontal | write | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B2:D2', 'top_left_value': 'Merged'}, actual={'merged_range': None} |
| merged_cells | pandas | merge_vertical | write | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B3:B5', 'top_left_value': 'Vertical'}, actual={'merged_range': None} |
| merged_cells | pandas | merge_value_off_top_left | write | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B6:D6', 'top_left_value': 'OffTop', 'non_top_left_nonempty': 0}, actual={'merged_range': None} |
| merged_cells | pandas | merge_top_left_fill | write | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B7:D7', 'top_left_value': 'Fill', 'top_left_bg_color': '#FF0000'}, actual={'merged_range': None} |
| merged_cells | openpyxl-readonly | merge_horizontal | read | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B2:D2', 'top_left_value': 'Merged'}, actual={'merged_range': None} |
| merged_cells | openpyxl-readonly | merge_vertical | read | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B3:B5', 'top_left_value': 'Vertical'}, actual={'merged_range': None} |
| merged_cells | openpyxl-readonly | merge_value_off_top_left | read | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B6:D6', 'top_left_value': 'OffTop', 'non_top_left_nonempty': 0}, actual={'merged_range': None} |
| merged_cells | openpyxl-readonly | merge_top_left_fill | read | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B7:D7', 'top_left_value': 'Fill', 'top_left_bg_color': '#FF0000'}, actual={'merged_range': None} |
| merged_cells | polars | merge_horizontal | read | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B2:D2', 'top_left_value': 'Merged'}, actual={'merged_range': None} |
| merged_cells | polars | merge_vertical | read | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B3:B5', 'top_left_value': 'Vertical'}, actual={'merged_range': None} |
| merged_cells | polars | merge_value_off_top_left | read | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B6:D6', 'top_left_value': 'OffTop', 'non_top_left_nonempty': 0}, actual={'merged_range': None} |
| merged_cells | polars | merge_top_left_fill | read | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B7:D7', 'top_left_value': 'Fill', 'top_left_bg_color': '#FF0000'}, actual={'merged_range': None} |
| merged_cells | tablib | merge_horizontal | read | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B2:D2', 'top_left_value': 'Merged'}, actual={'merged_range': None} |
| merged_cells | tablib | merge_vertical | read | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B3:B5', 'top_left_value': 'Vertical'}, actual={'merged_range': None} |
| merged_cells | tablib | merge_value_off_top_left | read | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B6:D6', 'top_left_value': 'OffTop', 'non_top_left_nonempty': 0}, actual={'merged_range': None} |
| merged_cells | tablib | merge_top_left_fill | read | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B7:D7', 'top_left_value': 'Fill', 'top_left_bg_color': '#FF0000'}, actual={'merged_range': None} |
| merged_cells | tablib | merge_horizontal | write | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B2:D2', 'top_left_value': 'Merged'}, actual={'merged_range': None} |
| merged_cells | tablib | merge_vertical | write | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B3:B5', 'top_left_value': 'Vertical'}, actual={'merged_range': None} |
| merged_cells | tablib | merge_value_off_top_left | write | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B6:D6', 'top_left_value': 'OffTop', 'non_top_left_nonempty': 0}, actual={'merged_range': None} |
| merged_cells | tablib | merge_top_left_fill | write | data_mismatch | error | Expected values did not match actual values: expected={'merged_range': 'B7:D7', 'top_left_value': 'Fill', 'top_left_bg_color': '#FF0000'}, actual={'merged_range': None} |
| conditional_formatting | python-calamine | cf_cell_gt | read | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'cellIs', 'operator': 'greaterThan', 'formula': '5', 'priority': 1, 'format': {'bg_color': '#FFFF00'}}}, actual={} |
| conditional_formatting | python-calamine | cf_formula_cross_sheet | read | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'expression', 'formula': '=Ref!$A$1>5', 'priority': 2, 'format': {'bg_color': '#FF00FF'}}}, actual={} |
| conditional_formatting | python-calamine | cf_text_contains | read | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'expression', 'formula': '=ISNUMBER(SEARCH("foo",B2))', 'priority': 3, 'format': {'bg_color': '#FFFF00'}}}, actual={} |
| conditional_formatting | python-calamine | cf_data_bar | read | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'dataBar', 'priority': 4}}, actual={} |
| conditional_formatting | python-calamine | cf_color_scale | read | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'colorScale', 'priority': 5}}, actual={} |
| conditional_formatting | python-calamine | cf_stop_if_true | read | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B7:B9', 'rule_type': 'cellIs', 'operator': 'lessThan', 'formula': '3', 'priority': 1, 'stop_if_true': True, 'format': {'bg_color': '#FF0000'}}}, actual={} |
| conditional_formatting | pylightxl | cf_cell_gt | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| conditional_formatting | pylightxl | cf_formula_cross_sheet | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| conditional_formatting | pylightxl | cf_text_contains | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| conditional_formatting | pylightxl | cf_data_bar | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| conditional_formatting | pylightxl | cf_color_scale | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| conditional_formatting | pylightxl | cf_stop_if_true | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| conditional_formatting | pylightxl | cf_cell_gt | write | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'cellIs', 'operator': 'greaterThan', 'formula': '5', 'priority': 1, 'format': {'bg_color': '#FFFF00'}}}, actual={} |
| conditional_formatting | pylightxl | cf_formula_cross_sheet | write | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'expression', 'formula': '=Ref!$A$1>5', 'priority': 2, 'format': {'bg_color': '#FF00FF'}}}, actual={} |
| conditional_formatting | pylightxl | cf_text_contains | write | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'expression', 'formula': '=ISNUMBER(SEARCH("foo",B2))', 'priority': 3, 'format': {'bg_color': '#FFFF00'}}}, actual={} |
| conditional_formatting | pylightxl | cf_data_bar | write | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'dataBar', 'priority': 4}}, actual={} |
| conditional_formatting | pylightxl | cf_color_scale | write | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'colorScale', 'priority': 5}}, actual={} |
| conditional_formatting | pylightxl | cf_stop_if_true | write | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B7:B9', 'rule_type': 'cellIs', 'operator': 'lessThan', 'formula': '3', 'priority': 1, 'stop_if_true': True, 'format': {'bg_color': '#FF0000'}}}, actual={} |
| conditional_formatting | pyexcel | cf_cell_gt | read | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'cellIs', 'operator': 'greaterThan', 'formula': '5', 'priority': 1, 'format': {'bg_color': '#FFFF00'}}}, actual={} |
| conditional_formatting | pyexcel | cf_formula_cross_sheet | read | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'expression', 'formula': '=Ref!$A$1>5', 'priority': 2, 'format': {'bg_color': '#FF00FF'}}}, actual={} |
| conditional_formatting | pyexcel | cf_text_contains | read | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'expression', 'formula': '=ISNUMBER(SEARCH("foo",B2))', 'priority': 3, 'format': {'bg_color': '#FFFF00'}}}, actual={} |
| conditional_formatting | pyexcel | cf_data_bar | read | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'dataBar', 'priority': 4}}, actual={} |
| conditional_formatting | pyexcel | cf_color_scale | read | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'colorScale', 'priority': 5}}, actual={} |
| conditional_formatting | pyexcel | cf_stop_if_true | read | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B7:B9', 'rule_type': 'cellIs', 'operator': 'lessThan', 'formula': '3', 'priority': 1, 'stop_if_true': True, 'format': {'bg_color': '#FF0000'}}}, actual={} |
| conditional_formatting | pyexcel | cf_cell_gt | write | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'cellIs', 'operator': 'greaterThan', 'formula': '5', 'priority': 1, 'format': {'bg_color': '#FFFF00'}}}, actual={} |
| conditional_formatting | pyexcel | cf_formula_cross_sheet | write | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'expression', 'formula': '=Ref!$A$1>5', 'priority': 2, 'format': {'bg_color': '#FF00FF'}}}, actual={} |
| conditional_formatting | pyexcel | cf_text_contains | write | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'expression', 'formula': '=ISNUMBER(SEARCH("foo",B2))', 'priority': 3, 'format': {'bg_color': '#FFFF00'}}}, actual={} |
| conditional_formatting | pyexcel | cf_data_bar | write | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'dataBar', 'priority': 4}}, actual={} |
| conditional_formatting | pyexcel | cf_color_scale | write | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'colorScale', 'priority': 5}}, actual={} |
| conditional_formatting | pyexcel | cf_stop_if_true | write | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B7:B9', 'rule_type': 'cellIs', 'operator': 'lessThan', 'formula': '3', 'priority': 1, 'stop_if_true': True, 'format': {'bg_color': '#FF0000'}}}, actual={} |
| conditional_formatting | xlwt | cf_cell_gt | write | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'cellIs', 'operator': 'greaterThan', 'formula': '5', 'priority': 1, 'format': {'bg_color': '#FFFF00'}}}, actual={} |
| conditional_formatting | xlwt | cf_formula_cross_sheet | write | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'expression', 'formula': '=Ref!$A$1>5', 'priority': 2, 'format': {'bg_color': '#FF00FF'}}}, actual={} |
| conditional_formatting | xlwt | cf_text_contains | write | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'expression', 'formula': '=ISNUMBER(SEARCH("foo",B2))', 'priority': 3, 'format': {'bg_color': '#FFFF00'}}}, actual={} |
| conditional_formatting | xlwt | cf_data_bar | write | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'dataBar', 'priority': 4}}, actual={} |
| conditional_formatting | xlwt | cf_color_scale | write | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'colorScale', 'priority': 5}}, actual={} |
| conditional_formatting | xlwt | cf_stop_if_true | write | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B7:B9', 'rule_type': 'cellIs', 'operator': 'lessThan', 'formula': '3', 'priority': 1, 'stop_if_true': True, 'format': {'bg_color': '#FF0000'}}}, actual={} |
| conditional_formatting | pandas | cf_cell_gt | read | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'cellIs', 'operator': 'greaterThan', 'formula': '5', 'priority': 1, 'format': {'bg_color': '#FFFF00'}}}, actual={} |
| conditional_formatting | pandas | cf_formula_cross_sheet | read | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'expression', 'formula': '=Ref!$A$1>5', 'priority': 2, 'format': {'bg_color': '#FF00FF'}}}, actual={} |
| conditional_formatting | pandas | cf_text_contains | read | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'expression', 'formula': '=ISNUMBER(SEARCH("foo",B2))', 'priority': 3, 'format': {'bg_color': '#FFFF00'}}}, actual={} |
| conditional_formatting | pandas | cf_data_bar | read | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'dataBar', 'priority': 4}}, actual={} |
| conditional_formatting | pandas | cf_color_scale | read | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'colorScale', 'priority': 5}}, actual={} |
| conditional_formatting | pandas | cf_stop_if_true | read | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B7:B9', 'rule_type': 'cellIs', 'operator': 'lessThan', 'formula': '3', 'priority': 1, 'stop_if_true': True, 'format': {'bg_color': '#FF0000'}}}, actual={} |
| conditional_formatting | pandas | cf_cell_gt | write | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'cellIs', 'operator': 'greaterThan', 'formula': '5', 'priority': 1, 'format': {'bg_color': '#FFFF00'}}}, actual={} |
| conditional_formatting | pandas | cf_formula_cross_sheet | write | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'expression', 'formula': '=Ref!$A$1>5', 'priority': 2, 'format': {'bg_color': '#FF00FF'}}}, actual={} |
| conditional_formatting | pandas | cf_text_contains | write | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'expression', 'formula': '=ISNUMBER(SEARCH("foo",B2))', 'priority': 3, 'format': {'bg_color': '#FFFF00'}}}, actual={} |
| conditional_formatting | pandas | cf_data_bar | write | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'dataBar', 'priority': 4}}, actual={} |
| conditional_formatting | pandas | cf_color_scale | write | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'colorScale', 'priority': 5}}, actual={} |
| conditional_formatting | pandas | cf_stop_if_true | write | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B7:B9', 'rule_type': 'cellIs', 'operator': 'lessThan', 'formula': '3', 'priority': 1, 'stop_if_true': True, 'format': {'bg_color': '#FF0000'}}}, actual={} |
| conditional_formatting | openpyxl-readonly | cf_cell_gt | read | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'cellIs', 'operator': 'greaterThan', 'formula': '5', 'priority': 1, 'format': {'bg_color': '#FFFF00'}}}, actual={} |
| conditional_formatting | openpyxl-readonly | cf_formula_cross_sheet | read | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'expression', 'formula': '=Ref!$A$1>5', 'priority': 2, 'format': {'bg_color': '#FF00FF'}}}, actual={} |
| conditional_formatting | openpyxl-readonly | cf_text_contains | read | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'expression', 'formula': '=ISNUMBER(SEARCH("foo",B2))', 'priority': 3, 'format': {'bg_color': '#FFFF00'}}}, actual={} |
| conditional_formatting | openpyxl-readonly | cf_data_bar | read | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'dataBar', 'priority': 4}}, actual={} |
| conditional_formatting | openpyxl-readonly | cf_color_scale | read | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'colorScale', 'priority': 5}}, actual={} |
| conditional_formatting | openpyxl-readonly | cf_stop_if_true | read | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B7:B9', 'rule_type': 'cellIs', 'operator': 'lessThan', 'formula': '3', 'priority': 1, 'stop_if_true': True, 'format': {'bg_color': '#FF0000'}}}, actual={} |
| conditional_formatting | polars | cf_cell_gt | read | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'cellIs', 'operator': 'greaterThan', 'formula': '5', 'priority': 1, 'format': {'bg_color': '#FFFF00'}}}, actual={} |
| conditional_formatting | polars | cf_formula_cross_sheet | read | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'expression', 'formula': '=Ref!$A$1>5', 'priority': 2, 'format': {'bg_color': '#FF00FF'}}}, actual={} |
| conditional_formatting | polars | cf_text_contains | read | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'expression', 'formula': '=ISNUMBER(SEARCH("foo",B2))', 'priority': 3, 'format': {'bg_color': '#FFFF00'}}}, actual={} |
| conditional_formatting | polars | cf_data_bar | read | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'dataBar', 'priority': 4}}, actual={} |
| conditional_formatting | polars | cf_color_scale | read | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'colorScale', 'priority': 5}}, actual={} |
| conditional_formatting | polars | cf_stop_if_true | read | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B7:B9', 'rule_type': 'cellIs', 'operator': 'lessThan', 'formula': '3', 'priority': 1, 'stop_if_true': True, 'format': {'bg_color': '#FF0000'}}}, actual={} |
| conditional_formatting | tablib | cf_cell_gt | read | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'cellIs', 'operator': 'greaterThan', 'formula': '5', 'priority': 1, 'format': {'bg_color': '#FFFF00'}}}, actual={} |
| conditional_formatting | tablib | cf_formula_cross_sheet | read | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'expression', 'formula': '=Ref!$A$1>5', 'priority': 2, 'format': {'bg_color': '#FF00FF'}}}, actual={} |
| conditional_formatting | tablib | cf_text_contains | read | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'expression', 'formula': '=ISNUMBER(SEARCH("foo",B2))', 'priority': 3, 'format': {'bg_color': '#FFFF00'}}}, actual={} |
| conditional_formatting | tablib | cf_data_bar | read | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'dataBar', 'priority': 4}}, actual={} |
| conditional_formatting | tablib | cf_color_scale | read | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'colorScale', 'priority': 5}}, actual={} |
| conditional_formatting | tablib | cf_stop_if_true | read | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B7:B9', 'rule_type': 'cellIs', 'operator': 'lessThan', 'formula': '3', 'priority': 1, 'stop_if_true': True, 'format': {'bg_color': '#FF0000'}}}, actual={} |
| conditional_formatting | tablib | cf_cell_gt | write | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'cellIs', 'operator': 'greaterThan', 'formula': '5', 'priority': 1, 'format': {'bg_color': '#FFFF00'}}}, actual={} |
| conditional_formatting | tablib | cf_formula_cross_sheet | write | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'expression', 'formula': '=Ref!$A$1>5', 'priority': 2, 'format': {'bg_color': '#FF00FF'}}}, actual={} |
| conditional_formatting | tablib | cf_text_contains | write | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'expression', 'formula': '=ISNUMBER(SEARCH("foo",B2))', 'priority': 3, 'format': {'bg_color': '#FFFF00'}}}, actual={} |
| conditional_formatting | tablib | cf_data_bar | write | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'dataBar', 'priority': 4}}, actual={} |
| conditional_formatting | tablib | cf_color_scale | write | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B2:B6', 'rule_type': 'colorScale', 'priority': 5}}, actual={} |
| conditional_formatting | tablib | cf_stop_if_true | write | data_mismatch | error | Expected values did not match actual values: expected={'cf_rule': {'range': 'B7:B9', 'rule_type': 'cellIs', 'operator': 'lessThan', 'formula': '3', 'priority': 1, 'stop_if_true': True, 'format': {'bg_color': '#FF0000'}}}, actual={} |
| data_validation | python-calamine | dv_list_csv | read | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B2', 'validation_type': 'list', 'formula1': '"Red,Green,Blue"', 'allow_blank': True}}, actual={} |
| data_validation | python-calamine | dv_list_range | read | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B3', 'validation_type': 'list', 'formula1': '=$D$2:$D$4'}}, actual={} |
| data_validation | python-calamine | dv_cross_sheet | read | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B4', 'validation_type': 'list', 'formula1': '=RegionList'}}, actual={} |
| data_validation | python-calamine | dv_custom_formula | read | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B5', 'validation_type': 'custom', 'formula1': '=B5>C5'}}, actual={} |
| data_validation | python-calamine | dv_whole_between | read | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B6', 'validation_type': 'whole', 'operator': 'between', 'formula1': '1', 'formula2': '10', 'allow_blank': False, 'error_title': 'Invalid', 'error': 'Enter 1-10'}}, actual={} |
| data_validation | pylightxl | dv_list_csv | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| data_validation | pylightxl | dv_list_range | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| data_validation | pylightxl | dv_cross_sheet | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| data_validation | pylightxl | dv_custom_formula | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| data_validation | pylightxl | dv_whole_between | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| data_validation | pylightxl | dv_list_csv | write | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B2', 'validation_type': 'list', 'formula1': '"Red,Green,Blue"', 'allow_blank': True}}, actual={} |
| data_validation | pylightxl | dv_list_range | write | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B3', 'validation_type': 'list', 'formula1': '=$D$2:$D$4'}}, actual={} |
| data_validation | pylightxl | dv_cross_sheet | write | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B4', 'validation_type': 'list', 'formula1': '=RegionList'}}, actual={} |
| data_validation | pylightxl | dv_custom_formula | write | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B5', 'validation_type': 'custom', 'formula1': '=B5>C5'}}, actual={} |
| data_validation | pylightxl | dv_whole_between | write | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B6', 'validation_type': 'whole', 'operator': 'between', 'formula1': '1', 'formula2': '10', 'allow_blank': False, 'error_title': 'Invalid', 'error': 'Enter 1-10'}}, actual={} |
| data_validation | pyexcel | dv_list_csv | read | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B2', 'validation_type': 'list', 'formula1': '"Red,Green,Blue"', 'allow_blank': True}}, actual={} |
| data_validation | pyexcel | dv_list_range | read | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B3', 'validation_type': 'list', 'formula1': '=$D$2:$D$4'}}, actual={} |
| data_validation | pyexcel | dv_cross_sheet | read | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B4', 'validation_type': 'list', 'formula1': '=RegionList'}}, actual={} |
| data_validation | pyexcel | dv_custom_formula | read | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B5', 'validation_type': 'custom', 'formula1': '=B5>C5'}}, actual={} |
| data_validation | pyexcel | dv_whole_between | read | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B6', 'validation_type': 'whole', 'operator': 'between', 'formula1': '1', 'formula2': '10', 'allow_blank': False, 'error_title': 'Invalid', 'error': 'Enter 1-10'}}, actual={} |
| data_validation | pyexcel | dv_list_csv | write | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B2', 'validation_type': 'list', 'formula1': '"Red,Green,Blue"', 'allow_blank': True}}, actual={} |
| data_validation | pyexcel | dv_list_range | write | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B3', 'validation_type': 'list', 'formula1': '=$D$2:$D$4'}}, actual={} |
| data_validation | pyexcel | dv_cross_sheet | write | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B4', 'validation_type': 'list', 'formula1': '=RegionList'}}, actual={} |
| data_validation | pyexcel | dv_custom_formula | write | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B5', 'validation_type': 'custom', 'formula1': '=B5>C5'}}, actual={} |
| data_validation | pyexcel | dv_whole_between | write | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B6', 'validation_type': 'whole', 'operator': 'between', 'formula1': '1', 'formula2': '10', 'allow_blank': False, 'error_title': 'Invalid', 'error': 'Enter 1-10'}}, actual={} |
| data_validation | xlwt | dv_list_csv | write | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B2', 'validation_type': 'list', 'formula1': '"Red,Green,Blue"', 'allow_blank': True}}, actual={} |
| data_validation | xlwt | dv_list_range | write | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B3', 'validation_type': 'list', 'formula1': '=$D$2:$D$4'}}, actual={} |
| data_validation | xlwt | dv_cross_sheet | write | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B4', 'validation_type': 'list', 'formula1': '=RegionList'}}, actual={} |
| data_validation | xlwt | dv_custom_formula | write | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B5', 'validation_type': 'custom', 'formula1': '=B5>C5'}}, actual={} |
| data_validation | xlwt | dv_whole_between | write | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B6', 'validation_type': 'whole', 'operator': 'between', 'formula1': '1', 'formula2': '10', 'allow_blank': False, 'error_title': 'Invalid', 'error': 'Enter 1-10'}}, actual={} |
| data_validation | pandas | dv_list_csv | read | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B2', 'validation_type': 'list', 'formula1': '"Red,Green,Blue"', 'allow_blank': True}}, actual={} |
| data_validation | pandas | dv_list_range | read | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B3', 'validation_type': 'list', 'formula1': '=$D$2:$D$4'}}, actual={} |
| data_validation | pandas | dv_cross_sheet | read | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B4', 'validation_type': 'list', 'formula1': '=RegionList'}}, actual={} |
| data_validation | pandas | dv_custom_formula | read | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B5', 'validation_type': 'custom', 'formula1': '=B5>C5'}}, actual={} |
| data_validation | pandas | dv_whole_between | read | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B6', 'validation_type': 'whole', 'operator': 'between', 'formula1': '1', 'formula2': '10', 'allow_blank': False, 'error_title': 'Invalid', 'error': 'Enter 1-10'}}, actual={} |
| data_validation | pandas | dv_list_csv | write | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B2', 'validation_type': 'list', 'formula1': '"Red,Green,Blue"', 'allow_blank': True}}, actual={} |
| data_validation | pandas | dv_list_range | write | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B3', 'validation_type': 'list', 'formula1': '=$D$2:$D$4'}}, actual={} |
| data_validation | pandas | dv_cross_sheet | write | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B4', 'validation_type': 'list', 'formula1': '=RegionList'}}, actual={} |
| data_validation | pandas | dv_custom_formula | write | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B5', 'validation_type': 'custom', 'formula1': '=B5>C5'}}, actual={} |
| data_validation | pandas | dv_whole_between | write | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B6', 'validation_type': 'whole', 'operator': 'between', 'formula1': '1', 'formula2': '10', 'allow_blank': False, 'error_title': 'Invalid', 'error': 'Enter 1-10'}}, actual={} |
| data_validation | openpyxl-readonly | dv_list_csv | read | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B2', 'validation_type': 'list', 'formula1': '"Red,Green,Blue"', 'allow_blank': True}}, actual={} |
| data_validation | openpyxl-readonly | dv_list_range | read | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B3', 'validation_type': 'list', 'formula1': '=$D$2:$D$4'}}, actual={} |
| data_validation | openpyxl-readonly | dv_cross_sheet | read | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B4', 'validation_type': 'list', 'formula1': '=RegionList'}}, actual={} |
| data_validation | openpyxl-readonly | dv_custom_formula | read | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B5', 'validation_type': 'custom', 'formula1': '=B5>C5'}}, actual={} |
| data_validation | openpyxl-readonly | dv_whole_between | read | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B6', 'validation_type': 'whole', 'operator': 'between', 'formula1': '1', 'formula2': '10', 'allow_blank': False, 'error_title': 'Invalid', 'error': 'Enter 1-10'}}, actual={} |
| data_validation | polars | dv_list_csv | read | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B2', 'validation_type': 'list', 'formula1': '"Red,Green,Blue"', 'allow_blank': True}}, actual={} |
| data_validation | polars | dv_list_range | read | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B3', 'validation_type': 'list', 'formula1': '=$D$2:$D$4'}}, actual={} |
| data_validation | polars | dv_cross_sheet | read | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B4', 'validation_type': 'list', 'formula1': '=RegionList'}}, actual={} |
| data_validation | polars | dv_custom_formula | read | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B5', 'validation_type': 'custom', 'formula1': '=B5>C5'}}, actual={} |
| data_validation | polars | dv_whole_between | read | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B6', 'validation_type': 'whole', 'operator': 'between', 'formula1': '1', 'formula2': '10', 'allow_blank': False, 'error_title': 'Invalid', 'error': 'Enter 1-10'}}, actual={} |
| data_validation | tablib | dv_list_csv | read | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B2', 'validation_type': 'list', 'formula1': '"Red,Green,Blue"', 'allow_blank': True}}, actual={} |
| data_validation | tablib | dv_list_range | read | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B3', 'validation_type': 'list', 'formula1': '=$D$2:$D$4'}}, actual={} |
| data_validation | tablib | dv_cross_sheet | read | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B4', 'validation_type': 'list', 'formula1': '=RegionList'}}, actual={} |
| data_validation | tablib | dv_custom_formula | read | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B5', 'validation_type': 'custom', 'formula1': '=B5>C5'}}, actual={} |
| data_validation | tablib | dv_whole_between | read | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B6', 'validation_type': 'whole', 'operator': 'between', 'formula1': '1', 'formula2': '10', 'allow_blank': False, 'error_title': 'Invalid', 'error': 'Enter 1-10'}}, actual={} |
| data_validation | tablib | dv_list_csv | write | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B2', 'validation_type': 'list', 'formula1': '"Red,Green,Blue"', 'allow_blank': True}}, actual={} |
| data_validation | tablib | dv_list_range | write | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B3', 'validation_type': 'list', 'formula1': '=$D$2:$D$4'}}, actual={} |
| data_validation | tablib | dv_cross_sheet | write | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B4', 'validation_type': 'list', 'formula1': '=RegionList'}}, actual={} |
| data_validation | tablib | dv_custom_formula | write | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B5', 'validation_type': 'custom', 'formula1': '=B5>C5'}}, actual={} |
| data_validation | tablib | dv_whole_between | write | data_mismatch | error | Expected values did not match actual values: expected={'validation': {'range': 'B6', 'validation_type': 'whole', 'operator': 'between', 'formula1': '1', 'formula2': '10', 'allow_blank': False, 'error_title': 'Invalid', 'error': 'Enter 1-10'}}, actual={} |
| hyperlinks | python-calamine | link_external | read | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B2', 'target': 'https://example.com/docs', 'display': 'Example Docs', 'tooltip': 'Go to docs', 'internal': False}}, actual={} |
| hyperlinks | python-calamine | link_internal | read | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B3', 'target': 'Targets!A1', 'display': 'Go Target', 'tooltip': 'Jump to target', 'internal': True}}, actual={} |
| hyperlinks | python-calamine | link_mailto | read | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B4', 'target': 'mailto:test@example.com', 'display': 'Email', 'tooltip': 'Send email', 'internal': False}}, actual={} |
| hyperlinks | python-calamine | link_long | read | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B5', 'target': 'https://example.com/search?q=excel%20bench&sort=desc#section-2', 'display': 'Search', 'tooltip': 'Encoded URL', 'internal': False}}, actual={} |
| hyperlinks | pylightxl | link_external | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| hyperlinks | pylightxl | link_internal | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| hyperlinks | pylightxl | link_mailto | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| hyperlinks | pylightxl | link_long | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| hyperlinks | pylightxl | link_external | write | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B2', 'target': 'https://example.com/docs', 'display': 'Example Docs', 'tooltip': 'Go to docs', 'internal': False}}, actual={} |
| hyperlinks | pylightxl | link_internal | write | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B3', 'target': 'Targets!A1', 'display': 'Go Target', 'tooltip': 'Jump to target', 'internal': True}}, actual={} |
| hyperlinks | pylightxl | link_mailto | write | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B4', 'target': 'mailto:test@example.com', 'display': 'Email', 'tooltip': 'Send email', 'internal': False}}, actual={} |
| hyperlinks | pylightxl | link_long | write | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B5', 'target': 'https://example.com/search?q=excel%20bench&sort=desc#section-2', 'display': 'Search', 'tooltip': 'Encoded URL', 'internal': False}}, actual={} |
| hyperlinks | pyexcel | link_external | read | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B2', 'target': 'https://example.com/docs', 'display': 'Example Docs', 'tooltip': 'Go to docs', 'internal': False}}, actual={} |
| hyperlinks | pyexcel | link_internal | read | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B3', 'target': 'Targets!A1', 'display': 'Go Target', 'tooltip': 'Jump to target', 'internal': True}}, actual={} |
| hyperlinks | pyexcel | link_mailto | read | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B4', 'target': 'mailto:test@example.com', 'display': 'Email', 'tooltip': 'Send email', 'internal': False}}, actual={} |
| hyperlinks | pyexcel | link_long | read | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B5', 'target': 'https://example.com/search?q=excel%20bench&sort=desc#section-2', 'display': 'Search', 'tooltip': 'Encoded URL', 'internal': False}}, actual={} |
| hyperlinks | pyexcel | link_external | write | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B2', 'target': 'https://example.com/docs', 'display': 'Example Docs', 'tooltip': 'Go to docs', 'internal': False}}, actual={} |
| hyperlinks | pyexcel | link_internal | write | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B3', 'target': 'Targets!A1', 'display': 'Go Target', 'tooltip': 'Jump to target', 'internal': True}}, actual={} |
| hyperlinks | pyexcel | link_mailto | write | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B4', 'target': 'mailto:test@example.com', 'display': 'Email', 'tooltip': 'Send email', 'internal': False}}, actual={} |
| hyperlinks | pyexcel | link_long | write | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B5', 'target': 'https://example.com/search?q=excel%20bench&sort=desc#section-2', 'display': 'Search', 'tooltip': 'Encoded URL', 'internal': False}}, actual={} |
| hyperlinks | xlwt | link_external | write | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B2', 'target': 'https://example.com/docs', 'display': 'Example Docs', 'tooltip': 'Go to docs', 'internal': False}}, actual={} |
| hyperlinks | xlwt | link_internal | write | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B3', 'target': 'Targets!A1', 'display': 'Go Target', 'tooltip': 'Jump to target', 'internal': True}}, actual={} |
| hyperlinks | xlwt | link_mailto | write | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B4', 'target': 'mailto:test@example.com', 'display': 'Email', 'tooltip': 'Send email', 'internal': False}}, actual={} |
| hyperlinks | xlwt | link_long | write | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B5', 'target': 'https://example.com/search?q=excel%20bench&sort=desc#section-2', 'display': 'Search', 'tooltip': 'Encoded URL', 'internal': False}}, actual={} |
| hyperlinks | pandas | link_external | read | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B2', 'target': 'https://example.com/docs', 'display': 'Example Docs', 'tooltip': 'Go to docs', 'internal': False}}, actual={} |
| hyperlinks | pandas | link_internal | read | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B3', 'target': 'Targets!A1', 'display': 'Go Target', 'tooltip': 'Jump to target', 'internal': True}}, actual={} |
| hyperlinks | pandas | link_mailto | read | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B4', 'target': 'mailto:test@example.com', 'display': 'Email', 'tooltip': 'Send email', 'internal': False}}, actual={} |
| hyperlinks | pandas | link_long | read | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B5', 'target': 'https://example.com/search?q=excel%20bench&sort=desc#section-2', 'display': 'Search', 'tooltip': 'Encoded URL', 'internal': False}}, actual={} |
| hyperlinks | pandas | link_external | write | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B2', 'target': 'https://example.com/docs', 'display': 'Example Docs', 'tooltip': 'Go to docs', 'internal': False}}, actual={} |
| hyperlinks | pandas | link_internal | write | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B3', 'target': 'Targets!A1', 'display': 'Go Target', 'tooltip': 'Jump to target', 'internal': True}}, actual={} |
| hyperlinks | pandas | link_mailto | write | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B4', 'target': 'mailto:test@example.com', 'display': 'Email', 'tooltip': 'Send email', 'internal': False}}, actual={} |
| hyperlinks | pandas | link_long | write | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B5', 'target': 'https://example.com/search?q=excel%20bench&sort=desc#section-2', 'display': 'Search', 'tooltip': 'Encoded URL', 'internal': False}}, actual={} |
| hyperlinks | openpyxl-readonly | link_external | read | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B2', 'target': 'https://example.com/docs', 'display': 'Example Docs', 'tooltip': 'Go to docs', 'internal': False}}, actual={} |
| hyperlinks | openpyxl-readonly | link_internal | read | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B3', 'target': 'Targets!A1', 'display': 'Go Target', 'tooltip': 'Jump to target', 'internal': True}}, actual={} |
| hyperlinks | openpyxl-readonly | link_mailto | read | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B4', 'target': 'mailto:test@example.com', 'display': 'Email', 'tooltip': 'Send email', 'internal': False}}, actual={} |
| hyperlinks | openpyxl-readonly | link_long | read | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B5', 'target': 'https://example.com/search?q=excel%20bench&sort=desc#section-2', 'display': 'Search', 'tooltip': 'Encoded URL', 'internal': False}}, actual={} |
| hyperlinks | polars | link_external | read | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B2', 'target': 'https://example.com/docs', 'display': 'Example Docs', 'tooltip': 'Go to docs', 'internal': False}}, actual={} |
| hyperlinks | polars | link_internal | read | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B3', 'target': 'Targets!A1', 'display': 'Go Target', 'tooltip': 'Jump to target', 'internal': True}}, actual={} |
| hyperlinks | polars | link_mailto | read | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B4', 'target': 'mailto:test@example.com', 'display': 'Email', 'tooltip': 'Send email', 'internal': False}}, actual={} |
| hyperlinks | polars | link_long | read | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B5', 'target': 'https://example.com/search?q=excel%20bench&sort=desc#section-2', 'display': 'Search', 'tooltip': 'Encoded URL', 'internal': False}}, actual={} |
| hyperlinks | tablib | link_external | read | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B2', 'target': 'https://example.com/docs', 'display': 'Example Docs', 'tooltip': 'Go to docs', 'internal': False}}, actual={} |
| hyperlinks | tablib | link_internal | read | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B3', 'target': 'Targets!A1', 'display': 'Go Target', 'tooltip': 'Jump to target', 'internal': True}}, actual={} |
| hyperlinks | tablib | link_mailto | read | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B4', 'target': 'mailto:test@example.com', 'display': 'Email', 'tooltip': 'Send email', 'internal': False}}, actual={} |
| hyperlinks | tablib | link_long | read | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B5', 'target': 'https://example.com/search?q=excel%20bench&sort=desc#section-2', 'display': 'Search', 'tooltip': 'Encoded URL', 'internal': False}}, actual={} |
| hyperlinks | tablib | link_external | write | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B2', 'target': 'https://example.com/docs', 'display': 'Example Docs', 'tooltip': 'Go to docs', 'internal': False}}, actual={} |
| hyperlinks | tablib | link_internal | write | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B3', 'target': 'Targets!A1', 'display': 'Go Target', 'tooltip': 'Jump to target', 'internal': True}}, actual={} |
| hyperlinks | tablib | link_mailto | write | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B4', 'target': 'mailto:test@example.com', 'display': 'Email', 'tooltip': 'Send email', 'internal': False}}, actual={} |
| hyperlinks | tablib | link_long | write | data_mismatch | error | Expected values did not match actual values: expected={'hyperlink': {'cell': 'B5', 'target': 'https://example.com/search?q=excel%20bench&sort=desc#section-2', 'display': 'Search', 'tooltip': 'Encoded URL', 'internal': False}}, actual={} |
| images | python-calamine | image_one_cell | read | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'B2', 'path': 'fixtures/images/sample.png', 'anchor': 'oneCell'}}, actual={} |
| images | python-calamine | image_two_cell_offset | read | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'D6', 'path': 'fixtures/images/sample.jpg', 'anchor': 'oneCell'}}, actual={} |
| images | rust_xlsxwriter | image_one_cell | write | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'B2', 'path': 'fixtures/images/sample.png', 'anchor': 'oneCell'}}, actual={} |
| images | rust_xlsxwriter | image_two_cell_offset | write | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'D6', 'path': 'fixtures/images/sample.jpg', 'anchor': 'oneCell'}}, actual={} |
| images | wolfxl | image_one_cell | read | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'B2', 'path': 'fixtures/images/sample.png', 'anchor': 'oneCell'}}, actual={} |
| images | wolfxl | image_two_cell_offset | read | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'D6', 'path': 'fixtures/images/sample.jpg', 'anchor': 'oneCell'}}, actual={} |
| images | wolfxl | image_one_cell | write | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'B2', 'path': 'fixtures/images/sample.png', 'anchor': 'oneCell'}}, actual={} |
| images | wolfxl | image_two_cell_offset | write | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'D6', 'path': 'fixtures/images/sample.jpg', 'anchor': 'oneCell'}}, actual={} |
| images | pylightxl | image_one_cell | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| images | pylightxl | image_two_cell_offset | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| images | pylightxl | image_one_cell | write | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'B2', 'path': 'fixtures/images/sample.png', 'anchor': 'oneCell'}}, actual={} |
| images | pylightxl | image_two_cell_offset | write | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'D6', 'path': 'fixtures/images/sample.jpg', 'anchor': 'oneCell'}}, actual={} |
| images | pyexcel | image_one_cell | read | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'B2', 'path': 'fixtures/images/sample.png', 'anchor': 'oneCell'}}, actual={} |
| images | pyexcel | image_two_cell_offset | read | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'D6', 'path': 'fixtures/images/sample.jpg', 'anchor': 'oneCell'}}, actual={} |
| images | pyexcel | image_one_cell | write | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'B2', 'path': 'fixtures/images/sample.png', 'anchor': 'oneCell'}}, actual={} |
| images | pyexcel | image_two_cell_offset | write | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'D6', 'path': 'fixtures/images/sample.jpg', 'anchor': 'oneCell'}}, actual={} |
| images | xlwt | image_one_cell | write | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'B2', 'path': 'fixtures/images/sample.png', 'anchor': 'oneCell'}}, actual={} |
| images | xlwt | image_two_cell_offset | write | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'D6', 'path': 'fixtures/images/sample.jpg', 'anchor': 'oneCell'}}, actual={} |
| images | pandas | image_one_cell | read | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'B2', 'path': 'fixtures/images/sample.png', 'anchor': 'oneCell'}}, actual={} |
| images | pandas | image_two_cell_offset | read | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'D6', 'path': 'fixtures/images/sample.jpg', 'anchor': 'oneCell'}}, actual={} |
| images | pandas | image_one_cell | write | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'B2', 'path': 'fixtures/images/sample.png', 'anchor': 'oneCell'}}, actual={} |
| images | pandas | image_two_cell_offset | write | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'D6', 'path': 'fixtures/images/sample.jpg', 'anchor': 'oneCell'}}, actual={} |
| images | xlsxwriter-constmem | image_one_cell | write | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'B2', 'path': 'fixtures/images/sample.png', 'anchor': 'oneCell'}}, actual={} |
| images | xlsxwriter-constmem | image_two_cell_offset | write | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'D6', 'path': 'fixtures/images/sample.jpg', 'anchor': 'oneCell'}}, actual={} |
| images | openpyxl-readonly | image_one_cell | read | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'B2', 'path': 'fixtures/images/sample.png', 'anchor': 'oneCell'}}, actual={} |
| images | openpyxl-readonly | image_two_cell_offset | read | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'D6', 'path': 'fixtures/images/sample.jpg', 'anchor': 'oneCell'}}, actual={} |
| images | polars | image_one_cell | read | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'B2', 'path': 'fixtures/images/sample.png', 'anchor': 'oneCell'}}, actual={} |
| images | polars | image_two_cell_offset | read | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'D6', 'path': 'fixtures/images/sample.jpg', 'anchor': 'oneCell'}}, actual={} |
| images | tablib | image_one_cell | read | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'B2', 'path': 'fixtures/images/sample.png', 'anchor': 'oneCell'}}, actual={} |
| images | tablib | image_two_cell_offset | read | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'D6', 'path': 'fixtures/images/sample.jpg', 'anchor': 'oneCell'}}, actual={} |
| images | tablib | image_one_cell | write | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'B2', 'path': 'fixtures/images/sample.png', 'anchor': 'oneCell'}}, actual={} |
| images | tablib | image_two_cell_offset | write | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'D6', 'path': 'fixtures/images/sample.jpg', 'anchor': 'oneCell'}}, actual={} |
| comments | python-calamine | comment_legacy | read | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B2', 'text': 'Legacy note', 'threaded': False}}, actual={} |
| comments | python-calamine | comment_threaded | read | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B3', 'text': 'Threaded fallback', 'threaded': False}}, actual={} |
| comments | python-calamine | comment_author | read | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B4', 'text': 'Another note', 'threaded': False}}, actual={} |
| comments | pylightxl | comment_legacy | read | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B2', 'text': 'Legacy note', 'threaded': False}}, actual={} |
| comments | pylightxl | comment_threaded | read | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B3', 'text': 'Threaded fallback', 'threaded': False}}, actual={} |
| comments | pylightxl | comment_author | read | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B4', 'text': 'Another note', 'threaded': False}}, actual={} |
| comments | pylightxl | comment_legacy | write | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B2', 'text': 'Legacy note', 'threaded': False}}, actual={} |
| comments | pylightxl | comment_threaded | write | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B3', 'text': 'Threaded fallback', 'threaded': False}}, actual={} |
| comments | pylightxl | comment_author | write | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B4', 'text': 'Another note', 'threaded': False}}, actual={} |
| comments | pyexcel | comment_legacy | read | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B2', 'text': 'Legacy note', 'threaded': False}}, actual={} |
| comments | pyexcel | comment_threaded | read | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B3', 'text': 'Threaded fallback', 'threaded': False}}, actual={} |
| comments | pyexcel | comment_author | read | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B4', 'text': 'Another note', 'threaded': False}}, actual={} |
| comments | pyexcel | comment_legacy | write | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B2', 'text': 'Legacy note', 'threaded': False}}, actual={} |
| comments | pyexcel | comment_threaded | write | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B3', 'text': 'Threaded fallback', 'threaded': False}}, actual={} |
| comments | pyexcel | comment_author | write | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B4', 'text': 'Another note', 'threaded': False}}, actual={} |
| comments | xlwt | comment_legacy | write | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B2', 'text': 'Legacy note', 'threaded': False}}, actual={} |
| comments | xlwt | comment_threaded | write | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B3', 'text': 'Threaded fallback', 'threaded': False}}, actual={} |
| comments | xlwt | comment_author | write | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B4', 'text': 'Another note', 'threaded': False}}, actual={} |
| comments | pandas | comment_legacy | read | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B2', 'text': 'Legacy note', 'threaded': False}}, actual={} |
| comments | pandas | comment_threaded | read | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B3', 'text': 'Threaded fallback', 'threaded': False}}, actual={} |
| comments | pandas | comment_author | read | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B4', 'text': 'Another note', 'threaded': False}}, actual={} |
| comments | pandas | comment_legacy | write | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B2', 'text': 'Legacy note', 'threaded': False}}, actual={} |
| comments | pandas | comment_threaded | write | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B3', 'text': 'Threaded fallback', 'threaded': False}}, actual={} |
| comments | pandas | comment_author | write | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B4', 'text': 'Another note', 'threaded': False}}, actual={} |
| comments | xlsxwriter-constmem | comment_legacy | write | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B2', 'text': 'Legacy note', 'threaded': False}}, actual={} |
| comments | xlsxwriter-constmem | comment_threaded | write | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B3', 'text': 'Threaded fallback', 'threaded': False}}, actual={} |
| comments | xlsxwriter-constmem | comment_author | write | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B4', 'text': 'Another note', 'threaded': False}}, actual={} |
| comments | openpyxl-readonly | comment_legacy | read | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B2', 'text': 'Legacy note', 'threaded': False}}, actual={} |
| comments | openpyxl-readonly | comment_threaded | read | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B3', 'text': 'Threaded fallback', 'threaded': False}}, actual={} |
| comments | openpyxl-readonly | comment_author | read | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B4', 'text': 'Another note', 'threaded': False}}, actual={} |
| comments | polars | comment_legacy | read | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B2', 'text': 'Legacy note', 'threaded': False}}, actual={} |
| comments | polars | comment_threaded | read | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B3', 'text': 'Threaded fallback', 'threaded': False}}, actual={} |
| comments | polars | comment_author | read | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B4', 'text': 'Another note', 'threaded': False}}, actual={} |
| comments | tablib | comment_legacy | read | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B2', 'text': 'Legacy note', 'threaded': False}}, actual={} |
| comments | tablib | comment_threaded | read | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B3', 'text': 'Threaded fallback', 'threaded': False}}, actual={} |
| comments | tablib | comment_author | read | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B4', 'text': 'Another note', 'threaded': False}}, actual={} |
| comments | tablib | comment_legacy | write | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B2', 'text': 'Legacy note', 'threaded': False}}, actual={} |
| comments | tablib | comment_threaded | write | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B3', 'text': 'Threaded fallback', 'threaded': False}}, actual={} |
| comments | tablib | comment_author | write | data_mismatch | error | Expected values did not match actual values: expected={'comment': {'cell': 'B4', 'text': 'Another note', 'threaded': False}}, actual={} |
| freeze_panes | python-calamine | freeze_b2 | read | data_mismatch | error | Expected values did not match actual values: expected={'freeze': {'mode': 'freeze', 'top_left_cell': 'B2'}}, actual={'freeze': {'mode': None, 'top_left_cell': None}} |
| freeze_panes | python-calamine | freeze_d5 | read | data_mismatch | error | Expected values did not match actual values: expected={'freeze': {'mode': 'freeze', 'top_left_cell': 'D5'}}, actual={'freeze': {'mode': None, 'top_left_cell': None}} |
| freeze_panes | python-calamine | split_2x1 | read | data_mismatch | error | Expected values did not match actual values: expected={'freeze': {'mode': 'split', 'x_split': 1, 'y_split': 2}}, actual={'freeze': {'mode': None, 'x_split': None, 'y_split': None}} |
| freeze_panes | pylightxl | freeze_b2 | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| freeze_panes | pylightxl | freeze_d5 | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| freeze_panes | pylightxl | split_2x1 | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| freeze_panes | pylightxl | freeze_b2 | write | data_mismatch | error | Expected values did not match actual values: expected={'freeze': {'mode': 'freeze', 'top_left_cell': 'B2'}}, actual={'freeze': {'mode': None, 'top_left_cell': None}} |
| freeze_panes | pylightxl | freeze_d5 | write | data_mismatch | error | Expected values did not match actual values: expected={'freeze': {'mode': 'freeze', 'top_left_cell': 'D5'}}, actual={'freeze': {'mode': None, 'top_left_cell': None}} |
| freeze_panes | pylightxl | split_2x1 | write | data_mismatch | error | Expected values did not match actual values: expected={'freeze': {'mode': 'split', 'x_split': 1, 'y_split': 2}}, actual={'freeze': {'mode': None, 'x_split': None, 'y_split': None}} |
| freeze_panes | pyexcel | freeze_b2 | read | data_mismatch | error | Expected values did not match actual values: expected={'freeze': {'mode': 'freeze', 'top_left_cell': 'B2'}}, actual={'freeze': {'mode': None, 'top_left_cell': None}} |
| freeze_panes | pyexcel | freeze_d5 | read | data_mismatch | error | Expected values did not match actual values: expected={'freeze': {'mode': 'freeze', 'top_left_cell': 'D5'}}, actual={'freeze': {'mode': None, 'top_left_cell': None}} |
| freeze_panes | pyexcel | split_2x1 | read | data_mismatch | error | Expected values did not match actual values: expected={'freeze': {'mode': 'split', 'x_split': 1, 'y_split': 2}}, actual={'freeze': {'mode': None, 'x_split': None, 'y_split': None}} |
| freeze_panes | pyexcel | freeze_b2 | write | data_mismatch | error | Expected values did not match actual values: expected={'freeze': {'mode': 'freeze', 'top_left_cell': 'B2'}}, actual={'freeze': {'mode': None, 'top_left_cell': None}} |
| freeze_panes | pyexcel | freeze_d5 | write | data_mismatch | error | Expected values did not match actual values: expected={'freeze': {'mode': 'freeze', 'top_left_cell': 'D5'}}, actual={'freeze': {'mode': None, 'top_left_cell': None}} |
| freeze_panes | pyexcel | split_2x1 | write | data_mismatch | error | Expected values did not match actual values: expected={'freeze': {'mode': 'split', 'x_split': 1, 'y_split': 2}}, actual={'freeze': {'mode': None, 'x_split': None, 'y_split': None}} |
| freeze_panes | xlwt | freeze_b2 | write | internal | error | AttributeError: 'Sheet' object has no attribute 'frozen_row_count' |
| freeze_panes | xlwt | freeze_d5 | write | internal | error | AttributeError: 'Sheet' object has no attribute 'frozen_row_count' |
| freeze_panes | xlwt | split_2x1 | write | internal | error | AttributeError: 'Sheet' object has no attribute 'frozen_row_count' |
| freeze_panes | pandas | freeze_b2 | read | data_mismatch | error | Expected values did not match actual values: expected={'freeze': {'mode': 'freeze', 'top_left_cell': 'B2'}}, actual={'freeze': {'mode': None, 'top_left_cell': None}} |
| freeze_panes | pandas | freeze_d5 | read | data_mismatch | error | Expected values did not match actual values: expected={'freeze': {'mode': 'freeze', 'top_left_cell': 'D5'}}, actual={'freeze': {'mode': None, 'top_left_cell': None}} |
| freeze_panes | pandas | split_2x1 | read | data_mismatch | error | Expected values did not match actual values: expected={'freeze': {'mode': 'split', 'x_split': 1, 'y_split': 2}}, actual={'freeze': {'mode': None, 'x_split': None, 'y_split': None}} |
| freeze_panes | pandas | freeze_b2 | write | data_mismatch | error | Expected values did not match actual values: expected={'freeze': {'mode': 'freeze', 'top_left_cell': 'B2'}}, actual={'freeze': {'mode': None, 'top_left_cell': None}} |
| freeze_panes | pandas | freeze_d5 | write | data_mismatch | error | Expected values did not match actual values: expected={'freeze': {'mode': 'freeze', 'top_left_cell': 'D5'}}, actual={'freeze': {'mode': None, 'top_left_cell': None}} |
| freeze_panes | pandas | split_2x1 | write | data_mismatch | error | Expected values did not match actual values: expected={'freeze': {'mode': 'split', 'x_split': 1, 'y_split': 2}}, actual={'freeze': {'mode': None, 'x_split': None, 'y_split': None}} |
| freeze_panes | openpyxl-readonly | freeze_b2 | read | data_mismatch | error | Expected values did not match actual values: expected={'freeze': {'mode': 'freeze', 'top_left_cell': 'B2'}}, actual={'freeze': {'mode': None, 'top_left_cell': None}} |
| freeze_panes | openpyxl-readonly | freeze_d5 | read | data_mismatch | error | Expected values did not match actual values: expected={'freeze': {'mode': 'freeze', 'top_left_cell': 'D5'}}, actual={'freeze': {'mode': None, 'top_left_cell': None}} |
| freeze_panes | openpyxl-readonly | split_2x1 | read | data_mismatch | error | Expected values did not match actual values: expected={'freeze': {'mode': 'split', 'x_split': 1, 'y_split': 2}}, actual={'freeze': {'mode': None, 'x_split': None, 'y_split': None}} |
| freeze_panes | polars | freeze_b2 | read | data_mismatch | error | Expected values did not match actual values: expected={'freeze': {'mode': 'freeze', 'top_left_cell': 'B2'}}, actual={'freeze': {'mode': None, 'top_left_cell': None}} |
| freeze_panes | polars | freeze_d5 | read | data_mismatch | error | Expected values did not match actual values: expected={'freeze': {'mode': 'freeze', 'top_left_cell': 'D5'}}, actual={'freeze': {'mode': None, 'top_left_cell': None}} |
| freeze_panes | polars | split_2x1 | read | data_mismatch | error | Expected values did not match actual values: expected={'freeze': {'mode': 'split', 'x_split': 1, 'y_split': 2}}, actual={'freeze': {'mode': None, 'x_split': None, 'y_split': None}} |
| freeze_panes | tablib | freeze_b2 | read | data_mismatch | error | Expected values did not match actual values: expected={'freeze': {'mode': 'freeze', 'top_left_cell': 'B2'}}, actual={'freeze': {'mode': None, 'top_left_cell': None}} |
| freeze_panes | tablib | freeze_d5 | read | data_mismatch | error | Expected values did not match actual values: expected={'freeze': {'mode': 'freeze', 'top_left_cell': 'D5'}}, actual={'freeze': {'mode': None, 'top_left_cell': None}} |
| freeze_panes | tablib | split_2x1 | read | data_mismatch | error | Expected values did not match actual values: expected={'freeze': {'mode': 'split', 'x_split': 1, 'y_split': 2}}, actual={'freeze': {'mode': None, 'x_split': None, 'y_split': None}} |
| freeze_panes | tablib | freeze_b2 | write | data_mismatch | error | Expected values did not match actual values: expected={'freeze': {'mode': 'freeze', 'top_left_cell': 'B2'}}, actual={'freeze': {'mode': None, 'top_left_cell': None}} |
| freeze_panes | tablib | freeze_d5 | write | data_mismatch | error | Expected values did not match actual values: expected={'freeze': {'mode': 'freeze', 'top_left_cell': 'D5'}}, actual={'freeze': {'mode': None, 'top_left_cell': None}} |
| freeze_panes | tablib | split_2x1 | write | data_mismatch | error | Expected values did not match actual values: expected={'freeze': {'mode': 'split', 'x_split': 1, 'y_split': 2}}, actual={'freeze': {'mode': None, 'x_split': None, 'y_split': None}} |
| named_ranges | xlsxwriter | nr_simple_cell | write | unsupported_feature | warning | NotImplementedError: xlsxwriter does not implement named range writes |
| named_ranges | xlsxwriter | nr_cell_range | write | unsupported_feature | warning | NotImplementedError: xlsxwriter does not implement named range writes |
| named_ranges | xlsxwriter | nr_formula_ref | write | unsupported_feature | warning | NotImplementedError: xlsxwriter does not implement named range writes |
| named_ranges | xlsxwriter | nr_sheet_scope | write | unsupported_feature | warning | NotImplementedError: xlsxwriter does not implement named range writes |
| named_ranges | xlsxwriter | nr_cross_sheet | write | unsupported_feature | warning | NotImplementedError: xlsxwriter does not implement named range writes |
| named_ranges | xlsxwriter | nr_special_chars | write | unsupported_feature | warning | NotImplementedError: xlsxwriter does not implement named range writes |
| named_ranges | python-calamine | nr_simple_cell | read | unsupported_feature | warning | NotImplementedError: python-calamine does not implement named range reads |
| named_ranges | python-calamine | nr_cell_range | read | unsupported_feature | warning | NotImplementedError: python-calamine does not implement named range reads |
| named_ranges | python-calamine | nr_formula_ref | read | unsupported_feature | warning | NotImplementedError: python-calamine does not implement named range reads |
| named_ranges | python-calamine | nr_sheet_scope | read | unsupported_feature | warning | NotImplementedError: python-calamine does not implement named range reads |
| named_ranges | python-calamine | nr_cross_sheet | read | unsupported_feature | warning | NotImplementedError: python-calamine does not implement named range reads |
| named_ranges | python-calamine | nr_special_chars | read | unsupported_feature | warning | NotImplementedError: python-calamine does not implement named range reads |
| named_ranges | pylightxl | nr_simple_cell | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| named_ranges | pylightxl | nr_cell_range | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| named_ranges | pylightxl | nr_formula_ref | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| named_ranges | pylightxl | nr_sheet_scope | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| named_ranges | pylightxl | nr_cross_sheet | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| named_ranges | pylightxl | nr_special_chars | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| named_ranges | pylightxl | nr_simple_cell | write | unsupported_feature | warning | NotImplementedError: pylightxl does not implement named range writes |
| named_ranges | pylightxl | nr_cell_range | write | unsupported_feature | warning | NotImplementedError: pylightxl does not implement named range writes |
| named_ranges | pylightxl | nr_formula_ref | write | unsupported_feature | warning | NotImplementedError: pylightxl does not implement named range writes |
| named_ranges | pylightxl | nr_sheet_scope | write | unsupported_feature | warning | NotImplementedError: pylightxl does not implement named range writes |
| named_ranges | pylightxl | nr_cross_sheet | write | unsupported_feature | warning | NotImplementedError: pylightxl does not implement named range writes |
| named_ranges | pylightxl | nr_special_chars | write | unsupported_feature | warning | NotImplementedError: pylightxl does not implement named range writes |
| named_ranges | pyexcel | nr_simple_cell | read | unsupported_feature | warning | NotImplementedError: pyexcel does not implement named range reads |
| named_ranges | pyexcel | nr_cell_range | read | unsupported_feature | warning | NotImplementedError: pyexcel does not implement named range reads |
| named_ranges | pyexcel | nr_formula_ref | read | unsupported_feature | warning | NotImplementedError: pyexcel does not implement named range reads |
| named_ranges | pyexcel | nr_sheet_scope | read | unsupported_feature | warning | NotImplementedError: pyexcel does not implement named range reads |
| named_ranges | pyexcel | nr_cross_sheet | read | unsupported_feature | warning | NotImplementedError: pyexcel does not implement named range reads |
| named_ranges | pyexcel | nr_special_chars | read | unsupported_feature | warning | NotImplementedError: pyexcel does not implement named range reads |
| named_ranges | pyexcel | nr_simple_cell | write | unsupported_feature | warning | NotImplementedError: pyexcel does not implement named range writes |
| named_ranges | pyexcel | nr_cell_range | write | unsupported_feature | warning | NotImplementedError: pyexcel does not implement named range writes |
| named_ranges | pyexcel | nr_formula_ref | write | unsupported_feature | warning | NotImplementedError: pyexcel does not implement named range writes |
| named_ranges | pyexcel | nr_sheet_scope | write | unsupported_feature | warning | NotImplementedError: pyexcel does not implement named range writes |
| named_ranges | pyexcel | nr_cross_sheet | write | unsupported_feature | warning | NotImplementedError: pyexcel does not implement named range writes |
| named_ranges | pyexcel | nr_special_chars | write | unsupported_feature | warning | NotImplementedError: pyexcel does not implement named range writes |
| named_ranges | xlwt | nr_simple_cell | write | unsupported_feature | warning | NotImplementedError: xlwt does not implement named range writes |
| named_ranges | xlwt | nr_cell_range | write | unsupported_feature | warning | NotImplementedError: xlwt does not implement named range writes |
| named_ranges | xlwt | nr_formula_ref | write | unsupported_feature | warning | NotImplementedError: xlwt does not implement named range writes |
| named_ranges | xlwt | nr_sheet_scope | write | unsupported_feature | warning | NotImplementedError: xlwt does not implement named range writes |
| named_ranges | xlwt | nr_cross_sheet | write | unsupported_feature | warning | NotImplementedError: xlwt does not implement named range writes |
| named_ranges | xlwt | nr_special_chars | write | unsupported_feature | warning | NotImplementedError: xlwt does not implement named range writes |
| named_ranges | pandas | nr_simple_cell | read | unsupported_feature | warning | NotImplementedError: pandas does not implement named range reads |
| named_ranges | pandas | nr_cell_range | read | unsupported_feature | warning | NotImplementedError: pandas does not implement named range reads |
| named_ranges | pandas | nr_formula_ref | read | unsupported_feature | warning | NotImplementedError: pandas does not implement named range reads |
| named_ranges | pandas | nr_sheet_scope | read | unsupported_feature | warning | NotImplementedError: pandas does not implement named range reads |
| named_ranges | pandas | nr_cross_sheet | read | unsupported_feature | warning | NotImplementedError: pandas does not implement named range reads |
| named_ranges | pandas | nr_special_chars | read | unsupported_feature | warning | NotImplementedError: pandas does not implement named range reads |
| named_ranges | pandas | nr_simple_cell | write | unsupported_feature | warning | NotImplementedError: pandas does not implement named range writes |
| named_ranges | pandas | nr_cell_range | write | unsupported_feature | warning | NotImplementedError: pandas does not implement named range writes |
| named_ranges | pandas | nr_formula_ref | write | unsupported_feature | warning | NotImplementedError: pandas does not implement named range writes |
| named_ranges | pandas | nr_sheet_scope | write | unsupported_feature | warning | NotImplementedError: pandas does not implement named range writes |
| named_ranges | pandas | nr_cross_sheet | write | unsupported_feature | warning | NotImplementedError: pandas does not implement named range writes |
| named_ranges | pandas | nr_special_chars | write | unsupported_feature | warning | NotImplementedError: pandas does not implement named range writes |
| named_ranges | xlsxwriter-constmem | nr_simple_cell | write | unsupported_feature | warning | NotImplementedError: xlsxwriter-constmem does not implement named range writes |
| named_ranges | xlsxwriter-constmem | nr_cell_range | write | unsupported_feature | warning | NotImplementedError: xlsxwriter-constmem does not implement named range writes |
| named_ranges | xlsxwriter-constmem | nr_formula_ref | write | unsupported_feature | warning | NotImplementedError: xlsxwriter-constmem does not implement named range writes |
| named_ranges | xlsxwriter-constmem | nr_sheet_scope | write | unsupported_feature | warning | NotImplementedError: xlsxwriter-constmem does not implement named range writes |
| named_ranges | xlsxwriter-constmem | nr_cross_sheet | write | unsupported_feature | warning | NotImplementedError: xlsxwriter-constmem does not implement named range writes |
| named_ranges | xlsxwriter-constmem | nr_special_chars | write | unsupported_feature | warning | NotImplementedError: xlsxwriter-constmem does not implement named range writes |
| named_ranges | openpyxl-readonly | nr_simple_cell | read | unsupported_feature | warning | NotImplementedError: openpyxl-readonly does not implement named range reads |
| named_ranges | openpyxl-readonly | nr_cell_range | read | unsupported_feature | warning | NotImplementedError: openpyxl-readonly does not implement named range reads |
| named_ranges | openpyxl-readonly | nr_formula_ref | read | unsupported_feature | warning | NotImplementedError: openpyxl-readonly does not implement named range reads |
| named_ranges | openpyxl-readonly | nr_sheet_scope | read | unsupported_feature | warning | NotImplementedError: openpyxl-readonly does not implement named range reads |
| named_ranges | openpyxl-readonly | nr_cross_sheet | read | unsupported_feature | warning | NotImplementedError: openpyxl-readonly does not implement named range reads |
| named_ranges | openpyxl-readonly | nr_special_chars | read | unsupported_feature | warning | NotImplementedError: openpyxl-readonly does not implement named range reads |
| named_ranges | polars | nr_simple_cell | read | unsupported_feature | warning | NotImplementedError: polars does not implement named range reads |
| named_ranges | polars | nr_cell_range | read | unsupported_feature | warning | NotImplementedError: polars does not implement named range reads |
| named_ranges | polars | nr_formula_ref | read | unsupported_feature | warning | NotImplementedError: polars does not implement named range reads |
| named_ranges | polars | nr_sheet_scope | read | unsupported_feature | warning | NotImplementedError: polars does not implement named range reads |
| named_ranges | polars | nr_cross_sheet | read | unsupported_feature | warning | NotImplementedError: polars does not implement named range reads |
| named_ranges | polars | nr_special_chars | read | unsupported_feature | warning | NotImplementedError: polars does not implement named range reads |
| named_ranges | tablib | nr_simple_cell | read | unsupported_feature | warning | NotImplementedError: tablib does not implement named range reads |
| named_ranges | tablib | nr_cell_range | read | unsupported_feature | warning | NotImplementedError: tablib does not implement named range reads |
| named_ranges | tablib | nr_formula_ref | read | unsupported_feature | warning | NotImplementedError: tablib does not implement named range reads |
| named_ranges | tablib | nr_sheet_scope | read | unsupported_feature | warning | NotImplementedError: tablib does not implement named range reads |
| named_ranges | tablib | nr_cross_sheet | read | unsupported_feature | warning | NotImplementedError: tablib does not implement named range reads |
| named_ranges | tablib | nr_special_chars | read | unsupported_feature | warning | NotImplementedError: tablib does not implement named range reads |
| named_ranges | tablib | nr_simple_cell | write | unsupported_feature | warning | NotImplementedError: tablib does not implement named range writes |
| named_ranges | tablib | nr_cell_range | write | unsupported_feature | warning | NotImplementedError: tablib does not implement named range writes |
| named_ranges | tablib | nr_formula_ref | write | unsupported_feature | warning | NotImplementedError: tablib does not implement named range writes |
| named_ranges | tablib | nr_sheet_scope | write | unsupported_feature | warning | NotImplementedError: tablib does not implement named range writes |
| named_ranges | tablib | nr_cross_sheet | write | unsupported_feature | warning | NotImplementedError: tablib does not implement named range writes |
| named_ranges | tablib | nr_special_chars | write | unsupported_feature | warning | NotImplementedError: tablib does not implement named range writes |
| tables | xlsxwriter | tbl_basic | write | unsupported_feature | warning | NotImplementedError: xlsxwriter does not implement table writes |
| tables | xlsxwriter | tbl_with_totals | write | unsupported_feature | warning | NotImplementedError: xlsxwriter does not implement table writes |
| tables | xlsxwriter | tbl_no_style | write | unsupported_feature | warning | NotImplementedError: xlsxwriter does not implement table writes |
| tables | xlsxwriter | tbl_single_col | write | unsupported_feature | warning | NotImplementedError: xlsxwriter does not implement table writes |
| tables | xlsxwriter | tbl_single_row | write | unsupported_feature | warning | NotImplementedError: xlsxwriter does not implement table writes |
| tables | xlsxwriter | tbl_autofilter | write | unsupported_feature | warning | NotImplementedError: xlsxwriter does not implement table writes |
| tables | python-calamine | tbl_basic | read | unsupported_feature | warning | NotImplementedError: python-calamine does not implement table reads |
| tables | python-calamine | tbl_with_totals | read | unsupported_feature | warning | NotImplementedError: python-calamine does not implement table reads |
| tables | python-calamine | tbl_no_style | read | unsupported_feature | warning | NotImplementedError: python-calamine does not implement table reads |
| tables | python-calamine | tbl_single_col | read | unsupported_feature | warning | NotImplementedError: python-calamine does not implement table reads |
| tables | python-calamine | tbl_single_row | read | unsupported_feature | warning | NotImplementedError: python-calamine does not implement table reads |
| tables | python-calamine | tbl_autofilter | read | unsupported_feature | warning | NotImplementedError: python-calamine does not implement table reads |
| tables | pylightxl | tbl_basic | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| tables | pylightxl | tbl_with_totals | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| tables | pylightxl | tbl_no_style | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| tables | pylightxl | tbl_single_col | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| tables | pylightxl | tbl_single_row | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| tables | pylightxl | tbl_autofilter | read | invalid_input | error | TypeError: expected string or bytes-like object, got 'NoneType' |
| tables | pylightxl | tbl_basic | write | unsupported_feature | warning | NotImplementedError: pylightxl does not implement table writes |
| tables | pylightxl | tbl_with_totals | write | unsupported_feature | warning | NotImplementedError: pylightxl does not implement table writes |
| tables | pylightxl | tbl_no_style | write | unsupported_feature | warning | NotImplementedError: pylightxl does not implement table writes |
| tables | pylightxl | tbl_single_col | write | unsupported_feature | warning | NotImplementedError: pylightxl does not implement table writes |
| tables | pylightxl | tbl_single_row | write | unsupported_feature | warning | NotImplementedError: pylightxl does not implement table writes |
| tables | pylightxl | tbl_autofilter | write | unsupported_feature | warning | NotImplementedError: pylightxl does not implement table writes |
| tables | pyexcel | tbl_basic | read | unsupported_feature | warning | NotImplementedError: pyexcel does not implement table reads |
| tables | pyexcel | tbl_with_totals | read | unsupported_feature | warning | NotImplementedError: pyexcel does not implement table reads |
| tables | pyexcel | tbl_no_style | read | unsupported_feature | warning | NotImplementedError: pyexcel does not implement table reads |
| tables | pyexcel | tbl_single_col | read | unsupported_feature | warning | NotImplementedError: pyexcel does not implement table reads |
| tables | pyexcel | tbl_single_row | read | unsupported_feature | warning | NotImplementedError: pyexcel does not implement table reads |
| tables | pyexcel | tbl_autofilter | read | unsupported_feature | warning | NotImplementedError: pyexcel does not implement table reads |
| tables | pyexcel | tbl_basic | write | unsupported_feature | warning | NotImplementedError: pyexcel does not implement table writes |
| tables | pyexcel | tbl_with_totals | write | unsupported_feature | warning | NotImplementedError: pyexcel does not implement table writes |
| tables | pyexcel | tbl_no_style | write | unsupported_feature | warning | NotImplementedError: pyexcel does not implement table writes |
| tables | pyexcel | tbl_single_col | write | unsupported_feature | warning | NotImplementedError: pyexcel does not implement table writes |
| tables | pyexcel | tbl_single_row | write | unsupported_feature | warning | NotImplementedError: pyexcel does not implement table writes |
| tables | pyexcel | tbl_autofilter | write | unsupported_feature | warning | NotImplementedError: pyexcel does not implement table writes |
| tables | xlwt | tbl_basic | write | unsupported_feature | warning | NotImplementedError: xlwt does not implement table writes |
| tables | xlwt | tbl_with_totals | write | unsupported_feature | warning | NotImplementedError: xlwt does not implement table writes |
| tables | xlwt | tbl_no_style | write | unsupported_feature | warning | NotImplementedError: xlwt does not implement table writes |
| tables | xlwt | tbl_single_col | write | unsupported_feature | warning | NotImplementedError: xlwt does not implement table writes |
| tables | xlwt | tbl_single_row | write | unsupported_feature | warning | NotImplementedError: xlwt does not implement table writes |
| tables | xlwt | tbl_autofilter | write | unsupported_feature | warning | NotImplementedError: xlwt does not implement table writes |
| tables | pandas | tbl_basic | read | unsupported_feature | warning | NotImplementedError: pandas does not implement table reads |
| tables | pandas | tbl_with_totals | read | unsupported_feature | warning | NotImplementedError: pandas does not implement table reads |
| tables | pandas | tbl_no_style | read | unsupported_feature | warning | NotImplementedError: pandas does not implement table reads |
| tables | pandas | tbl_single_col | read | unsupported_feature | warning | NotImplementedError: pandas does not implement table reads |
| tables | pandas | tbl_single_row | read | unsupported_feature | warning | NotImplementedError: pandas does not implement table reads |
| tables | pandas | tbl_autofilter | read | unsupported_feature | warning | NotImplementedError: pandas does not implement table reads |
| tables | pandas | tbl_basic | write | unsupported_feature | warning | NotImplementedError: pandas does not implement table writes |
| tables | pandas | tbl_with_totals | write | unsupported_feature | warning | NotImplementedError: pandas does not implement table writes |
| tables | pandas | tbl_no_style | write | unsupported_feature | warning | NotImplementedError: pandas does not implement table writes |
| tables | pandas | tbl_single_col | write | unsupported_feature | warning | NotImplementedError: pandas does not implement table writes |
| tables | pandas | tbl_single_row | write | unsupported_feature | warning | NotImplementedError: pandas does not implement table writes |
| tables | pandas | tbl_autofilter | write | unsupported_feature | warning | NotImplementedError: pandas does not implement table writes |
| tables | xlsxwriter-constmem | tbl_basic | write | unsupported_feature | warning | NotImplementedError: xlsxwriter-constmem does not implement table writes |
| tables | xlsxwriter-constmem | tbl_with_totals | write | unsupported_feature | warning | NotImplementedError: xlsxwriter-constmem does not implement table writes |
| tables | xlsxwriter-constmem | tbl_no_style | write | unsupported_feature | warning | NotImplementedError: xlsxwriter-constmem does not implement table writes |
| tables | xlsxwriter-constmem | tbl_single_col | write | unsupported_feature | warning | NotImplementedError: xlsxwriter-constmem does not implement table writes |
| tables | xlsxwriter-constmem | tbl_single_row | write | unsupported_feature | warning | NotImplementedError: xlsxwriter-constmem does not implement table writes |
| tables | xlsxwriter-constmem | tbl_autofilter | write | unsupported_feature | warning | NotImplementedError: xlsxwriter-constmem does not implement table writes |
| tables | openpyxl-readonly | tbl_basic | read | unsupported_feature | warning | NotImplementedError: openpyxl-readonly does not implement table reads |
| tables | openpyxl-readonly | tbl_with_totals | read | unsupported_feature | warning | NotImplementedError: openpyxl-readonly does not implement table reads |
| tables | openpyxl-readonly | tbl_no_style | read | unsupported_feature | warning | NotImplementedError: openpyxl-readonly does not implement table reads |
| tables | openpyxl-readonly | tbl_single_col | read | unsupported_feature | warning | NotImplementedError: openpyxl-readonly does not implement table reads |
| tables | openpyxl-readonly | tbl_single_row | read | unsupported_feature | warning | NotImplementedError: openpyxl-readonly does not implement table reads |
| tables | openpyxl-readonly | tbl_autofilter | read | unsupported_feature | warning | NotImplementedError: openpyxl-readonly does not implement table reads |
| tables | polars | tbl_basic | read | unsupported_feature | warning | NotImplementedError: polars does not implement table reads |
| tables | polars | tbl_with_totals | read | unsupported_feature | warning | NotImplementedError: polars does not implement table reads |
| tables | polars | tbl_no_style | read | unsupported_feature | warning | NotImplementedError: polars does not implement table reads |
| tables | polars | tbl_single_col | read | unsupported_feature | warning | NotImplementedError: polars does not implement table reads |
| tables | polars | tbl_single_row | read | unsupported_feature | warning | NotImplementedError: polars does not implement table reads |
| tables | polars | tbl_autofilter | read | unsupported_feature | warning | NotImplementedError: polars does not implement table reads |
| tables | tablib | tbl_basic | read | unsupported_feature | warning | NotImplementedError: tablib does not implement table reads |
| tables | tablib | tbl_with_totals | read | unsupported_feature | warning | NotImplementedError: tablib does not implement table reads |
| tables | tablib | tbl_no_style | read | unsupported_feature | warning | NotImplementedError: tablib does not implement table reads |
| tables | tablib | tbl_single_col | read | unsupported_feature | warning | NotImplementedError: tablib does not implement table reads |
| tables | tablib | tbl_single_row | read | unsupported_feature | warning | NotImplementedError: tablib does not implement table reads |
| tables | tablib | tbl_autofilter | read | unsupported_feature | warning | NotImplementedError: tablib does not implement table reads |
| tables | tablib | tbl_basic | write | unsupported_feature | warning | NotImplementedError: tablib does not implement table writes |
| tables | tablib | tbl_with_totals | write | unsupported_feature | warning | NotImplementedError: tablib does not implement table writes |
| tables | tablib | tbl_no_style | write | unsupported_feature | warning | NotImplementedError: tablib does not implement table writes |
| tables | tablib | tbl_single_col | write | unsupported_feature | warning | NotImplementedError: tablib does not implement table writes |
| tables | tablib | tbl_single_row | write | unsupported_feature | warning | NotImplementedError: tablib does not implement table writes |
| tables | tablib | tbl_autofilter | write | unsupported_feature | warning | NotImplementedError: tablib does not implement table writes |

## Detailed Results

<a id="alignment-details"></a>
### alignment

**openpyxl** — Read: 🟢 3 | Write: 🟢 3

**openpyxl-readonly** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Align - left | basic | ❌ |
| Align - center | basic | ❌ |
| Align - right | basic | ❌ |
| Align - top | basic | ❌ |
| Align - center | basic | ❌ |
| Align - bottom | basic | ❌ |
| Align - wrap text | basic | ❌ |
| Align - rotation 45 | basic | ❌ |
| Align - indent 2 | basic | ❌ |

**pandas** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Align - left | basic | ❌ | ❌ |
| Align - center | basic | ❌ | ❌ |
| Align - right | basic | ❌ | ❌ |
| Align - top | basic | ❌ | ❌ |
| Align - center | basic | ❌ | ❌ |
| Align - bottom | basic | ❌ | ❌ |
| Align - wrap text | basic | ❌ | ❌ |
| Align - rotation 45 | basic | ❌ | ❌ |
| Align - indent 2 | basic | ❌ | ❌ |

**polars** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Align - left | basic | ❌ |
| Align - center | basic | ❌ |
| Align - right | basic | ❌ |
| Align - top | basic | ❌ |
| Align - center | basic | ❌ |
| Align - bottom | basic | ❌ |
| Align - wrap text | basic | ❌ |
| Align - rotation 45 | basic | ❌ |
| Align - indent 2 | basic | ❌ |

**pyexcel** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Align - left | basic | ❌ | ❌ |
| Align - center | basic | ❌ | ❌ |
| Align - right | basic | ❌ | ❌ |
| Align - top | basic | ❌ | ❌ |
| Align - center | basic | ❌ | ❌ |
| Align - bottom | basic | ❌ | ❌ |
| Align - wrap text | basic | ❌ | ❌ |
| Align - rotation 45 | basic | ❌ | ❌ |
| Align - indent 2 | basic | ❌ | ❌ |

**pylightxl** — Read: 🔴 0 | Write: 🔴 0
- Notes: Known limitation: pylightxl alignment write is a no-op because the library does not support formatting writes.

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Align - left | basic | ❌ | ❌ |
| Align - center | basic | ❌ | ❌ |
| Align - right | basic | ❌ | ❌ |
| Align - top | basic | ❌ | ❌ |
| Align - center | basic | ❌ | ❌ |
| Align - bottom | basic | ❌ | ❌ |
| Align - wrap text | basic | ❌ | ❌ |
| Align - rotation 45 | basic | ❌ | ❌ |
| Align - indent 2 | basic | ❌ | ❌ |

**python-calamine** — Read: 🔴 0
- Notes: Known limitation: python-calamine alignment read is limited because its API does not expose style/alignment metadata.

| Test | Importance | Read |
|------|-----------|------|
| Align - left | basic | ❌ |
| Align - center | basic | ❌ |
| Align - right | basic | ❌ |
| Align - top | basic | ❌ |
| Align - center | basic | ❌ |
| Align - bottom | basic | ❌ |
| Align - wrap text | basic | ❌ |
| Align - rotation 45 | basic | ❌ |
| Align - indent 2 | basic | ❌ |

**rust_xlsxwriter** — Write: 🟠 1

| Test | Importance | Write |
|------|-----------|-------|
| Align - left | basic | ✅ |
| Align - center | basic | ✅ |
| Align - right | basic | ✅ |
| Align - top | basic | ✅ |
| Align - center | basic | ✅ |
| Align - bottom | basic | ❌ |
| Align - wrap text | basic | ✅ |
| Align - rotation 45 | basic | ✅ |
| Align - indent 2 | basic | ✅ |

**tablib** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Align - left | basic | ❌ | ❌ |
| Align - center | basic | ❌ | ❌ |
| Align - right | basic | ❌ | ❌ |
| Align - top | basic | ❌ | ❌ |
| Align - center | basic | ❌ | ❌ |
| Align - bottom | basic | ❌ | ❌ |
| Align - wrap text | basic | ❌ | ❌ |
| Align - rotation 45 | basic | ❌ | ❌ |
| Align - indent 2 | basic | ❌ | ❌ |

**wolfxl** — Read: 🟢 3 | Write: 🟠 1

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Align - left | basic | ✅ | ✅ |
| Align - center | basic | ✅ | ✅ |
| Align - right | basic | ✅ | ✅ |
| Align - top | basic | ✅ | ✅ |
| Align - center | basic | ✅ | ✅ |
| Align - bottom | basic | ✅ | ❌ |
| Align - wrap text | basic | ✅ | ✅ |
| Align - rotation 45 | basic | ✅ | ✅ |
| Align - indent 2 | basic | ✅ | ✅ |

**xlrd**
- Notes: Not applicable: xlrd does not support .xlsx input

**xlsxwriter** — Write: 🟠 1

| Test | Importance | Write |
|------|-----------|-------|
| Align - left | basic | ✅ |
| Align - center | basic | ✅ |
| Align - right | basic | ✅ |
| Align - top | basic | ✅ |
| Align - center | basic | ✅ |
| Align - bottom | basic | ❌ |
| Align - wrap text | basic | ✅ |
| Align - rotation 45 | basic | ✅ |
| Align - indent 2 | basic | ✅ |

**xlsxwriter-constmem** — Write: 🟠 1

| Test | Importance | Write |
|------|-----------|-------|
| Align - left | basic | ✅ |
| Align - center | basic | ✅ |
| Align - right | basic | ✅ |
| Align - top | basic | ✅ |
| Align - center | basic | ✅ |
| Align - bottom | basic | ❌ |
| Align - wrap text | basic | ✅ |
| Align - rotation 45 | basic | ✅ |
| Align - indent 2 | basic | ✅ |

**xlwt** — Write: 🟢 3

<a id="background_colors-details"></a>
### background_colors

**openpyxl** — Read: 🟢 3 | Write: 🟢 3

**openpyxl-readonly** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Background - red | basic | ❌ |
| Background - blue | basic | ❌ |
| Background - green | basic | ❌ |
| Background - custom (#8B4513) | basic | ❌ |

**pandas** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Background - red | basic | ❌ | ❌ |
| Background - blue | basic | ❌ | ❌ |
| Background - green | basic | ❌ | ❌ |
| Background - custom (#8B4513) | basic | ❌ | ❌ |

**polars** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Background - red | basic | ❌ |
| Background - blue | basic | ❌ |
| Background - green | basic | ❌ |
| Background - custom (#8B4513) | basic | ❌ |

**pyexcel** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Background - red | basic | ❌ | ❌ |
| Background - blue | basic | ❌ | ❌ |
| Background - green | basic | ❌ | ❌ |
| Background - custom (#8B4513) | basic | ❌ | ❌ |

**pylightxl** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Background - red | basic | ❌ | ❌ |
| Background - blue | basic | ❌ | ❌ |
| Background - green | basic | ❌ | ❌ |
| Background - custom (#8B4513) | basic | ❌ | ❌ |

**python-calamine** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Background - red | basic | ❌ |
| Background - blue | basic | ❌ |
| Background - green | basic | ❌ |
| Background - custom (#8B4513) | basic | ❌ |

**rust_xlsxwriter** — Write: 🟢 3

**tablib** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Background - red | basic | ❌ | ❌ |
| Background - blue | basic | ❌ | ❌ |
| Background - green | basic | ❌ | ❌ |
| Background - custom (#8B4513) | basic | ❌ | ❌ |

**wolfxl** — Read: 🟢 3 | Write: 🟢 3

**xlrd**
- Notes: Not applicable: xlrd does not support .xlsx input

**xlsxwriter** — Write: 🟢 3

**xlsxwriter-constmem** — Write: 🟢 3

**xlwt** — Write: 🟠 1

| Test | Importance | Write |
|------|-----------|-------|
| Background - red | basic | ✅ |
| Background - blue | basic | ✅ |
| Background - green | basic | ❌ |
| Background - custom (#8B4513) | basic | ❌ |

<a id="borders-details"></a>
### borders

**openpyxl** — Read: 🟢 3 | Write: 🟢 3

**openpyxl-readonly** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Border - thin all edges | basic | ❌ |
| Border - medium all edges | basic | ❌ |
| Border - thick all edges | basic | ❌ |
| Border - double line | basic | ❌ |
| Border - dashed | basic | ❌ |
| Border - dotted | basic | ❌ |
| Border - dash-dot | basic | ❌ |
| Border - dash-dot-dot | basic | ❌ |
| Border - top only | basic | ❌ |
| Border - bottom only | basic | ❌ |
| Border - left only | basic | ❌ |
| Border - right only | basic | ❌ |
| Border - diagonal up | basic | ❌ |
| Border - diagonal down | basic | ❌ |
| Border - diagonal both | basic | ❌ |
| Border - red color | basic | ❌ |
| Border - blue color | basic | ❌ |
| Border - custom color (#8B4513) | basic | ❌ |
| Border - mixed styles per edge | basic | ❌ |
| Border - mixed colors per edge | basic | ❌ |

**pandas** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Border - thin all edges | basic | ❌ | ❌ |
| Border - medium all edges | basic | ❌ | ❌ |
| Border - thick all edges | basic | ❌ | ❌ |
| Border - double line | basic | ❌ | ❌ |
| Border - dashed | basic | ❌ | ❌ |
| Border - dotted | basic | ❌ | ❌ |
| Border - dash-dot | basic | ❌ | ❌ |
| Border - dash-dot-dot | basic | ❌ | ❌ |
| Border - top only | basic | ❌ | ❌ |
| Border - bottom only | basic | ❌ | ❌ |
| Border - left only | basic | ❌ | ❌ |
| Border - right only | basic | ❌ | ❌ |
| Border - diagonal up | basic | ❌ | ❌ |
| Border - diagonal down | basic | ❌ | ❌ |
| Border - diagonal both | basic | ❌ | ❌ |
| Border - red color | basic | ❌ | ❌ |
| Border - blue color | basic | ❌ | ❌ |
| Border - custom color (#8B4513) | basic | ❌ | ❌ |
| Border - mixed styles per edge | basic | ❌ | ❌ |
| Border - mixed colors per edge | basic | ❌ | ❌ |

**polars** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Border - thin all edges | basic | ❌ |
| Border - medium all edges | basic | ❌ |
| Border - thick all edges | basic | ❌ |
| Border - double line | basic | ❌ |
| Border - dashed | basic | ❌ |
| Border - dotted | basic | ❌ |
| Border - dash-dot | basic | ❌ |
| Border - dash-dot-dot | basic | ❌ |
| Border - top only | basic | ❌ |
| Border - bottom only | basic | ❌ |
| Border - left only | basic | ❌ |
| Border - right only | basic | ❌ |
| Border - diagonal up | basic | ❌ |
| Border - diagonal down | basic | ❌ |
| Border - diagonal both | basic | ❌ |
| Border - red color | basic | ❌ |
| Border - blue color | basic | ❌ |
| Border - custom color (#8B4513) | basic | ❌ |
| Border - mixed styles per edge | basic | ❌ |
| Border - mixed colors per edge | basic | ❌ |

**pyexcel** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Border - thin all edges | basic | ❌ | ❌ |
| Border - medium all edges | basic | ❌ | ❌ |
| Border - thick all edges | basic | ❌ | ❌ |
| Border - double line | basic | ❌ | ❌ |
| Border - dashed | basic | ❌ | ❌ |
| Border - dotted | basic | ❌ | ❌ |
| Border - dash-dot | basic | ❌ | ❌ |
| Border - dash-dot-dot | basic | ❌ | ❌ |
| Border - top only | basic | ❌ | ❌ |
| Border - bottom only | basic | ❌ | ❌ |
| Border - left only | basic | ❌ | ❌ |
| Border - right only | basic | ❌ | ❌ |
| Border - diagonal up | basic | ❌ | ❌ |
| Border - diagonal down | basic | ❌ | ❌ |
| Border - diagonal both | basic | ❌ | ❌ |
| Border - red color | basic | ❌ | ❌ |
| Border - blue color | basic | ❌ | ❌ |
| Border - custom color (#8B4513) | basic | ❌ | ❌ |
| Border - mixed styles per edge | basic | ❌ | ❌ |
| Border - mixed colors per edge | basic | ❌ | ❌ |

**pylightxl** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Border - thin all edges | basic | ❌ | ❌ |
| Border - medium all edges | basic | ❌ | ❌ |
| Border - thick all edges | basic | ❌ | ❌ |
| Border - double line | basic | ❌ | ❌ |
| Border - dashed | basic | ❌ | ❌ |
| Border - dotted | basic | ❌ | ❌ |
| Border - dash-dot | basic | ❌ | ❌ |
| Border - dash-dot-dot | basic | ❌ | ❌ |
| Border - top only | basic | ❌ | ❌ |
| Border - bottom only | basic | ❌ | ❌ |
| Border - left only | basic | ❌ | ❌ |
| Border - right only | basic | ❌ | ❌ |
| Border - diagonal up | basic | ❌ | ❌ |
| Border - diagonal down | basic | ❌ | ❌ |
| Border - diagonal both | basic | ❌ | ❌ |
| Border - red color | basic | ❌ | ❌ |
| Border - blue color | basic | ❌ | ❌ |
| Border - custom color (#8B4513) | basic | ❌ | ❌ |
| Border - mixed styles per edge | basic | ❌ | ❌ |
| Border - mixed colors per edge | basic | ❌ | ❌ |

**python-calamine** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Border - thin all edges | basic | ❌ |
| Border - medium all edges | basic | ❌ |
| Border - thick all edges | basic | ❌ |
| Border - double line | basic | ❌ |
| Border - dashed | basic | ❌ |
| Border - dotted | basic | ❌ |
| Border - dash-dot | basic | ❌ |
| Border - dash-dot-dot | basic | ❌ |
| Border - top only | basic | ❌ |
| Border - bottom only | basic | ❌ |
| Border - left only | basic | ❌ |
| Border - right only | basic | ❌ |
| Border - diagonal up | basic | ❌ |
| Border - diagonal down | basic | ❌ |
| Border - diagonal both | basic | ❌ |
| Border - red color | basic | ❌ |
| Border - blue color | basic | ❌ |
| Border - custom color (#8B4513) | basic | ❌ |
| Border - mixed styles per edge | basic | ❌ |
| Border - mixed colors per edge | basic | ❌ |

**rust_xlsxwriter** — Write: 🟢 3

**tablib** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Border - thin all edges | basic | ❌ | ❌ |
| Border - medium all edges | basic | ❌ | ❌ |
| Border - thick all edges | basic | ❌ | ❌ |
| Border - double line | basic | ❌ | ❌ |
| Border - dashed | basic | ❌ | ❌ |
| Border - dotted | basic | ❌ | ❌ |
| Border - dash-dot | basic | ❌ | ❌ |
| Border - dash-dot-dot | basic | ❌ | ❌ |
| Border - top only | basic | ❌ | ❌ |
| Border - bottom only | basic | ❌ | ❌ |
| Border - left only | basic | ❌ | ❌ |
| Border - right only | basic | ❌ | ❌ |
| Border - diagonal up | basic | ❌ | ❌ |
| Border - diagonal down | basic | ❌ | ❌ |
| Border - diagonal both | basic | ❌ | ❌ |
| Border - red color | basic | ❌ | ❌ |
| Border - blue color | basic | ❌ | ❌ |
| Border - custom color (#8B4513) | basic | ❌ | ❌ |
| Border - mixed styles per edge | basic | ❌ | ❌ |
| Border - mixed colors per edge | basic | ❌ | ❌ |

**wolfxl** — Read: 🟢 3 | Write: 🟢 3

**xlrd**
- Notes: Not applicable: xlrd does not support .xlsx input

**xlsxwriter** — Write: 🟢 3

**xlsxwriter-constmem** — Write: 🟢 3

**xlwt** — Write: 🟠 1

| Test | Importance | Write |
|------|-----------|-------|
| Border - thin all edges | basic | ✅ |
| Border - medium all edges | basic | ✅ |
| Border - thick all edges | basic | ✅ |
| Border - double line | basic | ✅ |
| Border - dashed | basic | ✅ |
| Border - dotted | basic | ✅ |
| Border - dash-dot | basic | ✅ |
| Border - dash-dot-dot | basic | ✅ |
| Border - top only | basic | ✅ |
| Border - bottom only | basic | ✅ |
| Border - left only | basic | ✅ |
| Border - right only | basic | ✅ |
| Border - diagonal up | basic | ❌ |
| Border - diagonal down | basic | ❌ |
| Border - diagonal both | basic | ✅ |
| Border - red color | basic | ✅ |
| Border - blue color | basic | ✅ |
| Border - custom color (#8B4513) | basic | ❌ |
| Border - mixed styles per edge | basic | ✅ |
| Border - mixed colors per edge | basic | ❌ |

<a id="cell_values-details"></a>
### cell_values

**openpyxl** — Read: 🟢 3 | Write: 🟢 3

**openpyxl-readonly** — Read: 🟢 3

**pandas** — Read: 🟠 1 | Write: 🟢 3

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| String - simple | basic | ✅ | ✅ |
| String - unicode | basic | ✅ | ✅ |
| String - empty | basic | ✅ | ✅ |
| String - long (1000 chars) | basic | ✅ | ✅ |
| String - with newlines | basic | ✅ | ✅ |
| Number - integer | basic | ✅ | ✅ |
| Number - float | basic | ✅ | ✅ |
| Number - negative | basic | ✅ | ✅ |
| Number - large | basic | ✅ | ✅ |
| Number - scientific notation | basic | ✅ | ✅ |
| Date - standard | basic | ✅ | ✅ |
| DateTime - with time | basic | ✅ | ✅ |
| Boolean - TRUE | basic | ✅ | ✅ |
| Boolean - FALSE | basic | ✅ | ✅ |
| Error - #DIV/0! | basic | ❌ | ✅ |
| Error - #N/A | basic | ❌ | ✅ |
| Error - #VALUE! | basic | ❌ | ✅ |
| Blank cell | basic | ✅ | ✅ |

**polars** — Read: 🟠 1

| Test | Importance | Read |
|------|-----------|------|
| String - simple | basic | ✅ |
| String - unicode | basic | ✅ |
| String - empty | basic | ✅ |
| String - long (1000 chars) | basic | ✅ |
| String - with newlines | basic | ✅ |
| Number - integer | basic | ✅ |
| Number - float | basic | ✅ |
| Number - negative | basic | ✅ |
| Number - large | basic | ✅ |
| Number - scientific notation | basic | ✅ |
| Date - standard | basic | ✅ |
| DateTime - with time | basic | ✅ |
| Boolean - TRUE | basic | ✅ |
| Boolean - FALSE | basic | ✅ |
| Error - #DIV/0! | basic | ❌ |
| Error - #N/A | basic | ❌ |
| Error - #VALUE! | basic | ❌ |
| Blank cell | basic | ✅ |

**pyexcel** — Read: 🟢 3 | Write: 🟢 3

**pylightxl** — Read: 🟢 3 | Write: 🟠 1
- Notes: Known limitation: pylightxl cell-values write has date/boolean/error fidelity limits due to writer encoding behavior.

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| String - simple | basic | ✅ | ✅ |
| String - unicode | basic | ✅ | ✅ |
| String - empty | basic | ✅ | ✅ |
| String - long (1000 chars) | basic | ✅ | ✅ |
| String - with newlines | basic | ✅ | ✅ |
| Number - integer | basic | ✅ | ✅ |
| Number - float | basic | ✅ | ✅ |
| Number - negative | basic | ✅ | ✅ |
| Number - large | basic | ✅ | ✅ |
| Number - scientific notation | basic | ✅ | ✅ |
| Date - standard | basic | ✅ | ❌ |
| DateTime - with time | basic | ✅ | ❌ |
| Boolean - TRUE | basic | ✅ | ❌ |
| Boolean - FALSE | basic | ✅ | ❌ |
| Error - #DIV/0! | basic | ✅ | ✅ |
| Error - #N/A | basic | ✅ | ✅ |
| Error - #VALUE! | basic | ✅ | ✅ |
| Blank cell | basic | ✅ | ✅ |

**python-calamine** — Read: 🟠 1
- Notes: Known limitation: python-calamine can surface formula error cells as blank values in current API responses.

| Test | Importance | Read |
|------|-----------|------|
| String - simple | basic | ✅ |
| String - unicode | basic | ✅ |
| String - empty | basic | ✅ |
| String - long (1000 chars) | basic | ✅ |
| String - with newlines | basic | ✅ |
| Number - integer | basic | ✅ |
| Number - float | basic | ✅ |
| Number - negative | basic | ✅ |
| Number - large | basic | ✅ |
| Number - scientific notation | basic | ✅ |
| Date - standard | basic | ✅ |
| DateTime - with time | basic | ✅ |
| Boolean - TRUE | basic | ✅ |
| Boolean - FALSE | basic | ✅ |
| Error - #DIV/0! | basic | ❌ |
| Error - #N/A | basic | ❌ |
| Error - #VALUE! | basic | ❌ |
| Blank cell | basic | ✅ |

**rust_xlsxwriter** — Write: 🟢 3

**tablib** — Read: 🟢 3 | Write: 🟢 3

**wolfxl** — Read: 🟢 3 | Write: 🟢 3

**xlrd**
- Notes: Not applicable: xlrd does not support .xlsx input

**xlsxwriter** — Write: 🟢 3

**xlsxwriter-constmem** — Write: 🟢 3

**xlwt** — Write: 🟢 3

<a id="comments-details"></a>
### comments

**openpyxl** — Read: 🟢 3 | Write: 🟢 3

**openpyxl-readonly** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Comment: legacy note | basic | ❌ |
| Comment: threaded | edge | ❌ |
| Comment: second author | edge | ❌ |

**pandas** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Comment: legacy note | basic | ❌ | ❌ |
| Comment: threaded | edge | ❌ | ❌ |
| Comment: second author | edge | ❌ | ❌ |

**polars** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Comment: legacy note | basic | ❌ |
| Comment: threaded | edge | ❌ |
| Comment: second author | edge | ❌ |

**pyexcel** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Comment: legacy note | basic | ❌ | ❌ |
| Comment: threaded | edge | ❌ | ❌ |
| Comment: second author | edge | ❌ | ❌ |

**pylightxl** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Comment: legacy note | basic | ❌ | ❌ |
| Comment: threaded | edge | ❌ | ❌ |
| Comment: second author | edge | ❌ | ❌ |

**python-calamine** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Comment: legacy note | basic | ❌ |
| Comment: threaded | edge | ❌ |
| Comment: second author | edge | ❌ |

**rust_xlsxwriter** — Write: 🟢 3

**tablib** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Comment: legacy note | basic | ❌ | ❌ |
| Comment: threaded | edge | ❌ | ❌ |
| Comment: second author | edge | ❌ | ❌ |

**wolfxl** — Read: 🟢 3 | Write: 🟢 3

**xlrd**
- Notes: Not applicable: xlrd does not support .xlsx input

**xlsxwriter** — Write: 🟢 3

**xlsxwriter-constmem** — Write: 🔴 0

| Test | Importance | Write |
|------|-----------|-------|
| Comment: legacy note | basic | ❌ |
| Comment: threaded | edge | ❌ |
| Comment: second author | edge | ❌ |

**xlwt** — Write: 🔴 0

| Test | Importance | Write |
|------|-----------|-------|
| Comment: legacy note | basic | ❌ |
| Comment: threaded | edge | ❌ |
| Comment: second author | edge | ❌ |

<a id="conditional_formatting-details"></a>
### conditional_formatting

**openpyxl** — Read: 🟢 3 | Write: 🟢 3

**openpyxl-readonly** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| CF: cell > 5 (yellow fill) | basic | ❌ |
| CF: formula rule with cross-sheet ref | edge | ❌ |
| CF: text contains | edge | ❌ |
| CF: data bar | edge | ❌ |
| CF: 3-color scale | edge | ❌ |
| CF: stop-if-true priority | edge | ❌ |

**pandas** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| CF: cell > 5 (yellow fill) | basic | ❌ | ❌ |
| CF: formula rule with cross-sheet ref | edge | ❌ | ❌ |
| CF: text contains | edge | ❌ | ❌ |
| CF: data bar | edge | ❌ | ❌ |
| CF: 3-color scale | edge | ❌ | ❌ |
| CF: stop-if-true priority | edge | ❌ | ❌ |

**polars** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| CF: cell > 5 (yellow fill) | basic | ❌ |
| CF: formula rule with cross-sheet ref | edge | ❌ |
| CF: text contains | edge | ❌ |
| CF: data bar | edge | ❌ |
| CF: 3-color scale | edge | ❌ |
| CF: stop-if-true priority | edge | ❌ |

**pyexcel** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| CF: cell > 5 (yellow fill) | basic | ❌ | ❌ |
| CF: formula rule with cross-sheet ref | edge | ❌ | ❌ |
| CF: text contains | edge | ❌ | ❌ |
| CF: data bar | edge | ❌ | ❌ |
| CF: 3-color scale | edge | ❌ | ❌ |
| CF: stop-if-true priority | edge | ❌ | ❌ |

**pylightxl** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| CF: cell > 5 (yellow fill) | basic | ❌ | ❌ |
| CF: formula rule with cross-sheet ref | edge | ❌ | ❌ |
| CF: text contains | edge | ❌ | ❌ |
| CF: data bar | edge | ❌ | ❌ |
| CF: 3-color scale | edge | ❌ | ❌ |
| CF: stop-if-true priority | edge | ❌ | ❌ |

**python-calamine** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| CF: cell > 5 (yellow fill) | basic | ❌ |
| CF: formula rule with cross-sheet ref | edge | ❌ |
| CF: text contains | edge | ❌ |
| CF: data bar | edge | ❌ |
| CF: 3-color scale | edge | ❌ |
| CF: stop-if-true priority | edge | ❌ |

**rust_xlsxwriter** — Write: 🟢 3

**tablib** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| CF: cell > 5 (yellow fill) | basic | ❌ | ❌ |
| CF: formula rule with cross-sheet ref | edge | ❌ | ❌ |
| CF: text contains | edge | ❌ | ❌ |
| CF: data bar | edge | ❌ | ❌ |
| CF: 3-color scale | edge | ❌ | ❌ |
| CF: stop-if-true priority | edge | ❌ | ❌ |

**wolfxl** — Read: 🟢 3 | Write: 🟢 3

**xlrd**
- Notes: Not applicable: xlrd does not support .xlsx input

**xlsxwriter** — Write: 🟢 3

**xlsxwriter-constmem** — Write: 🟢 3

**xlwt** — Write: 🔴 0

| Test | Importance | Write |
|------|-----------|-------|
| CF: cell > 5 (yellow fill) | basic | ❌ |
| CF: formula rule with cross-sheet ref | edge | ❌ |
| CF: text contains | edge | ❌ |
| CF: data bar | edge | ❌ |
| CF: 3-color scale | edge | ❌ |
| CF: stop-if-true priority | edge | ❌ |

<a id="data_validation-details"></a>
### data_validation

**openpyxl** — Read: 🟢 3 | Write: 🟢 3

**openpyxl-readonly** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| DV: list from CSV | basic | ❌ |
| DV: list from range | edge | ❌ |
| DV: cross-sheet named range | edge | ❌ |
| DV: custom formula | edge | ❌ |
| DV: whole number with error | basic | ❌ |

**pandas** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| DV: list from CSV | basic | ❌ | ❌ |
| DV: list from range | edge | ❌ | ❌ |
| DV: cross-sheet named range | edge | ❌ | ❌ |
| DV: custom formula | edge | ❌ | ❌ |
| DV: whole number with error | basic | ❌ | ❌ |

**polars** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| DV: list from CSV | basic | ❌ |
| DV: list from range | edge | ❌ |
| DV: cross-sheet named range | edge | ❌ |
| DV: custom formula | edge | ❌ |
| DV: whole number with error | basic | ❌ |

**pyexcel** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| DV: list from CSV | basic | ❌ | ❌ |
| DV: list from range | edge | ❌ | ❌ |
| DV: cross-sheet named range | edge | ❌ | ❌ |
| DV: custom formula | edge | ❌ | ❌ |
| DV: whole number with error | basic | ❌ | ❌ |

**pylightxl** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| DV: list from CSV | basic | ❌ | ❌ |
| DV: list from range | edge | ❌ | ❌ |
| DV: cross-sheet named range | edge | ❌ | ❌ |
| DV: custom formula | edge | ❌ | ❌ |
| DV: whole number with error | basic | ❌ | ❌ |

**python-calamine** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| DV: list from CSV | basic | ❌ |
| DV: list from range | edge | ❌ |
| DV: cross-sheet named range | edge | ❌ |
| DV: custom formula | edge | ❌ |
| DV: whole number with error | basic | ❌ |

**rust_xlsxwriter** — Write: 🟢 3

**tablib** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| DV: list from CSV | basic | ❌ | ❌ |
| DV: list from range | edge | ❌ | ❌ |
| DV: cross-sheet named range | edge | ❌ | ❌ |
| DV: custom formula | edge | ❌ | ❌ |
| DV: whole number with error | basic | ❌ | ❌ |

**wolfxl** — Read: 🟢 3 | Write: 🟢 3

**xlrd**
- Notes: Not applicable: xlrd does not support .xlsx input

**xlsxwriter** — Write: 🟢 3

**xlsxwriter-constmem** — Write: 🟢 3

**xlwt** — Write: 🔴 0

| Test | Importance | Write |
|------|-----------|-------|
| DV: list from CSV | basic | ❌ |
| DV: list from range | edge | ❌ |
| DV: cross-sheet named range | edge | ❌ |
| DV: custom formula | edge | ❌ |
| DV: whole number with error | basic | ❌ |

<a id="dimensions-details"></a>
### dimensions

**openpyxl** — Read: 🟢 3 | Write: 🟢 3

**openpyxl-readonly** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Row height - 30 | basic | ❌ |
| Row height - 45 | basic | ❌ |
| Column width - D = 20 | basic | ❌ |
| Column width - E = 8 | basic | ❌ |

**pandas** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Row height - 30 | basic | ❌ | ❌ |
| Row height - 45 | basic | ❌ | ❌ |
| Column width - D = 20 | basic | ❌ | ❌ |
| Column width - E = 8 | basic | ❌ | ❌ |

**polars** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Row height - 30 | basic | ❌ |
| Row height - 45 | basic | ❌ |
| Column width - D = 20 | basic | ❌ |
| Column width - E = 8 | basic | ❌ |

**pyexcel** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Row height - 30 | basic | ❌ | ❌ |
| Row height - 45 | basic | ❌ | ❌ |
| Column width - D = 20 | basic | ❌ | ❌ |
| Column width - E = 8 | basic | ❌ | ❌ |

**pylightxl** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Row height - 30 | basic | ❌ | ❌ |
| Row height - 45 | basic | ❌ | ❌ |
| Column width - D = 20 | basic | ❌ | ❌ |
| Column width - E = 8 | basic | ❌ | ❌ |

**python-calamine** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Row height - 30 | basic | ❌ |
| Row height - 45 | basic | ❌ |
| Column width - D = 20 | basic | ❌ |
| Column width - E = 8 | basic | ❌ |

**rust_xlsxwriter** — Write: 🟢 3

**tablib** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Row height - 30 | basic | ❌ | ❌ |
| Row height - 45 | basic | ❌ | ❌ |
| Column width - D = 20 | basic | ❌ | ❌ |
| Column width - E = 8 | basic | ❌ | ❌ |

**wolfxl** — Read: 🟢 3 | Write: 🟢 3

**xlrd**
- Notes: Not applicable: xlrd does not support .xlsx input

**xlsxwriter** — Write: 🟢 3

**xlsxwriter-constmem** — Write: 🟠 1

| Test | Importance | Write |
|------|-----------|-------|
| Row height - 30 | basic | ❌ |
| Row height - 45 | basic | ❌ |
| Column width - D = 20 | basic | ✅ |
| Column width - E = 8 | basic | ✅ |

**xlwt** — Write: 🟠 1

| Test | Importance | Write |
|------|-----------|-------|
| Row height - 30 | basic | ❌ |
| Row height - 45 | basic | ❌ |
| Column width - D = 20 | basic | ✅ |
| Column width - E = 8 | basic | ✅ |

<a id="formulas-details"></a>
### formulas

**openpyxl** — Read: 🟢 3 | Write: 🟢 3

**openpyxl-readonly** — Read: 🟢 3

**pandas** — Read: 🔴 0 | Write: 🟢 3

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Formula - SUM | basic | ❌ | ✅ |
| Formula - cell reference | basic | ❌ | ✅ |
| Formula - concat | basic | ❌ | ✅ |
| Formula - cross sheet | basic | ❌ | ✅ |

**polars** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Formula - SUM | basic | ❌ |
| Formula - cell reference | basic | ❌ |
| Formula - concat | basic | ❌ |
| Formula - cross sheet | basic | ❌ |

**pyexcel** — Read: 🔴 0 | Write: 🟢 3

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Formula - SUM | basic | ❌ | ✅ |
| Formula - cell reference | basic | ❌ | ✅ |
| Formula - concat | basic | ❌ | ✅ |
| Formula - cross sheet | basic | ❌ | ✅ |

**pylightxl** — Read: 🔴 0 | Write: 🟢 3

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Formula - SUM | basic | ❌ | ✅ |
| Formula - cell reference | basic | ❌ | ✅ |
| Formula - concat | basic | ❌ | ✅ |
| Formula - cross sheet | basic | ❌ | ✅ |

**python-calamine** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Formula - SUM | basic | ❌ |
| Formula - cell reference | basic | ❌ |
| Formula - concat | basic | ❌ |
| Formula - cross sheet | basic | ❌ |

**rust_xlsxwriter** — Write: 🟢 3

**tablib** — Read: 🔴 0 | Write: 🟢 3

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Formula - SUM | basic | ❌ | ✅ |
| Formula - cell reference | basic | ❌ | ✅ |
| Formula - concat | basic | ❌ | ✅ |
| Formula - cross sheet | basic | ❌ | ✅ |

**wolfxl** — Read: 🟢 3 | Write: 🟢 3

**xlrd**
- Notes: Not applicable: xlrd does not support .xlsx input

**xlsxwriter** — Write: 🟢 3

**xlsxwriter-constmem** — Write: 🟢 3

**xlwt** — Write: 🔴 0

| Test | Importance | Write |
|------|-----------|-------|
| Formula - SUM | basic | ❌ |
| Formula - cell reference | basic | ❌ |
| Formula - concat | basic | ❌ |
| Formula - cross sheet | basic | ❌ |

<a id="freeze_panes-details"></a>
### freeze_panes

**openpyxl** — Read: 🟢 3 | Write: 🟢 3

**openpyxl-readonly** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Freeze panes at B2 | basic | ❌ |
| Freeze panes at D5 | edge | ❌ |
| Split panes row=2 col=1 | edge | ❌ |

**pandas** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Freeze panes at B2 | basic | ❌ | ❌ |
| Freeze panes at D5 | edge | ❌ | ❌ |
| Split panes row=2 col=1 | edge | ❌ | ❌ |

**polars** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Freeze panes at B2 | basic | ❌ |
| Freeze panes at D5 | edge | ❌ |
| Split panes row=2 col=1 | edge | ❌ |

**pyexcel** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Freeze panes at B2 | basic | ❌ | ❌ |
| Freeze panes at D5 | edge | ❌ | ❌ |
| Split panes row=2 col=1 | edge | ❌ | ❌ |

**pylightxl** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Freeze panes at B2 | basic | ❌ | ❌ |
| Freeze panes at D5 | edge | ❌ | ❌ |
| Split panes row=2 col=1 | edge | ❌ | ❌ |

**python-calamine** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Freeze panes at B2 | basic | ❌ |
| Freeze panes at D5 | edge | ❌ |
| Split panes row=2 col=1 | edge | ❌ |

**rust_xlsxwriter** — Write: 🟢 3

**tablib** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Freeze panes at B2 | basic | ❌ | ❌ |
| Freeze panes at D5 | edge | ❌ | ❌ |
| Split panes row=2 col=1 | edge | ❌ | ❌ |

**wolfxl** — Read: 🟢 3 | Write: 🟢 3

**xlrd**
- Notes: Not applicable: xlrd does not support .xlsx input

**xlsxwriter** — Write: 🟢 3

**xlsxwriter-constmem** — Write: 🟢 3

**xlwt** — Write: 🔴 0

| Test | Importance | Write |
|------|-----------|-------|
| Freeze panes at B2 | basic | ❌ |
| Freeze panes at D5 | edge | ❌ |
| Split panes row=2 col=1 | edge | ❌ |

<a id="hyperlinks-details"></a>
### hyperlinks

**openpyxl** — Read: 🟢 3 | Write: 🟢 3

**openpyxl-readonly** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Hyperlink: external URL | basic | ❌ |
| Hyperlink: internal sheet | edge | ❌ |
| Hyperlink: mailto | basic | ❌ |
| Hyperlink: long encoded URL | edge | ❌ |

**pandas** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Hyperlink: external URL | basic | ❌ | ❌ |
| Hyperlink: internal sheet | edge | ❌ | ❌ |
| Hyperlink: mailto | basic | ❌ | ❌ |
| Hyperlink: long encoded URL | edge | ❌ | ❌ |

**polars** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Hyperlink: external URL | basic | ❌ |
| Hyperlink: internal sheet | edge | ❌ |
| Hyperlink: mailto | basic | ❌ |
| Hyperlink: long encoded URL | edge | ❌ |

**pyexcel** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Hyperlink: external URL | basic | ❌ | ❌ |
| Hyperlink: internal sheet | edge | ❌ | ❌ |
| Hyperlink: mailto | basic | ❌ | ❌ |
| Hyperlink: long encoded URL | edge | ❌ | ❌ |

**pylightxl** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Hyperlink: external URL | basic | ❌ | ❌ |
| Hyperlink: internal sheet | edge | ❌ | ❌ |
| Hyperlink: mailto | basic | ❌ | ❌ |
| Hyperlink: long encoded URL | edge | ❌ | ❌ |

**python-calamine** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Hyperlink: external URL | basic | ❌ |
| Hyperlink: internal sheet | edge | ❌ |
| Hyperlink: mailto | basic | ❌ |
| Hyperlink: long encoded URL | edge | ❌ |

**rust_xlsxwriter** — Write: 🟢 3

**tablib** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Hyperlink: external URL | basic | ❌ | ❌ |
| Hyperlink: internal sheet | edge | ❌ | ❌ |
| Hyperlink: mailto | basic | ❌ | ❌ |
| Hyperlink: long encoded URL | edge | ❌ | ❌ |

**wolfxl** — Read: 🟢 3 | Write: 🟢 3

**xlrd**
- Notes: Not applicable: xlrd does not support .xlsx input

**xlsxwriter** — Write: 🟢 3

**xlsxwriter-constmem** — Write: 🟢 3

**xlwt** — Write: 🔴 0

| Test | Importance | Write |
|------|-----------|-------|
| Hyperlink: external URL | basic | ❌ |
| Hyperlink: internal sheet | edge | ❌ |
| Hyperlink: mailto | basic | ❌ |
| Hyperlink: long encoded URL | edge | ❌ |

<a id="images-details"></a>
### images

**openpyxl** — Read: 🟢 3 | Write: 🟢 3

**openpyxl-readonly** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Image: one-cell anchor | basic | ❌ |
| Image: two-cell anchor with offset | edge | ❌ |

**pandas** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Image: one-cell anchor | basic | ❌ | ❌ |
| Image: two-cell anchor with offset | edge | ❌ | ❌ |

**polars** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Image: one-cell anchor | basic | ❌ |
| Image: two-cell anchor with offset | edge | ❌ |

**pyexcel** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Image: one-cell anchor | basic | ❌ | ❌ |
| Image: two-cell anchor with offset | edge | ❌ | ❌ |

**pylightxl** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Image: one-cell anchor | basic | ❌ | ❌ |
| Image: two-cell anchor with offset | edge | ❌ | ❌ |

**python-calamine** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Image: one-cell anchor | basic | ❌ |
| Image: two-cell anchor with offset | edge | ❌ |

**rust_xlsxwriter** — Write: 🔴 0

| Test | Importance | Write |
|------|-----------|-------|
| Image: one-cell anchor | basic | ❌ |
| Image: two-cell anchor with offset | edge | ❌ |

**tablib** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Image: one-cell anchor | basic | ❌ | ❌ |
| Image: two-cell anchor with offset | edge | ❌ | ❌ |

**wolfxl** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Image: one-cell anchor | basic | ❌ | ❌ |
| Image: two-cell anchor with offset | edge | ❌ | ❌ |

**xlrd**
- Notes: Not applicable: xlrd does not support .xlsx input

**xlsxwriter** — Write: 🟢 3

**xlsxwriter-constmem** — Write: 🔴 0

| Test | Importance | Write |
|------|-----------|-------|
| Image: one-cell anchor | basic | ❌ |
| Image: two-cell anchor with offset | edge | ❌ |

**xlwt** — Write: 🔴 0

| Test | Importance | Write |
|------|-----------|-------|
| Image: one-cell anchor | basic | ❌ |
| Image: two-cell anchor with offset | edge | ❌ |

<a id="merged_cells-details"></a>
### merged_cells

**openpyxl** — Read: 🟢 3 | Write: 🟢 3

**openpyxl-readonly** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Merge horizontal B2:D2 | basic | ❌ |
| Merge vertical B3:B5 | basic | ❌ |
| Merge with non-top-left value | edge | ❌ |
| Merge with top-left fill | edge | ❌ |

**pandas** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Merge horizontal B2:D2 | basic | ❌ | ❌ |
| Merge vertical B3:B5 | basic | ❌ | ❌ |
| Merge with non-top-left value | edge | ❌ | ❌ |
| Merge with top-left fill | edge | ❌ | ❌ |

**polars** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Merge horizontal B2:D2 | basic | ❌ |
| Merge vertical B3:B5 | basic | ❌ |
| Merge with non-top-left value | edge | ❌ |
| Merge with top-left fill | edge | ❌ |

**pyexcel** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Merge horizontal B2:D2 | basic | ❌ | ❌ |
| Merge vertical B3:B5 | basic | ❌ | ❌ |
| Merge with non-top-left value | edge | ❌ | ❌ |
| Merge with top-left fill | edge | ❌ | ❌ |

**pylightxl** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Merge horizontal B2:D2 | basic | ❌ | ❌ |
| Merge vertical B3:B5 | basic | ❌ | ❌ |
| Merge with non-top-left value | edge | ❌ | ❌ |
| Merge with top-left fill | edge | ❌ | ❌ |

**python-calamine** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Merge horizontal B2:D2 | basic | ❌ |
| Merge vertical B3:B5 | basic | ❌ |
| Merge with non-top-left value | edge | ❌ |
| Merge with top-left fill | edge | ❌ |

**rust_xlsxwriter** — Write: 🟢 3

**tablib** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Merge horizontal B2:D2 | basic | ❌ | ❌ |
| Merge vertical B3:B5 | basic | ❌ | ❌ |
| Merge with non-top-left value | edge | ❌ | ❌ |
| Merge with top-left fill | edge | ❌ | ❌ |

**wolfxl** — Read: 🟢 3 | Write: 🟢 3

**xlrd**
- Notes: Not applicable: xlrd does not support .xlsx input

**xlsxwriter** — Write: 🟢 3

**xlsxwriter-constmem** — Write: 🟢 3

**xlwt** — Write: 🔴 0

| Test | Importance | Write |
|------|-----------|-------|
| Merge horizontal B2:D2 | basic | ❌ |
| Merge vertical B3:B5 | basic | ❌ |
| Merge with non-top-left value | edge | ❌ |
| Merge with top-left fill | edge | ❌ |

<a id="multiple_sheets-details"></a>
### multiple_sheets

**openpyxl** — Read: 🟢 3 | Write: 🟢 3

**openpyxl-readonly** — Read: 🟢 3

**pandas** — Read: 🟢 3 | Write: 🟢 3

**polars** — Read: 🟠 1

| Test | Importance | Read |
|------|-----------|------|
| Sheet names | basic | ✅ |
| Alpha value | basic | ✅ |
| Beta value | basic | ❌ |
| Gamma value | basic | ❌ |

**pyexcel** — Read: 🟢 3 | Write: 🟢 3

**pylightxl** — Read: 🟢 3 | Write: 🟢 3

**python-calamine** — Read: 🟢 3

**rust_xlsxwriter** — Write: 🟢 3

**tablib** — Read: 🟢 3 | Write: 🟢 3

**wolfxl** — Read: 🟢 3 | Write: 🟢 3

**xlrd**
- Notes: Not applicable: xlrd does not support .xlsx input

**xlsxwriter** — Write: 🟢 3

**xlsxwriter-constmem** — Write: 🟢 3

**xlwt** — Write: 🟢 3

<a id="named_ranges-details"></a>
### named_ranges

**openpyxl** — Read: 🟢 3 | Write: 🟢 3

**openpyxl-readonly** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Named range: single cell | basic | ❌ |
| Named range: cell range | basic | ❌ |
| Named range: used in formula | basic | ❌ |
| Named range: sheet-scoped | edge | ❌ |
| Named range: cross-sheet reference | edge | ❌ |
| Named range: underscore name | edge | ❌ |

**pandas** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Named range: single cell | basic | ❌ | ❌ |
| Named range: cell range | basic | ❌ | ❌ |
| Named range: used in formula | basic | ❌ | ❌ |
| Named range: sheet-scoped | edge | ❌ | ❌ |
| Named range: cross-sheet reference | edge | ❌ | ❌ |
| Named range: underscore name | edge | ❌ | ❌ |

**polars** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Named range: single cell | basic | ❌ |
| Named range: cell range | basic | ❌ |
| Named range: used in formula | basic | ❌ |
| Named range: sheet-scoped | edge | ❌ |
| Named range: cross-sheet reference | edge | ❌ |
| Named range: underscore name | edge | ❌ |

**pyexcel** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Named range: single cell | basic | ❌ | ❌ |
| Named range: cell range | basic | ❌ | ❌ |
| Named range: used in formula | basic | ❌ | ❌ |
| Named range: sheet-scoped | edge | ❌ | ❌ |
| Named range: cross-sheet reference | edge | ❌ | ❌ |
| Named range: underscore name | edge | ❌ | ❌ |

**pylightxl** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Named range: single cell | basic | ❌ | ❌ |
| Named range: cell range | basic | ❌ | ❌ |
| Named range: used in formula | basic | ❌ | ❌ |
| Named range: sheet-scoped | edge | ❌ | ❌ |
| Named range: cross-sheet reference | edge | ❌ | ❌ |
| Named range: underscore name | edge | ❌ | ❌ |

**python-calamine** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Named range: single cell | basic | ❌ |
| Named range: cell range | basic | ❌ |
| Named range: used in formula | basic | ❌ |
| Named range: sheet-scoped | edge | ❌ |
| Named range: cross-sheet reference | edge | ❌ |
| Named range: underscore name | edge | ❌ |

**rust_xlsxwriter** — Write: 🟢 3

**tablib** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Named range: single cell | basic | ❌ | ❌ |
| Named range: cell range | basic | ❌ | ❌ |
| Named range: used in formula | basic | ❌ | ❌ |
| Named range: sheet-scoped | edge | ❌ | ❌ |
| Named range: cross-sheet reference | edge | ❌ | ❌ |
| Named range: underscore name | edge | ❌ | ❌ |

**wolfxl** — Read: 🟢 3 | Write: 🟢 3

**xlrd**
- Notes: Not applicable: xlrd does not support .xlsx input

**xlsxwriter** — Write: 🔴 0

| Test | Importance | Write |
|------|-----------|-------|
| Named range: single cell | basic | ❌ |
| Named range: cell range | basic | ❌ |
| Named range: used in formula | basic | ❌ |
| Named range: sheet-scoped | edge | ❌ |
| Named range: cross-sheet reference | edge | ❌ |
| Named range: underscore name | edge | ❌ |

**xlsxwriter-constmem** — Write: 🔴 0

| Test | Importance | Write |
|------|-----------|-------|
| Named range: single cell | basic | ❌ |
| Named range: cell range | basic | ❌ |
| Named range: used in formula | basic | ❌ |
| Named range: sheet-scoped | edge | ❌ |
| Named range: cross-sheet reference | edge | ❌ |
| Named range: underscore name | edge | ❌ |

**xlwt** — Write: 🔴 0

| Test | Importance | Write |
|------|-----------|-------|
| Named range: single cell | basic | ❌ |
| Named range: cell range | basic | ❌ |
| Named range: used in formula | basic | ❌ |
| Named range: sheet-scoped | edge | ❌ |
| Named range: cross-sheet reference | edge | ❌ |
| Named range: underscore name | edge | ❌ |

<a id="number_formats-details"></a>
### number_formats

**openpyxl** — Read: 🟢 3 | Write: 🟢 3

**openpyxl-readonly** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Format - currency | basic | ❌ |
| Format - percent | basic | ❌ |
| Format - date | basic | ❌ |
| Format - scientific | basic | ❌ |
| Format - custom text | basic | ❌ |

**pandas** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Format - currency | basic | ❌ | ❌ |
| Format - percent | basic | ❌ | ❌ |
| Format - date | basic | ❌ | ❌ |
| Format - scientific | basic | ❌ | ❌ |
| Format - custom text | basic | ❌ | ❌ |

**polars** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Format - currency | basic | ❌ |
| Format - percent | basic | ❌ |
| Format - date | basic | ❌ |
| Format - scientific | basic | ❌ |
| Format - custom text | basic | ❌ |

**pyexcel** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Format - currency | basic | ❌ | ❌ |
| Format - percent | basic | ❌ | ❌ |
| Format - date | basic | ❌ | ❌ |
| Format - scientific | basic | ❌ | ❌ |
| Format - custom text | basic | ❌ | ❌ |

**pylightxl** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Format - currency | basic | ❌ | ❌ |
| Format - percent | basic | ❌ | ❌ |
| Format - date | basic | ❌ | ❌ |
| Format - scientific | basic | ❌ | ❌ |
| Format - custom text | basic | ❌ | ❌ |

**python-calamine** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Format - currency | basic | ❌ |
| Format - percent | basic | ❌ |
| Format - date | basic | ❌ |
| Format - scientific | basic | ❌ |
| Format - custom text | basic | ❌ |

**rust_xlsxwriter** — Write: 🟢 3

**tablib** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Format - currency | basic | ❌ | ❌ |
| Format - percent | basic | ❌ | ❌ |
| Format - date | basic | ❌ | ❌ |
| Format - scientific | basic | ❌ | ❌ |
| Format - custom text | basic | ❌ | ❌ |

**wolfxl** — Read: 🟢 3 | Write: 🟢 3

**xlrd**
- Notes: Not applicable: xlrd does not support .xlsx input

**xlsxwriter** — Write: 🟢 3

**xlsxwriter-constmem** — Write: 🟢 3

**xlwt** — Write: 🟢 3

<a id="pivot_tables-details"></a>
### pivot_tables

**openpyxl**
- Notes: Unsupported on macOS without a Windows-generated pivot fixture (fixtures/excel/tier2/15_pivot_tables.xlsx).

**openpyxl-readonly**
- Notes: Unsupported on macOS without a Windows-generated pivot fixture (fixtures/excel/tier2/15_pivot_tables.xlsx).

**pandas**
- Notes: Unsupported on macOS without a Windows-generated pivot fixture (fixtures/excel/tier2/15_pivot_tables.xlsx).

**polars**
- Notes: Unsupported on macOS without a Windows-generated pivot fixture (fixtures/excel/tier2/15_pivot_tables.xlsx).

**pyexcel**
- Notes: Unsupported on macOS without a Windows-generated pivot fixture (fixtures/excel/tier2/15_pivot_tables.xlsx).

**pylightxl**
- Notes: Unsupported on macOS without a Windows-generated pivot fixture (fixtures/excel/tier2/15_pivot_tables.xlsx).

**python-calamine**
- Notes: Unsupported on macOS without a Windows-generated pivot fixture (fixtures/excel/tier2/15_pivot_tables.xlsx).

**rust_xlsxwriter**
- Notes: Unsupported on macOS without a Windows-generated pivot fixture (fixtures/excel/tier2/15_pivot_tables.xlsx).

**tablib**
- Notes: Unsupported on macOS without a Windows-generated pivot fixture (fixtures/excel/tier2/15_pivot_tables.xlsx).

**wolfxl**
- Notes: Unsupported on macOS without a Windows-generated pivot fixture (fixtures/excel/tier2/15_pivot_tables.xlsx).

**xlrd**
- Notes: Not applicable: xlrd does not support .xlsx input

**xlsxwriter**
- Notes: Unsupported on macOS without a Windows-generated pivot fixture (fixtures/excel/tier2/15_pivot_tables.xlsx).

**xlsxwriter-constmem**
- Notes: Unsupported on macOS without a Windows-generated pivot fixture (fixtures/excel/tier2/15_pivot_tables.xlsx).

**xlwt**
- Notes: Unsupported on macOS without a Windows-generated pivot fixture (fixtures/excel/tier2/15_pivot_tables.xlsx).

<a id="tables-details"></a>
### tables

**openpyxl** — Read: 🟢 3 | Write: 🟢 3

**openpyxl-readonly** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Table: basic 3-col | basic | ❌ |
| Table: with totals row | basic | ❌ |
| Table: no style | basic | ❌ |
| Table: single column | edge | ❌ |
| Table: header only (no data rows) | edge | ❌ |
| Table: with autoFilter | edge | ❌ |

**pandas** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Table: basic 3-col | basic | ❌ | ❌ |
| Table: with totals row | basic | ❌ | ❌ |
| Table: no style | basic | ❌ | ❌ |
| Table: single column | edge | ❌ | ❌ |
| Table: header only (no data rows) | edge | ❌ | ❌ |
| Table: with autoFilter | edge | ❌ | ❌ |

**polars** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Table: basic 3-col | basic | ❌ |
| Table: with totals row | basic | ❌ |
| Table: no style | basic | ❌ |
| Table: single column | edge | ❌ |
| Table: header only (no data rows) | edge | ❌ |
| Table: with autoFilter | edge | ❌ |

**pyexcel** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Table: basic 3-col | basic | ❌ | ❌ |
| Table: with totals row | basic | ❌ | ❌ |
| Table: no style | basic | ❌ | ❌ |
| Table: single column | edge | ❌ | ❌ |
| Table: header only (no data rows) | edge | ❌ | ❌ |
| Table: with autoFilter | edge | ❌ | ❌ |

**pylightxl** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Table: basic 3-col | basic | ❌ | ❌ |
| Table: with totals row | basic | ❌ | ❌ |
| Table: no style | basic | ❌ | ❌ |
| Table: single column | edge | ❌ | ❌ |
| Table: header only (no data rows) | edge | ❌ | ❌ |
| Table: with autoFilter | edge | ❌ | ❌ |

**python-calamine** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Table: basic 3-col | basic | ❌ |
| Table: with totals row | basic | ❌ |
| Table: no style | basic | ❌ |
| Table: single column | edge | ❌ |
| Table: header only (no data rows) | edge | ❌ |
| Table: with autoFilter | edge | ❌ |

**rust_xlsxwriter** — Write: 🟢 3

**tablib** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Table: basic 3-col | basic | ❌ | ❌ |
| Table: with totals row | basic | ❌ | ❌ |
| Table: no style | basic | ❌ | ❌ |
| Table: single column | edge | ❌ | ❌ |
| Table: header only (no data rows) | edge | ❌ | ❌ |
| Table: with autoFilter | edge | ❌ | ❌ |

**wolfxl** — Read: 🟢 3 | Write: 🟢 3

**xlrd**
- Notes: Not applicable: xlrd does not support .xlsx input

**xlsxwriter** — Write: 🔴 0

| Test | Importance | Write |
|------|-----------|-------|
| Table: basic 3-col | basic | ❌ |
| Table: with totals row | basic | ❌ |
| Table: no style | basic | ❌ |
| Table: single column | edge | ❌ |
| Table: header only (no data rows) | edge | ❌ |
| Table: with autoFilter | edge | ❌ |

**xlsxwriter-constmem** — Write: 🔴 0

| Test | Importance | Write |
|------|-----------|-------|
| Table: basic 3-col | basic | ❌ |
| Table: with totals row | basic | ❌ |
| Table: no style | basic | ❌ |
| Table: single column | edge | ❌ |
| Table: header only (no data rows) | edge | ❌ |
| Table: with autoFilter | edge | ❌ |

**xlwt** — Write: 🔴 0

| Test | Importance | Write |
|------|-----------|-------|
| Table: basic 3-col | basic | ❌ |
| Table: with totals row | basic | ❌ |
| Table: no style | basic | ❌ |
| Table: single column | edge | ❌ |
| Table: header only (no data rows) | edge | ❌ |
| Table: with autoFilter | edge | ❌ |

<a id="text_formatting-details"></a>
### text_formatting

**openpyxl** — Read: 🟢 3 | Write: 🟢 3

**openpyxl-readonly** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Bold | basic | ❌ |
| Italic | basic | ❌ |
| Underline - single | basic | ❌ |
| Underline - double | basic | ❌ |
| Strikethrough | basic | ❌ |
| Bold + Italic | basic | ❌ |
| Font size 8 | basic | ❌ |
| Font size 14 | basic | ❌ |
| Font size 24 | basic | ❌ |
| Font size 36 | basic | ❌ |
| Font - Arial | basic | ❌ |
| Font - Times New Roman | basic | ❌ |
| Font - Courier New | basic | ❌ |
| Font color - red | basic | ❌ |
| Font color - blue | basic | ❌ |
| Font color - green | basic | ❌ |
| Font color - custom (#8B4513) | basic | ❌ |
| Combined - bold, 16pt, red | basic | ❌ |

**pandas** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Bold | basic | ❌ | ❌ |
| Italic | basic | ❌ | ❌ |
| Underline - single | basic | ❌ | ❌ |
| Underline - double | basic | ❌ | ❌ |
| Strikethrough | basic | ❌ | ❌ |
| Bold + Italic | basic | ❌ | ❌ |
| Font size 8 | basic | ❌ | ❌ |
| Font size 14 | basic | ❌ | ❌ |
| Font size 24 | basic | ❌ | ❌ |
| Font size 36 | basic | ❌ | ❌ |
| Font - Arial | basic | ❌ | ❌ |
| Font - Times New Roman | basic | ❌ | ❌ |
| Font - Courier New | basic | ❌ | ❌ |
| Font color - red | basic | ❌ | ❌ |
| Font color - blue | basic | ❌ | ❌ |
| Font color - green | basic | ❌ | ❌ |
| Font color - custom (#8B4513) | basic | ❌ | ❌ |
| Combined - bold, 16pt, red | basic | ❌ | ❌ |

**polars** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Bold | basic | ❌ |
| Italic | basic | ❌ |
| Underline - single | basic | ❌ |
| Underline - double | basic | ❌ |
| Strikethrough | basic | ❌ |
| Bold + Italic | basic | ❌ |
| Font size 8 | basic | ❌ |
| Font size 14 | basic | ❌ |
| Font size 24 | basic | ❌ |
| Font size 36 | basic | ❌ |
| Font - Arial | basic | ❌ |
| Font - Times New Roman | basic | ❌ |
| Font - Courier New | basic | ❌ |
| Font color - red | basic | ❌ |
| Font color - blue | basic | ❌ |
| Font color - green | basic | ❌ |
| Font color - custom (#8B4513) | basic | ❌ |
| Combined - bold, 16pt, red | basic | ❌ |

**pyexcel** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Bold | basic | ❌ | ❌ |
| Italic | basic | ❌ | ❌ |
| Underline - single | basic | ❌ | ❌ |
| Underline - double | basic | ❌ | ❌ |
| Strikethrough | basic | ❌ | ❌ |
| Bold + Italic | basic | ❌ | ❌ |
| Font size 8 | basic | ❌ | ❌ |
| Font size 14 | basic | ❌ | ❌ |
| Font size 24 | basic | ❌ | ❌ |
| Font size 36 | basic | ❌ | ❌ |
| Font - Arial | basic | ❌ | ❌ |
| Font - Times New Roman | basic | ❌ | ❌ |
| Font - Courier New | basic | ❌ | ❌ |
| Font color - red | basic | ❌ | ❌ |
| Font color - blue | basic | ❌ | ❌ |
| Font color - green | basic | ❌ | ❌ |
| Font color - custom (#8B4513) | basic | ❌ | ❌ |
| Combined - bold, 16pt, red | basic | ❌ | ❌ |

**pylightxl** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Bold | basic | ❌ | ❌ |
| Italic | basic | ❌ | ❌ |
| Underline - single | basic | ❌ | ❌ |
| Underline - double | basic | ❌ | ❌ |
| Strikethrough | basic | ❌ | ❌ |
| Bold + Italic | basic | ❌ | ❌ |
| Font size 8 | basic | ❌ | ❌ |
| Font size 14 | basic | ❌ | ❌ |
| Font size 24 | basic | ❌ | ❌ |
| Font size 36 | basic | ❌ | ❌ |
| Font - Arial | basic | ❌ | ❌ |
| Font - Times New Roman | basic | ❌ | ❌ |
| Font - Courier New | basic | ❌ | ❌ |
| Font color - red | basic | ❌ | ❌ |
| Font color - blue | basic | ❌ | ❌ |
| Font color - green | basic | ❌ | ❌ |
| Font color - custom (#8B4513) | basic | ❌ | ❌ |
| Combined - bold, 16pt, red | basic | ❌ | ❌ |

**python-calamine** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Bold | basic | ❌ |
| Italic | basic | ❌ |
| Underline - single | basic | ❌ |
| Underline - double | basic | ❌ |
| Strikethrough | basic | ❌ |
| Bold + Italic | basic | ❌ |
| Font size 8 | basic | ❌ |
| Font size 14 | basic | ❌ |
| Font size 24 | basic | ❌ |
| Font size 36 | basic | ❌ |
| Font - Arial | basic | ❌ |
| Font - Times New Roman | basic | ❌ |
| Font - Courier New | basic | ❌ |
| Font color - red | basic | ❌ |
| Font color - blue | basic | ❌ |
| Font color - green | basic | ❌ |
| Font color - custom (#8B4513) | basic | ❌ |
| Combined - bold, 16pt, red | basic | ❌ |

**rust_xlsxwriter** — Write: 🟢 3

**tablib** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Bold | basic | ❌ | ❌ |
| Italic | basic | ❌ | ❌ |
| Underline - single | basic | ❌ | ❌ |
| Underline - double | basic | ❌ | ❌ |
| Strikethrough | basic | ❌ | ❌ |
| Bold + Italic | basic | ❌ | ❌ |
| Font size 8 | basic | ❌ | ❌ |
| Font size 14 | basic | ❌ | ❌ |
| Font size 24 | basic | ❌ | ❌ |
| Font size 36 | basic | ❌ | ❌ |
| Font - Arial | basic | ❌ | ❌ |
| Font - Times New Roman | basic | ❌ | ❌ |
| Font - Courier New | basic | ❌ | ❌ |
| Font color - red | basic | ❌ | ❌ |
| Font color - blue | basic | ❌ | ❌ |
| Font color - green | basic | ❌ | ❌ |
| Font color - custom (#8B4513) | basic | ❌ | ❌ |
| Combined - bold, 16pt, red | basic | ❌ | ❌ |

**wolfxl** — Read: 🟢 3 | Write: 🟢 3

**xlrd**
- Notes: Not applicable: xlrd does not support .xlsx input

**xlsxwriter** — Write: 🟢 3

**xlsxwriter-constmem** — Write: 🟢 3

**xlwt** — Write: 🟠 1

| Test | Importance | Write |
|------|-----------|-------|
| Bold | basic | ✅ |
| Italic | basic | ✅ |
| Underline - single | basic | ✅ |
| Underline - double | basic | ✅ |
| Strikethrough | basic | ✅ |
| Bold + Italic | basic | ✅ |
| Font size 8 | basic | ✅ |
| Font size 14 | basic | ✅ |
| Font size 24 | basic | ✅ |
| Font size 36 | basic | ✅ |
| Font - Arial | basic | ✅ |
| Font - Times New Roman | basic | ✅ |
| Font - Courier New | basic | ✅ |
| Font color - red | basic | ✅ |
| Font color - blue | basic | ✅ |
| Font color - green | basic | ❌ |
| Font color - custom (#8B4513) | basic | ❌ |
| Combined - bold, 16pt, red | basic | ✅ |

---
*Benchmark version: 0.1.0*