# ExcelBench Results

*Generated: 2026-02-08 22:56 UTC*
*Profile: xlsx*
*Excel Version: 16.105.3*
*Platform: Darwin-arm64*

## Overview

> Condensed view — shows the **best score** across read/write for each library. See [Full Results Matrix](#full-results-matrix) for the complete R/W breakdown.

**Tier 0 — Basic Values**

| Feature | openpyxl | opxl-readonly | pandas | polars | pyexcel | pylightxl | calamine | tablib | xlsxwriter | xlsx-constmem | xlwt |
|---------|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|
| Cell Values | 🟢 | 🟢 | 🟢 | 🟠 | 🟢 | 🟢 | 🟠 | 🟢 | 🟢 | 🟢 | 🟢 |
| Formulas | 🟢 | 🟢 | 🟢 | 🔴 | 🟢 | 🟢 | 🔴 | 🟢 | 🟢 | 🟢 | 🔴 |
| Sheets | 🟢 | 🟢 | 🟢 | 🟠 | 🟢 | 🟢 | 🟢 | 🟢 | 🟢 | 🟢 | 🟢 |

**Tier 1 — Formatting**

| Feature | openpyxl | opxl-readonly | pandas | polars | pyexcel | pylightxl | calamine | tablib | xlsxwriter | xlsx-constmem | xlwt |
|---------|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|
| Alignment | 🟢 | 🟠 | 🟠 | 🟠 | 🟠 | 🟠 | 🟠 | 🟠 | 🟢 | 🟢 | 🟢 |
| Bg Colors | 🟢 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🟢 | 🟢 | 🟠 |
| Borders | 🟢 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🟢 | 🟢 | 🟠 |
| Dimensions | 🟢 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🟢 | 🟠 | 🟠 |
| Num Fmt | 🟢 | 🔴 | 🔴 | 🔴 | 🟠 | 🔴 | 🔴 | 🟠 | 🟢 | 🟢 | 🟢 |
| Text Fmt | 🟢 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🟢 | 🟢 | 🟠 |

**Tier 2 — Advanced**

| Feature | openpyxl | opxl-readonly | pandas | polars | pyexcel | pylightxl | calamine | tablib | xlsxwriter | xlsx-constmem | xlwt |
|---------|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|
| Comments | 🟢 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🟢 | 🔴 | 🔴 |
| Cond Fmt | 🟢 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🟢 | 🟢 | 🔴 |
| Validation | 🟢 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🟢 | 🟢 | 🔴 |
| Freeze | 🟢 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🟢 | 🟢 | 🔴 |
| Hyperlinks | 🟢 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🟢 | 🟢 | 🔴 |
| Images | 🟢 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🟢 | 🔴 | 🔴 |
| Merged | 🟢 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🔴 | 🟢 | 🟢 | 🔴 |

## Library Tiers

> Libraries ranked by their best capability (max of read/write green features).

| Tier | Library | Caps | Green Features | Summary |
|:----:|---------|:----:|:--------------:|---------|
| **S** | openpyxl | R+W | 16/16 | Reference adapter — full read + write fidelity |
| **S** | xlsxwriter | W | 16/16 | Best write-only option — full formatting support |
| **A** | xlsxwriter-constmem | W | 13/16 | Memory-optimized write — loses images, comments, row height |
| **B** | xlwt | W | 4/16 | Legacy .xls writer — basic formatting subset |
| **C** | openpyxl-readonly | R | 3/16 | Streaming read — loses all formatting metadata |
| **C** | pandas | R+W | 3/16 | DataFrame abstraction — errors coerced to NaN on read |
| **C** | pyexcel | R+W | 3/16 | Meta-library wrapping openpyxl — preserves error values |
| **C** | tablib | R+W | 3/16 | Dataset wrapper — matches pyexcel on fidelity |
| **C** | pylightxl | R+W | 2/16 | Lightweight — basic values, no formatting API |
| **C** | python-calamine | R | 1/16 | Fast Rust-backed reader — cell values + sheet names only |
| **D** | polars | R | 0/16 | Rust DataFrame reader — columnar type coercion drops fidelity |

## Score Legend

| Score | Meaning |
|-------|---------|
| 🟢 3 | Complete — full fidelity |
| 🟡 2 | Functional — works for common cases |
| 🟠 1 | Minimal — basic recognition only |
| 🔴 0 | Unsupported — errors or data loss |
| ➖ | Not applicable |

## Full Results Matrix

**Tier 0 — Basic Values**

| Feature | openpyxl (R) | openpyxl (W) | openpyxl-readonly (R) | pandas (R) | pandas (W) | polars (R) | pyexcel (R) | pyexcel (W) | pylightxl (R) | pylightxl (W) | python-calamine (R) | tablib (R) | tablib (W) | xlrd (R) | xlsxwriter (W) | xlsxwriter-constmem (W) | xlwt (W) |
|---------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|
| [cell_values](#cell_values-details) | 🟢 3 | 🟢 3 | 🟢 3 | 🟠 1 | 🟢 3 | 🟠 1 | 🟢 3 | 🟢 3 | 🟢 3 | 🟠 1 | 🟠 1 | 🟢 3 | 🟢 3 | ➖ | 🟢 3 | 🟢 3 | 🟢 3 |
| [formulas](#formulas-details) | 🟢 3 | 🟢 3 | 🟢 3 | 🔴 0 | 🟢 3 | 🔴 0 | 🔴 0 | 🟢 3 | 🔴 0 | 🟢 3 | 🔴 0 | 🔴 0 | 🟢 3 | ➖ | 🟢 3 | 🟢 3 | 🔴 0 |
| [multiple_sheets](#multiple_sheets-details) | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 | 🟠 1 | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 | ➖ | 🟢 3 | 🟢 3 | 🟢 3 |

**Tier 1 — Formatting**

| Feature | openpyxl (R) | openpyxl (W) | openpyxl-readonly (R) | pandas (R) | pandas (W) | polars (R) | pyexcel (R) | pyexcel (W) | pylightxl (R) | pylightxl (W) | python-calamine (R) | tablib (R) | tablib (W) | xlrd (R) | xlsxwriter (W) | xlsxwriter-constmem (W) | xlwt (W) |
|---------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|
| [alignment](#alignment-details) | 🟢 3 | 🟢 3 | 🟠 1 | 🟠 1 | 🟠 1 | 🟠 1 | 🟠 1 | 🟠 1 | 🔴 0 | 🟠 1 | 🟠 1 | 🟠 1 | 🟠 1 | ➖ | 🟢 3 | 🟢 3 | 🟢 3 |
| [background_colors](#background_colors-details) | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | ➖ | 🟢 3 | 🟢 3 | 🟠 1 |
| [borders](#borders-details) | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | ➖ | 🟢 3 | 🟢 3 | 🟠 1 |
| [dimensions](#dimensions-details) | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | ➖ | 🟢 3 | 🟠 1 | 🟠 1 |
| [number_formats](#number_formats-details) | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🟠 1 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🟠 1 | ➖ | 🟢 3 | 🟢 3 | 🟢 3 |
| [text_formatting](#text_formatting-details) | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | ➖ | 🟢 3 | 🟢 3 | 🟠 1 |

**Tier 2 — Advanced**

| Feature | openpyxl (R) | openpyxl (W) | openpyxl-readonly (R) | pandas (R) | pandas (W) | polars (R) | pyexcel (R) | pyexcel (W) | pylightxl (R) | pylightxl (W) | python-calamine (R) | tablib (R) | tablib (W) | xlrd (R) | xlsxwriter (W) | xlsxwriter-constmem (W) | xlwt (W) |
|---------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|------------|
| [comments](#comments-details) | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | ➖ | 🟢 3 | 🔴 0 | 🔴 0 |
| [conditional_formatting](#conditional_formatting-details) | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | ➖ | 🟢 3 | 🟢 3 | 🔴 0 |
| [data_validation](#data_validation-details) | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | ➖ | 🟢 3 | 🟢 3 | 🔴 0 |
| [freeze_panes](#freeze_panes-details) | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | ➖ | 🟢 3 | 🟢 3 | 🔴 0 |
| [hyperlinks](#hyperlinks-details) | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | ➖ | 🟢 3 | 🟢 3 | 🔴 0 |
| [images](#images-details) | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | ➖ | 🟢 3 | 🔴 0 | 🔴 0 |
| [merged_cells](#merged_cells-details) | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | ➖ | 🟢 3 | 🟢 3 | 🔴 0 |
| [pivot_tables](#pivot_tables-details) | ➖ | ➖ | ➖ | ➖ | ➖ | ➖ | ➖ | ➖ | ➖ | ➖ | ➖ | ➖ | ➖ | ➖ | ➖ | ➖ | ➖ |

## Notes

- **alignment**: Known limitation: pylightxl alignment write is a no-op because the library does not support formatting writes.
- **cell_values**: Known limitation: pylightxl cell-values write has date/boolean/error fidelity limits due to writer encoding behavior.
- **alignment**: Known limitation: python-calamine alignment read is limited because its API does not expose style/alignment metadata.
- **cell_values**: Known limitation: python-calamine can surface formula error cells as blank values in current API responses.
- **cell_values, formulas, ... (17 features)**: Not applicable: xlrd does not support .xlsx input
- **pivot_tables**: Unsupported on macOS without a Windows-generated pivot fixture (fixtures/excel/tier2/15_pivot_tables.xlsx).

## Statistics

| Library | Mode | Tests | Passed | Failed | Pass Rate | Green Features |
|---------|------|-------|--------|--------|-----------|----------------|
| openpyxl | R | 113 | 113 | 0 | 100% | 16/16 |
| openpyxl | W | 113 | 113 | 0 | 100% | 16/16 |
| openpyxl-readonly | R | 113 | 27 | 86 | 24% | 3/16 |
| pandas | R | 113 | 20 | 93 | 18% | 1/16 |
| pandas | W | 113 | 27 | 86 | 24% | 3/16 |
| polars | R | 113 | 18 | 95 | 16% | 0/16 |
| pyexcel | R | 113 | 23 | 90 | 20% | 2/16 |
| pyexcel | W | 113 | 28 | 85 | 25% | 3/16 |
| pylightxl | R | 113 | 22 | 91 | 19% | 2/16 |
| pylightxl | W | 113 | 23 | 90 | 20% | 2/16 |
| python-calamine | R | 113 | 20 | 93 | 18% | 1/16 |
| tablib | R | 113 | 23 | 90 | 20% | 2/16 |
| tablib | W | 113 | 28 | 85 | 25% | 3/16 |
| xlsxwriter | W | 113 | 113 | 0 | 100% | 16/16 |
| xlsxwriter-constmem | W | 113 | 106 | 7 | 94% | 13/16 |
| xlwt | W | 113 | 72 | 41 | 64% | 4/16 |

## Libraries Tested

- **openpyxl** v3.1.5 (python) - read, write
- **openpyxl-readonly** v3.1.5 (python) - read
- **pandas** v3.0.0 (python) - read, write
- **polars** v1.38.1 (python) - read
- **pyexcel** v0.7.4 (python) - read, write
- **pylightxl** v1.61 (python) - read, write
- **python-calamine** v0.6.1 (python) - read
- **tablib** v3.9.0 (python) - read, write
- **xlrd** v2.0.2 (python) - read
- **xlsxwriter** v3.2.9 (python) - write
- **xlsxwriter-constmem** v3.2.9 (python) - write
- **xlwt** v1.3.0 (python) - write

## Diagnostics Summary

No diagnostics recorded.

## Detailed Results

<a id="alignment-details"></a>
### alignment

**openpyxl** — Read: 🟢 3 | Write: 🟢 3

**openpyxl-readonly** — Read: 🟠 1

| Test | Importance | Read |
|------|-----------|------|
| Align - left | basic | ❌ |
| Align - center | basic | ❌ |
| Align - right | basic | ❌ |
| Align - top | basic | ❌ |
| Align - center | basic | ❌ |
| Align - bottom | basic | ✅ |
| Align - wrap text | basic | ❌ |
| Align - rotation 45 | basic | ❌ |
| Align - indent 2 | basic | ❌ |

**pandas** — Read: 🟠 1 | Write: 🟠 1

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Align - left | basic | ❌ | ❌ |
| Align - center | basic | ❌ | ❌ |
| Align - right | basic | ❌ | ❌ |
| Align - top | basic | ❌ | ❌ |
| Align - center | basic | ❌ | ❌ |
| Align - bottom | basic | ✅ | ✅ |
| Align - wrap text | basic | ❌ | ❌ |
| Align - rotation 45 | basic | ❌ | ❌ |
| Align - indent 2 | basic | ❌ | ❌ |

**polars** — Read: 🟠 1

| Test | Importance | Read |
|------|-----------|------|
| Align - left | basic | ❌ |
| Align - center | basic | ❌ |
| Align - right | basic | ❌ |
| Align - top | basic | ❌ |
| Align - center | basic | ❌ |
| Align - bottom | basic | ✅ |
| Align - wrap text | basic | ❌ |
| Align - rotation 45 | basic | ❌ |
| Align - indent 2 | basic | ❌ |

**pyexcel** — Read: 🟠 1 | Write: 🟠 1

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Align - left | basic | ❌ | ❌ |
| Align - center | basic | ❌ | ❌ |
| Align - right | basic | ❌ | ❌ |
| Align - top | basic | ❌ | ❌ |
| Align - center | basic | ❌ | ❌ |
| Align - bottom | basic | ✅ | ✅ |
| Align - wrap text | basic | ❌ | ❌ |
| Align - rotation 45 | basic | ❌ | ❌ |
| Align - indent 2 | basic | ❌ | ❌ |

**pylightxl** — Read: 🔴 0 | Write: 🟠 1
- Notes: Known limitation: pylightxl alignment write is a no-op because the library does not support formatting writes.

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Align - left | basic | ❌ | ❌ |
| Align - center | basic | ❌ | ❌ |
| Align - right | basic | ❌ | ❌ |
| Align - top | basic | ❌ | ❌ |
| Align - center | basic | ❌ | ❌ |
| Align - bottom | basic | ❌ | ✅ |
| Align - wrap text | basic | ❌ | ❌ |
| Align - rotation 45 | basic | ❌ | ❌ |
| Align - indent 2 | basic | ❌ | ❌ |

**python-calamine** — Read: 🟠 1
- Notes: Known limitation: python-calamine alignment read is limited because its API does not expose style/alignment metadata.

| Test | Importance | Read |
|------|-----------|------|
| Align - left | basic | ❌ |
| Align - center | basic | ❌ |
| Align - right | basic | ❌ |
| Align - top | basic | ❌ |
| Align - center | basic | ❌ |
| Align - bottom | basic | ✅ |
| Align - wrap text | basic | ❌ |
| Align - rotation 45 | basic | ❌ |
| Align - indent 2 | basic | ❌ |

**tablib** — Read: 🟠 1 | Write: 🟠 1

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Align - left | basic | ❌ | ❌ |
| Align - center | basic | ❌ | ❌ |
| Align - right | basic | ❌ | ❌ |
| Align - top | basic | ❌ | ❌ |
| Align - center | basic | ❌ | ❌ |
| Align - bottom | basic | ✅ | ✅ |
| Align - wrap text | basic | ❌ | ❌ |
| Align - rotation 45 | basic | ❌ | ❌ |
| Align - indent 2 | basic | ❌ | ❌ |

**xlrd**
- Notes: Not applicable: xlrd does not support .xlsx input

**xlsxwriter** — Write: 🟢 3

**xlsxwriter-constmem** — Write: 🟢 3

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

**tablib** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Background - red | basic | ❌ | ❌ |
| Background - blue | basic | ❌ | ❌ |
| Background - green | basic | ❌ | ❌ |
| Background - custom (#8B4513) | basic | ❌ | ❌ |

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

**tablib** — Read: 🟢 3 | Write: 🟢 3

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

**tablib** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Comment: legacy note | basic | ❌ | ❌ |
| Comment: threaded | edge | ❌ | ❌ |
| Comment: second author | edge | ❌ | ❌ |

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

**tablib** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| CF: cell > 5 (yellow fill) | basic | ❌ | ❌ |
| CF: formula rule with cross-sheet ref | edge | ❌ | ❌ |
| CF: text contains | edge | ❌ | ❌ |
| CF: data bar | edge | ❌ | ❌ |
| CF: 3-color scale | edge | ❌ | ❌ |
| CF: stop-if-true priority | edge | ❌ | ❌ |

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

**tablib** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| DV: list from CSV | basic | ❌ | ❌ |
| DV: list from range | edge | ❌ | ❌ |
| DV: cross-sheet named range | edge | ❌ | ❌ |
| DV: custom formula | edge | ❌ | ❌ |
| DV: whole number with error | basic | ❌ | ❌ |

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

**tablib** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Row height - 30 | basic | ❌ | ❌ |
| Row height - 45 | basic | ❌ | ❌ |
| Column width - D = 20 | basic | ❌ | ❌ |
| Column width - E = 8 | basic | ❌ | ❌ |

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

**tablib** — Read: 🔴 0 | Write: 🟢 3

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Formula - SUM | basic | ❌ | ✅ |
| Formula - cell reference | basic | ❌ | ✅ |
| Formula - concat | basic | ❌ | ✅ |
| Formula - cross sheet | basic | ❌ | ✅ |

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

**tablib** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Freeze panes at B2 | basic | ❌ | ❌ |
| Freeze panes at D5 | edge | ❌ | ❌ |
| Split panes row=2 col=1 | edge | ❌ | ❌ |

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

**tablib** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Hyperlink: external URL | basic | ❌ | ❌ |
| Hyperlink: internal sheet | edge | ❌ | ❌ |
| Hyperlink: mailto | basic | ❌ | ❌ |
| Hyperlink: long encoded URL | edge | ❌ | ❌ |

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

**tablib** — Read: 🔴 0 | Write: 🔴 0

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

**tablib** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Merge horizontal B2:D2 | basic | ❌ | ❌ |
| Merge vertical B3:B5 | basic | ❌ | ❌ |
| Merge with non-top-left value | edge | ❌ | ❌ |
| Merge with top-left fill | edge | ❌ | ❌ |

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

**tablib** — Read: 🟢 3 | Write: 🟢 3

**xlrd**
- Notes: Not applicable: xlrd does not support .xlsx input

**xlsxwriter** — Write: 🟢 3

**xlsxwriter-constmem** — Write: 🟢 3

**xlwt** — Write: 🟢 3

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

**pyexcel** — Read: 🔴 0 | Write: 🟠 1

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Format - currency | basic | ❌ | ❌ |
| Format - percent | basic | ❌ | ❌ |
| Format - date | basic | ❌ | ✅ |
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

**tablib** — Read: 🔴 0 | Write: 🟠 1

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Format - currency | basic | ❌ | ❌ |
| Format - percent | basic | ❌ | ❌ |
| Format - date | basic | ❌ | ✅ |
| Format - scientific | basic | ❌ | ❌ |
| Format - custom text | basic | ❌ | ❌ |

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

**tablib**
- Notes: Unsupported on macOS without a Windows-generated pivot fixture (fixtures/excel/tier2/15_pivot_tables.xlsx).

**xlrd**
- Notes: Not applicable: xlrd does not support .xlsx input

**xlsxwriter**
- Notes: Unsupported on macOS without a Windows-generated pivot fixture (fixtures/excel/tier2/15_pivot_tables.xlsx).

**xlsxwriter-constmem**
- Notes: Unsupported on macOS without a Windows-generated pivot fixture (fixtures/excel/tier2/15_pivot_tables.xlsx).

**xlwt**
- Notes: Unsupported on macOS without a Windows-generated pivot fixture (fixtures/excel/tier2/15_pivot_tables.xlsx).

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