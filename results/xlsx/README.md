# ExcelBench Results

*Generated: 2026-02-15 05:58 UTC*
*Profile: xlsx*
*Excel Version: 16.105.3*
*Platform: Darwin-arm64*

## Overview

> Condensed view — shows the **best score** across read/write for each library. See [Full Results Matrix](#full-results-matrix) for the complete R/W breakdown.

**Tier 0 — Basic Values**

| Feature | calamine-styled | rust_xlsxwriter | wolfxl |
|---------|:-:|:-:|:-:|
| Cell Values | 🟢 | 🟢 | 🟢 |
| Formulas | 🟢 | 🟢 | 🟢 |
| Sheets | 🟢 | 🟢 | 🟢 |

**Tier 1 — Formatting**

| Feature | calamine-styled | rust_xlsxwriter | wolfxl |
|---------|:-:|:-:|:-:|
| Alignment | 🟢 | 🟢 | 🟢 |
| Bg Colors | 🟢 | 🟢 | 🟢 |
| Borders | 🟠 | 🟢 | 🟢 |
| Dimensions | 🟢 | 🟢 | 🟢 |
| Num Fmt | 🟢 | 🟢 | 🟢 |
| Text Fmt | 🟢 | 🟢 | 🟢 |

**Tier 2 — Advanced**

| Feature | calamine-styled | rust_xlsxwriter | wolfxl |
|---------|:-:|:-:|:-:|
| Comments | 🟢 | 🟢 | 🟢 |
| Cond Fmt | 🟢 | 🟢 | 🟢 |
| Validation | 🟢 | 🟢 | 🟢 |
| Freeze | 🟢 | 🟢 | 🟢 |
| Hyperlinks | 🟢 | 🟢 | 🟢 |
| Images | 🔴 | 🔴 | 🔴 |
| Merged | 🟢 | 🟢 | 🟢 |

**Tier 3 — Workbook Metadata**

| Feature | calamine-styled | rust_xlsxwriter | wolfxl |
|---------|:-:|:-:|:-:|
| Named Ranges | 🟢 | 🟢 | 🟢 |
| Tables | 🟢 | 🟢 | 🟢 |

## Library Tiers

> Libraries ranked by their best capability (max of read/write green features).

| Tier | Library | Caps | Green Features | Summary |
|:----:|---------|:----:|:--------------:|---------|
| **A** | rust_xlsxwriter | W | 17/18 | 17/18 features with full fidelity |
| **A** | wolfxl | R+W | 17/18 | 17/18 features with full fidelity |
| **A** | calamine-styled | R | 16/18 | 16/18 features with full fidelity |

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

| Feature | calamine-styled (R) | rust_xlsxwriter (W) | wolfxl (R) | wolfxl (W) |
|---------|------------|------------|------------|------------|
| [cell_values](#cell_values-details) | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 |
| [formulas](#formulas-details) | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 |
| [multiple_sheets](#multiple_sheets-details) | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 |

**Tier 1 — Formatting**

| Feature | calamine-styled (R) | rust_xlsxwriter (W) | wolfxl (R) | wolfxl (W) |
|---------|------------|------------|------------|------------|
| [alignment](#alignment-details) | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 |
| [background_colors](#background_colors-details) | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 |
| [borders](#borders-details) | 🟠 1 | 🟢 3 | 🟠 1 | 🟢 3 |
| [dimensions](#dimensions-details) | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 |
| [number_formats](#number_formats-details) | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 |
| [text_formatting](#text_formatting-details) | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 |

**Tier 2 — Advanced**

| Feature | calamine-styled (R) | rust_xlsxwriter (W) | wolfxl (R) | wolfxl (W) |
|---------|------------|------------|------------|------------|
| [comments](#comments-details) | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 |
| [conditional_formatting](#conditional_formatting-details) | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 |
| [data_validation](#data_validation-details) | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 |
| [freeze_panes](#freeze_panes-details) | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 |
| [hyperlinks](#hyperlinks-details) | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 |
| [images](#images-details) | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 |
| [merged_cells](#merged_cells-details) | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 |
| [pivot_tables](#pivot_tables-details) | ➖ | ➖ | ➖ | ➖ |

**Tier 3 — Workbook Metadata**

| Feature | calamine-styled (R) | rust_xlsxwriter (W) | wolfxl (R) | wolfxl (W) |
|---------|------------|------------|------------|------------|
| [named_ranges](#named_ranges-details) | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 |
| [tables](#tables-details) | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 |

## Notes

- **pivot_tables**: Unsupported on macOS without a Windows-generated pivot fixture (fixtures/excel/tier2/15_pivot_tables.xlsx).

## Statistics

| Library | Mode | Tests | Passed | Failed | Pass Rate | Green Features |
|---------|------|-------|--------|--------|-----------|----------------|
| calamine-styled | R | 125 | 120 | 5 | 96% | 16/18 |
| rust_xlsxwriter | W | 125 | 123 | 2 | 98% | 17/18 |
| wolfxl | R | 125 | 120 | 5 | 96% | 16/18 |
| wolfxl | W | 125 | 123 | 2 | 98% | 17/18 |

## Libraries Tested

- **calamine-styled** v0.33.0 (rust) - read
- **rust_xlsxwriter** v0.79.4 (rust) - write
- **wolfxl** vcal=0.33.0+rxw=0.79.4 (rust) - read, write

## Diagnostics Summary

| Group | Value | Count |
|-------|-------|-------|
| category | data_mismatch | 14 |
| severity | error | 14 |

### Diagnostic Details

| Feature | Library | Test Case | Operation | Category | Severity | Message |
|---------|---------|-----------|-----------|----------|----------|---------|
| borders | wolfxl | diagonal_up | read | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_up': 'thin'}, actual={} |
| borders | wolfxl | diagonal_down | read | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_down': 'thin'}, actual={} |
| borders | wolfxl | diagonal_both | read | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_up': 'thin', 'border_diagonal_down': 'thin'}, actual={} |
| borders | calamine-styled | diagonal_up | read | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_up': 'thin'}, actual={} |
| borders | calamine-styled | diagonal_down | read | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_down': 'thin'}, actual={} |
| borders | calamine-styled | diagonal_both | read | data_mismatch | error | Expected values did not match actual values: expected={'border_diagonal_up': 'thin', 'border_diagonal_down': 'thin'}, actual={} |
| images | wolfxl | image_one_cell | read | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'B2', 'path': 'fixtures/images/sample.png', 'anchor': 'oneCell'}}, actual={} |
| images | wolfxl | image_one_cell | write | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'B2', 'path': 'fixtures/images/sample.png', 'anchor': 'oneCell'}}, actual={} |
| images | wolfxl | image_two_cell_offset | read | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'D6', 'path': 'fixtures/images/sample.jpg', 'anchor': 'oneCell'}}, actual={} |
| images | wolfxl | image_two_cell_offset | write | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'D6', 'path': 'fixtures/images/sample.jpg', 'anchor': 'oneCell'}}, actual={} |
| images | calamine-styled | image_one_cell | read | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'B2', 'path': 'fixtures/images/sample.png', 'anchor': 'oneCell'}}, actual={} |
| images | calamine-styled | image_two_cell_offset | read | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'D6', 'path': 'fixtures/images/sample.jpg', 'anchor': 'oneCell'}}, actual={} |
| images | rust_xlsxwriter | image_one_cell | write | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'B2', 'path': 'fixtures/images/sample.png', 'anchor': 'oneCell'}}, actual={} |
| images | rust_xlsxwriter | image_two_cell_offset | write | data_mismatch | error | Expected values did not match actual values: expected={'image': {'cell': 'D6', 'path': 'fixtures/images/sample.jpg', 'anchor': 'oneCell'}}, actual={} |

## Detailed Results

<a id="alignment-details"></a>
### alignment

**calamine-styled** — Read: 🟢 3

**rust_xlsxwriter** — Write: 🟢 3

**wolfxl** — Read: 🟢 3 | Write: 🟢 3

<a id="background_colors-details"></a>
### background_colors

**calamine-styled** — Read: 🟢 3

**rust_xlsxwriter** — Write: 🟢 3

**wolfxl** — Read: 🟢 3 | Write: 🟢 3

<a id="borders-details"></a>
### borders

**calamine-styled** — Read: 🟠 1

| Test | Importance | Read |
|------|-----------|------|
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
| Border - diagonal both | basic | ❌ |
| Border - red color | basic | ✅ |
| Border - blue color | basic | ✅ |
| Border - custom color (#8B4513) | basic | ✅ |
| Border - mixed styles per edge | basic | ✅ |
| Border - mixed colors per edge | basic | ✅ |

**rust_xlsxwriter** — Write: 🟢 3

**wolfxl** — Read: 🟠 1 | Write: 🟢 3

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Border - thin all edges | basic | ✅ | ✅ |
| Border - medium all edges | basic | ✅ | ✅ |
| Border - thick all edges | basic | ✅ | ✅ |
| Border - double line | basic | ✅ | ✅ |
| Border - dashed | basic | ✅ | ✅ |
| Border - dotted | basic | ✅ | ✅ |
| Border - dash-dot | basic | ✅ | ✅ |
| Border - dash-dot-dot | basic | ✅ | ✅ |
| Border - top only | basic | ✅ | ✅ |
| Border - bottom only | basic | ✅ | ✅ |
| Border - left only | basic | ✅ | ✅ |
| Border - right only | basic | ✅ | ✅ |
| Border - diagonal up | basic | ❌ | ✅ |
| Border - diagonal down | basic | ❌ | ✅ |
| Border - diagonal both | basic | ❌ | ✅ |
| Border - red color | basic | ✅ | ✅ |
| Border - blue color | basic | ✅ | ✅ |
| Border - custom color (#8B4513) | basic | ✅ | ✅ |
| Border - mixed styles per edge | basic | ✅ | ✅ |
| Border - mixed colors per edge | basic | ✅ | ✅ |

<a id="cell_values-details"></a>
### cell_values

**calamine-styled** — Read: 🟢 3

**rust_xlsxwriter** — Write: 🟢 3

**wolfxl** — Read: 🟢 3 | Write: 🟢 3

<a id="comments-details"></a>
### comments

**calamine-styled** — Read: 🟢 3

**rust_xlsxwriter** — Write: 🟢 3

**wolfxl** — Read: 🟢 3 | Write: 🟢 3

<a id="conditional_formatting-details"></a>
### conditional_formatting

**calamine-styled** — Read: 🟢 3

**rust_xlsxwriter** — Write: 🟢 3

**wolfxl** — Read: 🟢 3 | Write: 🟢 3

<a id="data_validation-details"></a>
### data_validation

**calamine-styled** — Read: 🟢 3

**rust_xlsxwriter** — Write: 🟢 3

**wolfxl** — Read: 🟢 3 | Write: 🟢 3

<a id="dimensions-details"></a>
### dimensions

**calamine-styled** — Read: 🟢 3

**rust_xlsxwriter** — Write: 🟢 3

**wolfxl** — Read: 🟢 3 | Write: 🟢 3

<a id="formulas-details"></a>
### formulas

**calamine-styled** — Read: 🟢 3

**rust_xlsxwriter** — Write: 🟢 3

**wolfxl** — Read: 🟢 3 | Write: 🟢 3

<a id="freeze_panes-details"></a>
### freeze_panes

**calamine-styled** — Read: 🟢 3

**rust_xlsxwriter** — Write: 🟢 3

**wolfxl** — Read: 🟢 3 | Write: 🟢 3

<a id="hyperlinks-details"></a>
### hyperlinks

**calamine-styled** — Read: 🟢 3

**rust_xlsxwriter** — Write: 🟢 3

**wolfxl** — Read: 🟢 3 | Write: 🟢 3

<a id="images-details"></a>
### images

**calamine-styled** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Image: one-cell anchor | basic | ❌ |
| Image: two-cell anchor with offset | edge | ❌ |

**rust_xlsxwriter** — Write: 🔴 0

| Test | Importance | Write |
|------|-----------|-------|
| Image: one-cell anchor | basic | ❌ |
| Image: two-cell anchor with offset | edge | ❌ |

**wolfxl** — Read: 🔴 0 | Write: 🔴 0

| Test | Importance | Read | Write |
|------|-----------|------|-------|
| Image: one-cell anchor | basic | ❌ | ❌ |
| Image: two-cell anchor with offset | edge | ❌ | ❌ |

<a id="merged_cells-details"></a>
### merged_cells

**calamine-styled** — Read: 🟢 3

**rust_xlsxwriter** — Write: 🟢 3

**wolfxl** — Read: 🟢 3 | Write: 🟢 3

<a id="multiple_sheets-details"></a>
### multiple_sheets

**calamine-styled** — Read: 🟢 3

**rust_xlsxwriter** — Write: 🟢 3

**wolfxl** — Read: 🟢 3 | Write: 🟢 3

<a id="named_ranges-details"></a>
### named_ranges

**calamine-styled** — Read: 🟢 3

**rust_xlsxwriter** — Write: 🟢 3

**wolfxl** — Read: 🟢 3 | Write: 🟢 3

<a id="number_formats-details"></a>
### number_formats

**calamine-styled** — Read: 🟢 3

**rust_xlsxwriter** — Write: 🟢 3

**wolfxl** — Read: 🟢 3 | Write: 🟢 3

<a id="pivot_tables-details"></a>
### pivot_tables

**calamine-styled**
- Notes: Unsupported on macOS without a Windows-generated pivot fixture (fixtures/excel/tier2/15_pivot_tables.xlsx).

**rust_xlsxwriter**
- Notes: Unsupported on macOS without a Windows-generated pivot fixture (fixtures/excel/tier2/15_pivot_tables.xlsx).

**wolfxl**
- Notes: Unsupported on macOS without a Windows-generated pivot fixture (fixtures/excel/tier2/15_pivot_tables.xlsx).

<a id="tables-details"></a>
### tables

**calamine-styled** — Read: 🟢 3

**rust_xlsxwriter** — Write: 🟢 3

**wolfxl** — Read: 🟢 3 | Write: 🟢 3

<a id="text_formatting-details"></a>
### text_formatting

**calamine-styled** — Read: 🟢 3

**rust_xlsxwriter** — Write: 🟢 3

**wolfxl** — Read: 🟢 3 | Write: 🟢 3

---
*Benchmark version: 0.1.0*