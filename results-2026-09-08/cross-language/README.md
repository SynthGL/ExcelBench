# ExcelBench Results

*Generated: 2026-09-08 05:37 UTC*
*Profile: xlsx*
*Excel Version: 16.105.3*
*Platform: Darwin-arm64*

## Overview

> Condensed view — shows the **best score** across read/write for each library. See [Full Results Matrix](#full-results-matrix) for the complete R/W breakdown.

**Tier 0 — Basic Values**

| Feature | zavora-xlsx |
|---------|:-:|
| Cell Values | 🔴 |
| Formulas | 🟢 |
| Sheets | 🟢 |

**Tier 1 — Formatting**

| Feature | zavora-xlsx |
|---------|:-:|
| Alignment | 🟠 |
| Bg Colors | 🟢 |
| Borders | 🔴 |
| Dimensions | 🟠 |
| Num Fmt | 🟢 |
| Text Fmt | 🟢 |

**Tier 2 — Advanced**

| Feature | zavora-xlsx |
|---------|:-:|
| chart_anchor | 🔴 |
| Comments | 🟢 |
| Cond Fmt | 🔴 |
| Validation | 🔴 |
| Freeze | 🔴 |
| Hyperlinks | 🔴 |
| Images | 🔴 |
| Merged | 🟢 |
| page_setup | 🔴 |
| sheet_protection | 🔴 |

**Tier 3 — Workbook Metadata**

| Feature | zavora-xlsx |
|---------|:-:|
| Named Ranges | 🟢 |
| Tables | 🟠 |

## Library Tiers

> Libraries ranked by their best capability (max of read/write green features).

| Tier | Library | Caps | Modify | Green Features | Summary |
|:----:|---------|:----:|:------:|:--------------:|---------|
| **B** | zavora-xlsx | W | No | 8/21 | 8/21 features with full fidelity |

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

| Feature | zavora-xlsx (W) |
|---------|------------|
| [cell_values](#cell_values-details) | 🔴 0 |
| [formulas](#formulas-details) | 🟢 3 |
| [multiple_sheets](#multiple_sheets-details) | 🟢 3 |

**Tier 1 — Formatting**

| Feature | zavora-xlsx (W) |
|---------|------------|
| [alignment](#alignment-details) | 🟠 1 |
| [background_colors](#background_colors-details) | 🟢 3 |
| [borders](#borders-details) | 🔴 0 |
| [dimensions](#dimensions-details) | 🟠 1 |
| [number_formats](#number_formats-details) | 🟢 3 |
| [text_formatting](#text_formatting-details) | 🟢 3 |

**Tier 2 — Advanced**

| Feature | zavora-xlsx (W) |
|---------|------------|
| [chart_anchor](#chart_anchor-details) | 🔴 0 |
| [comments](#comments-details) | 🟢 3 |
| [conditional_formatting](#conditional_formatting-details) | 🔴 0 |
| [data_validation](#data_validation-details) | 🔴 0 |
| [freeze_panes](#freeze_panes-details) | 🔴 0 |
| [hyperlinks](#hyperlinks-details) | 🔴 0 |
| [images](#images-details) | 🔴 0 |
| [merged_cells](#merged_cells-details) | 🟢 3 |
| [page_setup](#page_setup-details) | 🔴 0 |
| [pivot_tables](#pivot_tables-details) | ➖ |
| [sheet_protection](#sheet_protection-details) | 🔴 0 |

**Tier 3 — Workbook Metadata**

| Feature | zavora-xlsx (W) |
|---------|------------|
| [named_ranges](#named_ranges-details) | 🟢 3 |
| [tables](#tables-details) | 🟠 1 |

## Notes

- **pivot_tables**: Unsupported on macOS without a Windows-generated pivot fixture (fixtures/excel/tier2/15_pivot_tables.xlsx).

## Statistics

| Library | Mode | Tests | Passed | Failed | Pass Rate | Green Features |
|---------|------|-------|--------|--------|-----------|----------------|
| zavora-xlsx | W | 133 | 63 | 70 | 47% | 8/21 |

## Libraries Tested

- **zavora-xlsx** v0.1.2 (rust) - write; modify: No

## Diagnostics Summary

| Group | Value | Count |
|-------|-------|-------|
| category | data_mismatch | 4 |
| category | unsupported_feature | 66 |
| severity | error | 4 |
| severity | warning | 66 |

### Diagnostic Details

| Feature | Library | Test Case | Operation | Category | Severity | Message |
|---------|---------|-----------|-----------|----------|----------|---------|
| cell_values | zavora-xlsx | string_simple | write | unsupported_feature | warning | RuntimeError: unsupported cell type "date" at cell_values!B12 |
| cell_values | zavora-xlsx | string_unicode | write | unsupported_feature | warning | RuntimeError: unsupported cell type "date" at cell_values!B12 |
| cell_values | zavora-xlsx | string_empty | write | unsupported_feature | warning | RuntimeError: unsupported cell type "date" at cell_values!B12 |
| cell_values | zavora-xlsx | string_long | write | unsupported_feature | warning | RuntimeError: unsupported cell type "date" at cell_values!B12 |
| cell_values | zavora-xlsx | string_newline | write | unsupported_feature | warning | RuntimeError: unsupported cell type "date" at cell_values!B12 |
| cell_values | zavora-xlsx | number_integer | write | unsupported_feature | warning | RuntimeError: unsupported cell type "date" at cell_values!B12 |
| cell_values | zavora-xlsx | number_float | write | unsupported_feature | warning | RuntimeError: unsupported cell type "date" at cell_values!B12 |
| cell_values | zavora-xlsx | number_negative | write | unsupported_feature | warning | RuntimeError: unsupported cell type "date" at cell_values!B12 |
| cell_values | zavora-xlsx | number_large | write | unsupported_feature | warning | RuntimeError: unsupported cell type "date" at cell_values!B12 |
| cell_values | zavora-xlsx | number_scientific | write | unsupported_feature | warning | RuntimeError: unsupported cell type "date" at cell_values!B12 |
| cell_values | zavora-xlsx | date_standard | write | unsupported_feature | warning | RuntimeError: unsupported cell type "date" at cell_values!B12 |
| cell_values | zavora-xlsx | datetime | write | unsupported_feature | warning | RuntimeError: unsupported cell type "date" at cell_values!B12 |
| cell_values | zavora-xlsx | boolean_true | write | unsupported_feature | warning | RuntimeError: unsupported cell type "date" at cell_values!B12 |
| cell_values | zavora-xlsx | boolean_false | write | unsupported_feature | warning | RuntimeError: unsupported cell type "date" at cell_values!B12 |
| cell_values | zavora-xlsx | error_div0 | write | unsupported_feature | warning | RuntimeError: unsupported cell type "date" at cell_values!B12 |
| cell_values | zavora-xlsx | error_na | write | unsupported_feature | warning | RuntimeError: unsupported cell type "date" at cell_values!B12 |
| cell_values | zavora-xlsx | error_value | write | unsupported_feature | warning | RuntimeError: unsupported cell type "date" at cell_values!B12 |
| cell_values | zavora-xlsx | blank | write | unsupported_feature | warning | RuntimeError: unsupported cell type "date" at cell_values!B12 |
| alignment | zavora-xlsx | v_top | write | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'top'}, actual={} |
| borders | zavora-xlsx | thin_all | write | unsupported_feature | warning | RuntimeError: unsupported border style "dashDot" |
| borders | zavora-xlsx | medium_all | write | unsupported_feature | warning | RuntimeError: unsupported border style "dashDot" |
| borders | zavora-xlsx | thick_all | write | unsupported_feature | warning | RuntimeError: unsupported border style "dashDot" |
| borders | zavora-xlsx | double | write | unsupported_feature | warning | RuntimeError: unsupported border style "dashDot" |
| borders | zavora-xlsx | dashed | write | unsupported_feature | warning | RuntimeError: unsupported border style "dashDot" |
| borders | zavora-xlsx | dotted | write | unsupported_feature | warning | RuntimeError: unsupported border style "dashDot" |
| borders | zavora-xlsx | dash_dot | write | unsupported_feature | warning | RuntimeError: unsupported border style "dashDot" |
| borders | zavora-xlsx | dash_dot_dot | write | unsupported_feature | warning | RuntimeError: unsupported border style "dashDot" |
| borders | zavora-xlsx | top_only | write | unsupported_feature | warning | RuntimeError: unsupported border style "dashDot" |
| borders | zavora-xlsx | bottom_only | write | unsupported_feature | warning | RuntimeError: unsupported border style "dashDot" |
| borders | zavora-xlsx | left_only | write | unsupported_feature | warning | RuntimeError: unsupported border style "dashDot" |
| borders | zavora-xlsx | right_only | write | unsupported_feature | warning | RuntimeError: unsupported border style "dashDot" |
| borders | zavora-xlsx | diagonal_up | write | unsupported_feature | warning | RuntimeError: unsupported border style "dashDot" |
| borders | zavora-xlsx | diagonal_down | write | unsupported_feature | warning | RuntimeError: unsupported border style "dashDot" |
| borders | zavora-xlsx | diagonal_both | write | unsupported_feature | warning | RuntimeError: unsupported border style "dashDot" |
| borders | zavora-xlsx | color_red | write | unsupported_feature | warning | RuntimeError: unsupported border style "dashDot" |
| borders | zavora-xlsx | color_blue | write | unsupported_feature | warning | RuntimeError: unsupported border style "dashDot" |
| borders | zavora-xlsx | color_custom | write | unsupported_feature | warning | RuntimeError: unsupported border style "dashDot" |
| borders | zavora-xlsx | mixed_styles | write | unsupported_feature | warning | RuntimeError: unsupported border style "dashDot" |
| borders | zavora-xlsx | mixed_colors | write | unsupported_feature | warning | RuntimeError: unsupported border style "dashDot" |
| dimensions | zavora-xlsx | row_height_30 | write | data_mismatch | error | Expected values did not match actual values: expected={'row_height': 30}, actual={'row_height': None} |
| dimensions | zavora-xlsx | row_height_45 | write | data_mismatch | error | Expected values did not match actual values: expected={'row_height': 45}, actual={'row_height': None} |
| conditional_formatting | zavora-xlsx | cf_cell_gt | write | unsupported_feature | warning | RuntimeError: unsupported payload part: conditional_formats |
| conditional_formatting | zavora-xlsx | cf_formula_cross_sheet | write | unsupported_feature | warning | RuntimeError: unsupported payload part: conditional_formats |
| conditional_formatting | zavora-xlsx | cf_text_contains | write | unsupported_feature | warning | RuntimeError: unsupported payload part: conditional_formats |
| conditional_formatting | zavora-xlsx | cf_data_bar | write | unsupported_feature | warning | RuntimeError: unsupported payload part: conditional_formats |
| conditional_formatting | zavora-xlsx | cf_color_scale | write | unsupported_feature | warning | RuntimeError: unsupported payload part: conditional_formats |
| conditional_formatting | zavora-xlsx | cf_stop_if_true | write | unsupported_feature | warning | RuntimeError: unsupported payload part: conditional_formats |
| data_validation | zavora-xlsx | dv_list_csv | write | unsupported_feature | warning | RuntimeError: unsupported payload part: validations |
| data_validation | zavora-xlsx | dv_list_range | write | unsupported_feature | warning | RuntimeError: unsupported payload part: validations |
| data_validation | zavora-xlsx | dv_cross_sheet | write | unsupported_feature | warning | RuntimeError: unsupported payload part: validations |
| data_validation | zavora-xlsx | dv_custom_formula | write | unsupported_feature | warning | RuntimeError: unsupported payload part: validations |
| data_validation | zavora-xlsx | dv_whole_between | write | unsupported_feature | warning | RuntimeError: unsupported payload part: validations |
| hyperlinks | zavora-xlsx | link_external | write | unsupported_feature | warning | RuntimeError: unsupported hyperlink part: tooltip |
| hyperlinks | zavora-xlsx | link_internal | write | unsupported_feature | warning | RuntimeError: unsupported hyperlink part: tooltip |
| hyperlinks | zavora-xlsx | link_mailto | write | unsupported_feature | warning | RuntimeError: unsupported hyperlink part: tooltip |
| hyperlinks | zavora-xlsx | link_long | write | unsupported_feature | warning | RuntimeError: unsupported hyperlink part: tooltip |
| images | zavora-xlsx | image_one_cell | write | unsupported_feature | warning | RuntimeError: unsupported payload part: pictures |
| images | zavora-xlsx | image_two_cell_offset | write | unsupported_feature | warning | RuntimeError: unsupported payload part: pictures |
| freeze_panes | zavora-xlsx | freeze_b2 | write | unsupported_feature | warning | RuntimeError: unsupported pane part: top_left_cell |
| freeze_panes | zavora-xlsx | freeze_d5 | write | unsupported_feature | warning | RuntimeError: unsupported pane part: top_left_cell |
| freeze_panes | zavora-xlsx | split_2x1 | write | unsupported_feature | warning | RuntimeError: unsupported pane part: top_left_cell |
| tables | zavora-xlsx | tbl_with_totals | write | data_mismatch | error | Expected values did not match actual values: expected={'table': {'name': 'Summary', 'ref': 'E7:G11', 'header_row': True, 'totals_row': True, 'style': 'TableStyleLight1', 'columns': ['Item', 'Count', 'Total'], 'totals_row_count': 1}}, actual={'table': {'name': 'Summary', 'ref': 'E7:G11', 'header_row': True, 'totals_row': False, 'style': 'TableStyleLight1', 'columns': ['Item', 'Count', 'Total'], 'totals_row_count': 0}} |
| sheet_protection | zavora-xlsx | prot_basic | write | unsupported_feature | warning | NotImplementedError: zavora-xlsx does not implement sheet protection writes |
| sheet_protection | zavora-xlsx | prot_password | write | unsupported_feature | warning | NotImplementedError: zavora-xlsx does not implement sheet protection writes |
| sheet_protection | zavora-xlsx | prot_granular | write | unsupported_feature | warning | NotImplementedError: zavora-xlsx does not implement sheet protection writes |
| page_setup | zavora-xlsx | page_setup_landscape | write | unsupported_feature | warning | NotImplementedError: zavora-xlsx does not implement page setup writes |
| page_setup | zavora-xlsx | page_setup_scale | write | unsupported_feature | warning | NotImplementedError: zavora-xlsx does not implement page setup writes |
| page_setup | zavora-xlsx | page_setup_titles | write | unsupported_feature | warning | NotImplementedError: zavora-xlsx does not implement page setup writes |
| chart_anchor | zavora-xlsx | chart_bar | write | unsupported_feature | warning | NotImplementedError: zavora-xlsx does not implement chart anchor writes |
| chart_anchor | zavora-xlsx | chart_line | write | unsupported_feature | warning | NotImplementedError: zavora-xlsx does not implement chart anchor writes |

## Detailed Results

<a id="alignment-details"></a>
### alignment

**zavora-xlsx** — Write: 🟠 1

| Test | Importance | Write |
|------|-----------|-------|
| Align - left | basic | ✅ |
| Align - center | basic | ✅ |
| Align - right | basic | ✅ |
| Align - top | basic | ❌ |
| Align - center | basic | ✅ |
| Align - bottom | basic | ✅ |
| Align - wrap text | basic | ✅ |
| Align - rotation 45 | basic | ✅ |
| Align - indent 2 | basic | ✅ |

<a id="background_colors-details"></a>
### background_colors

**zavora-xlsx** — Write: 🟢 3

<a id="borders-details"></a>
### borders

**zavora-xlsx** — Write: 🔴 0

| Test | Importance | Write |
|------|-----------|-------|
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

<a id="cell_values-details"></a>
### cell_values

**zavora-xlsx** — Write: 🔴 0

| Test | Importance | Write |
|------|-----------|-------|
| String - simple | basic | ❌ |
| String - unicode | basic | ❌ |
| String - empty | basic | ❌ |
| String - long (1000 chars) | basic | ❌ |
| String - with newlines | basic | ❌ |
| Number - integer | basic | ❌ |
| Number - float | basic | ❌ |
| Number - negative | basic | ❌ |
| Number - large | basic | ❌ |
| Number - scientific notation | basic | ❌ |
| Date - standard | basic | ❌ |
| DateTime - with time | basic | ❌ |
| Boolean - TRUE | basic | ❌ |
| Boolean - FALSE | basic | ❌ |
| Error - #DIV/0! | basic | ❌ |
| Error - #N/A | basic | ❌ |
| Error - #VALUE! | basic | ❌ |
| Blank cell | basic | ❌ |

<a id="chart_anchor-details"></a>
### chart_anchor

**zavora-xlsx** — Write: 🔴 0

| Test | Importance | Write |
|------|-----------|-------|
| Chart: bar with twoCell anchor | basic | ❌ |
| Chart: line with twoCell anchor | edge | ❌ |

<a id="comments-details"></a>
### comments

**zavora-xlsx** — Write: 🟢 3

<a id="conditional_formatting-details"></a>
### conditional_formatting

**zavora-xlsx** — Write: 🔴 0

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

**zavora-xlsx** — Write: 🔴 0

| Test | Importance | Write |
|------|-----------|-------|
| DV: list from CSV | basic | ❌ |
| DV: list from range | edge | ❌ |
| DV: cross-sheet named range | edge | ❌ |
| DV: custom formula | edge | ❌ |
| DV: whole number with error | basic | ❌ |

<a id="dimensions-details"></a>
### dimensions

**zavora-xlsx** — Write: 🟠 1

| Test | Importance | Write |
|------|-----------|-------|
| Row height - 30 | basic | ❌ |
| Row height - 45 | basic | ❌ |
| Column width - D = 20 | basic | ✅ |
| Column width - E = 8 | basic | ✅ |

<a id="formulas-details"></a>
### formulas

**zavora-xlsx** — Write: 🟢 3

<a id="freeze_panes-details"></a>
### freeze_panes

**zavora-xlsx** — Write: 🔴 0

| Test | Importance | Write |
|------|-----------|-------|
| Freeze panes at B2 | basic | ❌ |
| Freeze panes at D5 | edge | ❌ |
| Split panes row=2 col=1 | edge | ❌ |

<a id="hyperlinks-details"></a>
### hyperlinks

**zavora-xlsx** — Write: 🔴 0

| Test | Importance | Write |
|------|-----------|-------|
| Hyperlink: external URL | basic | ❌ |
| Hyperlink: internal sheet | edge | ❌ |
| Hyperlink: mailto | basic | ❌ |
| Hyperlink: long encoded URL | edge | ❌ |

<a id="images-details"></a>
### images

**zavora-xlsx** — Write: 🔴 0

| Test | Importance | Write |
|------|-----------|-------|
| Image: one-cell anchor | basic | ❌ |
| Image: two-cell anchor with offset | edge | ❌ |

<a id="merged_cells-details"></a>
### merged_cells

**zavora-xlsx** — Write: 🟢 3

<a id="multiple_sheets-details"></a>
### multiple_sheets

**zavora-xlsx** — Write: 🟢 3

<a id="named_ranges-details"></a>
### named_ranges

**zavora-xlsx** — Write: 🟢 3

<a id="number_formats-details"></a>
### number_formats

**zavora-xlsx** — Write: 🟢 3

<a id="page_setup-details"></a>
### page_setup

**zavora-xlsx** — Write: 🔴 0

| Test | Importance | Write |
|------|-----------|-------|
| Page setup: landscape + fit-to-width | basic | ❌ |
| Page setup: portrait at 75% scale | basic | ❌ |
| Page setup: print titles + header/footer | edge | ❌ |

<a id="pivot_tables-details"></a>
### pivot_tables

**zavora-xlsx**
- Notes: Unsupported on macOS without a Windows-generated pivot fixture (fixtures/excel/tier2/15_pivot_tables.xlsx).

<a id="sheet_protection-details"></a>
### sheet_protection

**zavora-xlsx** — Write: 🔴 0

| Test | Importance | Write |
|------|-----------|-------|
| Protection: basic (OOXML defaults) | basic | ❌ |
| Protection: with password hash | basic | ❌ |
| Protection: granular attribute flags | edge | ❌ |

<a id="tables-details"></a>
### tables

**zavora-xlsx** — Write: 🟠 1

| Test | Importance | Write |
|------|-----------|-------|
| Table: basic 3-col | basic | ✅ |
| Table: with totals row | basic | ❌ |
| Table: no style | basic | ✅ |
| Table: single column | edge | ✅ |
| Table: header only (no data rows) | edge | ✅ |
| Table: with autoFilter | edge | ✅ |

<a id="text_formatting-details"></a>
### text_formatting

**zavora-xlsx** — Write: 🟢 3

---
*Benchmark version: 0.1.0*