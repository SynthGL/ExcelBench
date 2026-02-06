# ExcelBench Results

*Generated: 2026-02-06 06:09 UTC*
*Excel Version: 16.105.3*
*Platform: Darwin-arm64*

## Score Legend

| Score | Meaning |
|-------|---------|
| 🟢 3 | Complete - full fidelity |
| 🟡 2 | Functional - works for common cases |
| 🟠 1 | Minimal - basic recognition only |
| 🔴 0 | Unsupported - errors or data loss |
| ➖ | Not applicable (library doesn't support this operation) |

## Summary

| Feature | openpyxl (R) | openpyxl (W) | pylightxl (R) | pylightxl (W) | python-calamine (R) | xlsxwriter (W) |
|---------|------------|------------|------------|------------|------------|------------|
| alignment | 🟠 1 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🟠 1 |
| background_colors | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🟢 3 |
| borders | 🟢 3 | 🟠 1 | 🔴 0 | 🔴 0 | 🔴 0 | 🟠 1 |
| cell_values | 🟢 3 | 🟢 3 | 🟠 1 | 🔴 0 | 🟠 1 | 🟢 3 |
| comments | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🟢 3 |
| conditional_formatting | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 |
| data_validation | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🟠 1 |
| dimensions | 🟠 1 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🟠 1 |
| formulas | 🟠 1 | 🟢 3 | 🔴 0 | 🟢 3 | 🔴 0 | 🟢 3 |
| freeze_panes | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🔴 0 | 🟡 2 |
| hyperlinks | 🟠 1 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🟡 2 |
| images | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🟢 3 |
| merged_cells | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🟢 3 |
| multiple_sheets | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 |
| number_formats | 🟠 1 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🟢 3 |
| pivot_tables | ➖ | ➖ | ➖ | ➖ | ➖ | ➖ |
| text_formatting | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 | 🟢 3 |

Notes:
- pivot_tables: Unsupported on macOS without a Windows-generated pivot fixture (fixtures/excel/tier2/15_pivot_tables.xlsx).
- pivot_tables: Unsupported on macOS without a Windows-generated pivot fixture (fixtures/excel/tier2/15_pivot_tables.xlsx).
- pivot_tables: Unsupported on macOS without a Windows-generated pivot fixture (fixtures/excel/tier2/15_pivot_tables.xlsx).
- pivot_tables: Unsupported on macOS without a Windows-generated pivot fixture (fixtures/excel/tier2/15_pivot_tables.xlsx).

## Libraries Tested

- **openpyxl** v3.1.5 (python) - read, write
- **pylightxl** v1.61 (python) - read, write
- **python-calamine** v0.6.1 (python) - read
- **xlsxwriter** v3.2.9 (python) - write

## Detailed Results

### alignment

**openpyxl**
- Read: 🟠 1 (1/3)
- Write: 🟢 3 (3/3)
- Failed tests (8):
  - h_left (read)
  - h_center (read)
  - h_right (read)
  - v_top (read)
  - v_center (read)
  - ... and 3 more

**pylightxl**
- Read: 🔴 0 (0/3)
- Write: 🔴 0 (0/3)
- Failed tests (18):
  - h_left (read)
  - h_center (read)
  - h_right (read)
  - v_top (read)
  - v_center (read)
  - ... and 13 more

**python-calamine**
- Read: 🔴 0 (0/3)
- Failed tests (9):
  - h_left (read)
  - h_center (read)
  - h_right (read)
  - v_top (read)
  - v_center (read)
  - ... and 4 more

**xlsxwriter**
- Write: 🟠 1 (1/3)
- Failed tests (1):
  - v_bottom (write)

### background_colors

**openpyxl**
- Read: 🟢 3 (3/3)
- Write: 🟢 3 (3/3)

**pylightxl**
- Read: 🔴 0 (0/3)
- Write: 🔴 0 (0/3)
- Failed tests (8):
  - bg_red (read)
  - bg_blue (read)
  - bg_green (read)
  - bg_custom (read)
  - bg_red (write)
  - ... and 3 more

**python-calamine**
- Read: 🔴 0 (0/3)
- Failed tests (4):
  - bg_red (read)
  - bg_blue (read)
  - bg_green (read)
  - bg_custom (read)

**xlsxwriter**
- Write: 🟢 3 (3/3)

### borders

**openpyxl**
- Read: 🟢 3 (3/3)
- Write: 🟠 1 (1/3)
- Failed tests (1):
  - mixed_colors (write)

**pylightxl**
- Read: 🔴 0 (0/3)
- Write: 🔴 0 (0/3)
- Failed tests (40):
  - thin_all (read)
  - medium_all (read)
  - thick_all (read)
  - double (read)
  - dashed (read)
  - ... and 35 more

**python-calamine**
- Read: 🔴 0 (0/3)
- Failed tests (20):
  - thin_all (read)
  - medium_all (read)
  - thick_all (read)
  - double (read)
  - dashed (read)
  - ... and 15 more

**xlsxwriter**
- Write: 🟠 1 (1/3)
- Failed tests (3):
  - diagonal_up (write)
  - diagonal_down (write)
  - mixed_colors (write)

### cell_values

**openpyxl**
- Read: 🟢 3 (3/3)
- Write: 🟢 3 (3/3)

**pylightxl**
- Read: 🟠 1 (1/3)
- Write: 🔴 0 (0/3)
- Failed tests (21):
  - date_standard (read)
  - datetime (read)
  - error_na (read)
  - string_simple (write)
  - string_unicode (write)
  - ... and 16 more

**python-calamine**
- Read: 🟠 1 (1/3)
- Failed tests (3):
  - error_div0 (read)
  - error_na (read)
  - error_value (read)

**xlsxwriter**
- Write: 🟢 3 (3/3)

### comments

**openpyxl**
- Read: 🟢 3 (3/3)
- Write: 🟢 3 (3/3)

**pylightxl**
- Read: 🔴 0 (0/3)
- Write: 🔴 0 (0/3)
- Failed tests (6):
  - comment_legacy (read)
  - comment_threaded (read)
  - comment_author (read)
  - comment_legacy (write)
  - comment_threaded (write)
  - ... and 1 more

**python-calamine**
- Read: 🔴 0 (0/3)
- Failed tests (3):
  - comment_legacy (read)
  - comment_threaded (read)
  - comment_author (read)

**xlsxwriter**
- Write: 🟢 3 (3/3)

### conditional_formatting

**openpyxl**
- Read: 🟢 3 (3/3)
- Write: 🔴 0 (0/3)
- Failed tests (6):
  - cf_cell_gt (write)
  - cf_formula_cross_sheet (write)
  - cf_text_contains (write)
  - cf_data_bar (write)
  - cf_color_scale (write)
  - ... and 1 more

**pylightxl**
- Read: 🔴 0 (0/3)
- Write: 🔴 0 (0/3)
- Failed tests (12):
  - cf_cell_gt (read)
  - cf_formula_cross_sheet (read)
  - cf_text_contains (read)
  - cf_data_bar (read)
  - cf_color_scale (read)
  - ... and 7 more

**python-calamine**
- Read: 🔴 0 (0/3)
- Failed tests (6):
  - cf_cell_gt (read)
  - cf_formula_cross_sheet (read)
  - cf_text_contains (read)
  - cf_data_bar (read)
  - cf_color_scale (read)
  - ... and 1 more

**xlsxwriter**
- Write: 🔴 0 (0/3)
- Failed tests (4):
  - cf_cell_gt (write)
  - cf_formula_cross_sheet (write)
  - cf_text_contains (write)
  - cf_stop_if_true (write)

### data_validation

**openpyxl**
- Read: 🟢 3 (3/3)
- Write: 🟢 3 (3/3)

**pylightxl**
- Read: 🔴 0 (0/3)
- Write: 🔴 0 (0/3)
- Failed tests (10):
  - dv_list_csv (read)
  - dv_list_range (read)
  - dv_cross_sheet (read)
  - dv_custom_formula (read)
  - dv_whole_between (read)
  - ... and 5 more

**python-calamine**
- Read: 🔴 0 (0/3)
- Failed tests (5):
  - dv_list_csv (read)
  - dv_list_range (read)
  - dv_cross_sheet (read)
  - dv_custom_formula (read)
  - dv_whole_between (read)

**xlsxwriter**
- Write: 🟠 1 (1/3)
- Failed tests (1):
  - dv_whole_between (write)

### dimensions

**openpyxl**
- Read: 🟠 1 (1/3)
- Write: 🟢 3 (3/3)
- Failed tests (2):
  - col_width_20 (read)
  - col_width_8 (read)

**pylightxl**
- Read: 🔴 0 (0/3)
- Write: 🔴 0 (0/3)
- Failed tests (8):
  - row_height_30 (read)
  - row_height_45 (read)
  - col_width_20 (read)
  - col_width_8 (read)
  - row_height_30 (write)
  - ... and 3 more

**python-calamine**
- Read: 🔴 0 (0/3)
- Failed tests (4):
  - row_height_30 (read)
  - row_height_45 (read)
  - col_width_20 (read)
  - col_width_8 (read)

**xlsxwriter**
- Write: 🟠 1 (1/3)
- Failed tests (2):
  - col_width_20 (write)
  - col_width_8 (write)

### formulas

**openpyxl**
- Read: 🟠 1 (1/3)
- Write: 🟢 3 (3/3)
- Failed tests (1):
  - formula_cross_sheet (read)

**pylightxl**
- Read: 🔴 0 (0/3)
- Write: 🟢 3 (3/3)
- Failed tests (4):
  - formula_sum (read)
  - formula_cell_ref (read)
  - formula_concat (read)
  - formula_cross_sheet (read)

**python-calamine**
- Read: 🔴 0 (0/3)
- Failed tests (4):
  - formula_sum (read)
  - formula_cell_ref (read)
  - formula_concat (read)
  - formula_cross_sheet (read)

**xlsxwriter**
- Write: 🟢 3 (3/3)

### freeze_panes

**openpyxl**
- Read: 🟢 3 (3/3)
- Write: 🔴 0 (0/3)
- Failed tests (3):
  - freeze_b2 (write)
  - freeze_d5 (write)
  - split_2x1 (write)

**pylightxl**
- Read: 🔴 0 (0/3)
- Write: 🔴 0 (0/3)
- Failed tests (6):
  - freeze_b2 (read)
  - freeze_d5 (read)
  - split_2x1 (read)
  - freeze_b2 (write)
  - freeze_d5 (write)
  - ... and 1 more

**python-calamine**
- Read: 🔴 0 (0/3)
- Failed tests (3):
  - freeze_b2 (read)
  - freeze_d5 (read)
  - split_2x1 (read)

**xlsxwriter**
- Write: 🟡 2 (2/3)
- Failed tests (1):
  - split_2x1 (write)

### hyperlinks

**openpyxl**
- Read: 🟠 1 (1/3)
- Write: 🟢 3 (3/3)
- Failed tests (3):
  - link_internal (read)
  - link_mailto (read)
  - link_long (read)

**pylightxl**
- Read: 🔴 0 (0/3)
- Write: 🔴 0 (0/3)
- Failed tests (8):
  - link_external (read)
  - link_internal (read)
  - link_mailto (read)
  - link_long (read)
  - link_external (write)
  - ... and 3 more

**python-calamine**
- Read: 🔴 0 (0/3)
- Failed tests (4):
  - link_external (read)
  - link_internal (read)
  - link_mailto (read)
  - link_long (read)

**xlsxwriter**
- Write: 🟡 2 (2/3)
- Failed tests (1):
  - link_long (write)

### images

**openpyxl**
- Read: 🟢 3 (3/3)
- Write: 🟢 3 (3/3)

**pylightxl**
- Read: 🔴 0 (0/3)
- Write: 🔴 0 (0/3)
- Failed tests (4):
  - image_one_cell (read)
  - image_two_cell_offset (read)
  - image_one_cell (write)
  - image_two_cell_offset (write)

**python-calamine**
- Read: 🔴 0 (0/3)
- Failed tests (2):
  - image_one_cell (read)
  - image_two_cell_offset (read)

**xlsxwriter**
- Write: 🟢 3 (3/3)

### merged_cells

**openpyxl**
- Read: 🟢 3 (3/3)
- Write: 🟢 3 (3/3)

**pylightxl**
- Read: 🔴 0 (0/3)
- Write: 🔴 0 (0/3)
- Failed tests (8):
  - merge_horizontal (read)
  - merge_vertical (read)
  - merge_value_off_top_left (read)
  - merge_top_left_fill (read)
  - merge_horizontal (write)
  - ... and 3 more

**python-calamine**
- Read: 🔴 0 (0/3)
- Failed tests (4):
  - merge_horizontal (read)
  - merge_vertical (read)
  - merge_value_off_top_left (read)
  - merge_top_left_fill (read)

**xlsxwriter**
- Write: 🟢 3 (3/3)

### multiple_sheets

**openpyxl**
- Read: 🟢 3 (3/3)
- Write: 🟢 3 (3/3)

**pylightxl**
- Read: 🟢 3 (3/3)
- Write: 🟢 3 (3/3)

**python-calamine**
- Read: 🟢 3 (3/3)

**xlsxwriter**
- Write: 🟢 3 (3/3)

### number_formats

**openpyxl**
- Read: 🟠 1 (1/3)
- Write: 🟢 3 (3/3)
- Failed tests (3):
  - numfmt_currency (read)
  - numfmt_date (read)
  - numfmt_custom_text (read)

**pylightxl**
- Read: 🔴 0 (0/3)
- Write: 🔴 0 (0/3)
- Failed tests (10):
  - numfmt_currency (read)
  - numfmt_percent (read)
  - numfmt_date (read)
  - numfmt_scientific (read)
  - numfmt_custom_text (read)
  - ... and 5 more

**python-calamine**
- Read: 🔴 0 (0/3)
- Failed tests (5):
  - numfmt_currency (read)
  - numfmt_percent (read)
  - numfmt_date (read)
  - numfmt_scientific (read)
  - numfmt_custom_text (read)

**xlsxwriter**
- Write: 🟢 3 (3/3)

### pivot_tables

**openpyxl**
- Notes: Unsupported on macOS without a Windows-generated pivot fixture (fixtures/excel/tier2/15_pivot_tables.xlsx).

**pylightxl**
- Notes: Unsupported on macOS without a Windows-generated pivot fixture (fixtures/excel/tier2/15_pivot_tables.xlsx).

**python-calamine**
- Notes: Unsupported on macOS without a Windows-generated pivot fixture (fixtures/excel/tier2/15_pivot_tables.xlsx).

**xlsxwriter**
- Notes: Unsupported on macOS without a Windows-generated pivot fixture (fixtures/excel/tier2/15_pivot_tables.xlsx).

### text_formatting

**openpyxl**
- Read: 🟢 3 (3/3)
- Write: 🟢 3 (3/3)

**pylightxl**
- Read: 🔴 0 (0/3)
- Write: 🔴 0 (0/3)
- Failed tests (36):
  - bold (read)
  - italic (read)
  - underline_single (read)
  - underline_double (read)
  - strikethrough (read)
  - ... and 31 more

**python-calamine**
- Read: 🔴 0 (0/3)
- Failed tests (18):
  - bold (read)
  - italic (read)
  - underline_single (read)
  - underline_double (read)
  - strikethrough (read)
  - ... and 13 more

**xlsxwriter**
- Write: 🟢 3 (3/3)

---
*Benchmark version: 0.1.0*