# ExcelBench Results

*Generated: 2026-02-14 00:04 UTC*
*Profile: xls*
*Excel Version: xlwt*
*Platform: Darwin-arm64*

## Overview

> Condensed view — shows the **best score** across read/write for each library. See [Full Results Matrix](#full-results-matrix) for the complete R/W breakdown.

**Tier 0 — Basic Values**

| Feature | calamine | calamine | xlrd |
|---------|:-:|:-:|:-:|
| Cell Values | 🟠 | 🟢 | 🟢 |
| Sheets | 🟠 | 🟢 | 🟢 |

**Tier 1 — Formatting**

| Feature | calamine | calamine | xlrd |
|---------|:-:|:-:|:-:|
| Alignment | 🟠 | 🟠 | 🟢 |
| Dimensions | 🔴 | 🔴 | 🟢 |

## Library Tiers

> Libraries ranked by their best capability (max of read/write green features).

| Tier | Library | Caps | Green Features | Summary |
|:----:|---------|:----:|:--------------:|---------|
| **S** | xlrd | R | 4/4 | Legacy .xls reader — not applicable to .xlsx |
| **C** | python-calamine | R | 2/4 | Fast Rust-backed reader — cell values + sheet names only |
| **D** | calamine | R | 0/4 | 0/4 features with full fidelity |

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

| Feature | calamine (R) | python-calamine (R) | xlrd (R) |
|---------|------------|------------|------------|
| [cell_values](#cell_values-details) | 🟠 1 | 🟢 3 | 🟢 3 |
| [multiple_sheets](#multiple_sheets-details) | 🟠 1 | 🟢 3 | 🟢 3 |

**Tier 1 — Formatting**

| Feature | calamine (R) | python-calamine (R) | xlrd (R) |
|---------|------------|------------|------------|
| [alignment](#alignment-details) | 🟠 1 | 🟠 1 | 🟢 3 |
| [dimensions](#dimensions-details) | 🔴 0 | 🔴 0 | 🟢 3 |

## Notes

- **alignment**: Known limitation: python-calamine alignment read is limited because its API does not expose style/alignment metadata.

## Statistics

| Library | Mode | Tests | Passed | Failed | Pass Rate | Green Features |
|---------|------|-------|--------|--------|-----------|----------------|
| calamine | R | 35 | 19 | 16 | 54% | 0/4 |
| python-calamine | R | 35 | 23 | 12 | 66% | 2/4 |
| xlrd | R | 35 | 35 | 0 | 100% | 4/4 |

## Libraries Tested

- **calamine** v0.25.0 (rust) - read
- **python-calamine** v0.6.1 (python) - read
- **xlrd** v2.0.2 (python) - read

## Diagnostics Summary

| Group | Value | Count |
|-------|-------|-------|
| category | data_mismatch | 28 |
| severity | error | 28 |

### Diagnostic Details

| Feature | Library | Test Case | Operation | Category | Severity | Message |
|---------|---------|-----------|-----------|----------|----------|---------|
| cell_values | calamine | error_div0 | read | data_mismatch | error | Expected values did not match actual values: expected={'type': 'error', 'value': '#DIV/0!'}, actual={'type': 'string', 'value': '#DIV/0!'} |
| cell_values | calamine | error_na | read | data_mismatch | error | Expected values did not match actual values: expected={'type': 'error', 'value': '#N/A'}, actual={'type': 'string', 'value': '#N/A'} |
| cell_values | calamine | error_value | read | data_mismatch | error | Expected values did not match actual values: expected={'type': 'error', 'value': '#VALUE!'}, actual={'type': 'string', 'value': '#VALUE!'} |
| alignment | python-calamine | h_left | read | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'left'}, actual={'h_align': 'general', 'v_align': 'bottom'} |
| alignment | python-calamine | h_center | read | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'center'}, actual={'h_align': 'general', 'v_align': 'bottom'} |
| alignment | python-calamine | h_right | read | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'right'}, actual={'h_align': 'general', 'v_align': 'bottom'} |
| alignment | python-calamine | v_top | read | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'top'}, actual={'h_align': 'general', 'v_align': 'bottom'} |
| alignment | python-calamine | v_center | read | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'center'}, actual={'h_align': 'general', 'v_align': 'bottom'} |
| alignment | python-calamine | wrap_text | read | data_mismatch | error | Expected values did not match actual values: expected={'wrap': True}, actual={'h_align': 'general', 'v_align': 'bottom'} |
| alignment | python-calamine | rotation_45 | read | data_mismatch | error | Expected values did not match actual values: expected={'rotation': 45}, actual={'h_align': 'general', 'v_align': 'bottom'} |
| alignment | python-calamine | indent_2 | read | data_mismatch | error | Expected values did not match actual values: expected={'indent': 2}, actual={'h_align': 'general', 'v_align': 'bottom'} |
| alignment | calamine | h_left | read | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'left'}, actual={'h_align': 'general', 'v_align': 'bottom'} |
| alignment | calamine | h_center | read | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'center'}, actual={'h_align': 'general', 'v_align': 'bottom'} |
| alignment | calamine | h_right | read | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'right'}, actual={'h_align': 'general', 'v_align': 'bottom'} |
| alignment | calamine | v_top | read | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'top'}, actual={'h_align': 'general', 'v_align': 'bottom'} |
| alignment | calamine | v_center | read | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'center'}, actual={'h_align': 'general', 'v_align': 'bottom'} |
| alignment | calamine | wrap_text | read | data_mismatch | error | Expected values did not match actual values: expected={'wrap': True}, actual={'h_align': 'general', 'v_align': 'bottom'} |
| alignment | calamine | rotation_45 | read | data_mismatch | error | Expected values did not match actual values: expected={'rotation': 45}, actual={'h_align': 'general', 'v_align': 'bottom'} |
| alignment | calamine | indent_2 | read | data_mismatch | error | Expected values did not match actual values: expected={'indent': 2}, actual={'h_align': 'general', 'v_align': 'bottom'} |
| dimensions | python-calamine | row_height_30 | read | data_mismatch | error | Expected values did not match actual values: expected={'row_height': 30}, actual={'row_height': None} |
| dimensions | python-calamine | row_height_45 | read | data_mismatch | error | Expected values did not match actual values: expected={'row_height': 45}, actual={'row_height': None} |
| dimensions | python-calamine | col_width_20 | read | data_mismatch | error | Expected values did not match actual values: expected={'column_width': 20}, actual={'column_width': None} |
| dimensions | python-calamine | col_width_8 | read | data_mismatch | error | Expected values did not match actual values: expected={'column_width': 8}, actual={'column_width': None} |
| dimensions | calamine | row_height_30 | read | data_mismatch | error | Expected values did not match actual values: expected={'row_height': 30}, actual={'row_height': None} |
| dimensions | calamine | row_height_45 | read | data_mismatch | error | Expected values did not match actual values: expected={'row_height': 45}, actual={'row_height': None} |
| dimensions | calamine | col_width_20 | read | data_mismatch | error | Expected values did not match actual values: expected={'column_width': 20}, actual={'column_width': None} |
| dimensions | calamine | col_width_8 | read | data_mismatch | error | Expected values did not match actual values: expected={'column_width': 8}, actual={'column_width': None} |
| multiple_sheets | calamine | value_alpha | read | data_mismatch | error | Expected values did not match actual values: expected={'type': 'string', 'value': 'Alpha'}, actual={'type': 'blank'} |

## Detailed Results

<a id="alignment-details"></a>
### alignment

**calamine** — Read: 🟠 1

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

**xlrd** — Read: 🟢 3

<a id="cell_values-details"></a>
### cell_values

**calamine** — Read: 🟠 1

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

**python-calamine** — Read: 🟢 3

**xlrd** — Read: 🟢 3

<a id="dimensions-details"></a>
### dimensions

**calamine** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Row height - 30 | basic | ❌ |
| Row height - 45 | basic | ❌ |
| Column width - D = 20 | basic | ❌ |
| Column width - E = 8 | basic | ❌ |

**python-calamine** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Row height - 30 | basic | ❌ |
| Row height - 45 | basic | ❌ |
| Column width - D = 20 | basic | ❌ |
| Column width - E = 8 | basic | ❌ |

**xlrd** — Read: 🟢 3

<a id="multiple_sheets-details"></a>
### multiple_sheets

**calamine** — Read: 🟠 1

| Test | Importance | Read |
|------|-----------|------|
| Sheet names | basic | ✅ |
| Alpha value | basic | ❌ |
| Beta value | basic | ✅ |
| Gamma value | basic | ✅ |

**python-calamine** — Read: 🟢 3

**xlrd** — Read: 🟢 3

---
*Benchmark version: 0.1.0*