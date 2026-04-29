# ExcelBench Results

*Generated: 2026-04-29 14:26 UTC*
*Profile: xlsx*
*Excel Version: 16.105.3*
*Platform: Darwin-arm64*

## Overview

> Condensed view — shows the **best score** across read/write for each library. See [Full Results Matrix](#full-results-matrix) for the complete R/W breakdown.

**Tier 0 — Basic Values**

| Feature | apache-poi | excelize |
|---------|:-:|:-:|
| Cell Values | 🟢 | 🟢 |
| Formulas | 🟢 | 🟢 |
| Sheets | 🟢 | 🟢 |

**Tier 1 — Formatting**

| Feature | apache-poi | excelize |
|---------|:-:|:-:|
| Alignment | 🟢 | 🟢 |
| Bg Colors | 🟢 | 🟢 |
| Borders | 🟢 | 🟢 |
| Dimensions | 🟢 | 🟢 |
| Num Fmt | 🟢 | 🟢 |
| Text Fmt | 🟢 | 🟢 |

**Tier 2 — Advanced**

| Feature | apache-poi | excelize |
|---------|:-:|:-:|
| Comments | 🟢 | 🟢 |
| Cond Fmt | 🟢 | 🟢 |
| Validation | 🟢 | 🟢 |
| Freeze | 🟢 | 🟢 |
| Hyperlinks | 🟢 | 🟢 |
| Images | 🟢 | 🟢 |
| Merged | 🟢 | 🟢 |

**Tier 3 — Workbook Metadata**

| Feature | apache-poi | excelize |
|---------|:-:|:-:|
| Named Ranges | 🟢 | 🟢 |
| Tables | 🟢 | 🟢 |

## Library Tiers

> Libraries ranked by their best capability (max of read/write green features).

| Tier | Library | Caps | Modify | Green Features | Summary |
|:----:|---------|:----:|:------:|:--------------:|---------|
| **S** | apache-poi | W | No | 18/18 | 18/18 features with full fidelity |
| **S** | excelize | W | No | 18/18 | 18/18 features with full fidelity |

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

| Feature | apache-poi (W) | excelize (W) |
|---------|------------|------------|
| [cell_values](#cell_values-details) | 🟢 3 | 🟢 3 |
| [formulas](#formulas-details) | 🟢 3 | 🟢 3 |
| [multiple_sheets](#multiple_sheets-details) | 🟢 3 | 🟢 3 |

**Tier 1 — Formatting**

| Feature | apache-poi (W) | excelize (W) |
|---------|------------|------------|
| [alignment](#alignment-details) | 🟢 3 | 🟢 3 |
| [background_colors](#background_colors-details) | 🟢 3 | 🟢 3 |
| [borders](#borders-details) | 🟢 3 | 🟢 3 |
| [dimensions](#dimensions-details) | 🟢 3 | 🟢 3 |
| [number_formats](#number_formats-details) | 🟢 3 | 🟢 3 |
| [text_formatting](#text_formatting-details) | 🟢 3 | 🟢 3 |

**Tier 2 — Advanced**

| Feature | apache-poi (W) | excelize (W) |
|---------|------------|------------|
| [comments](#comments-details) | 🟢 3 | 🟢 3 |
| [conditional_formatting](#conditional_formatting-details) | 🟢 3 | 🟢 3 |
| [data_validation](#data_validation-details) | 🟢 3 | 🟢 3 |
| [freeze_panes](#freeze_panes-details) | 🟢 3 | 🟢 3 |
| [hyperlinks](#hyperlinks-details) | 🟢 3 | 🟢 3 |
| [images](#images-details) | 🟢 3 | 🟢 3 |
| [merged_cells](#merged_cells-details) | 🟢 3 | 🟢 3 |
| [pivot_tables](#pivot_tables-details) | ➖ | ➖ |

**Tier 3 — Workbook Metadata**

| Feature | apache-poi (W) | excelize (W) |
|---------|------------|------------|
| [named_ranges](#named_ranges-details) | 🟢 3 | 🟢 3 |
| [tables](#tables-details) | 🟢 3 | 🟢 3 |

## Notes

- **pivot_tables**: Unsupported on macOS without a Windows-generated pivot fixture (fixtures/excel/tier2/15_pivot_tables.xlsx).

## Statistics

| Library | Mode | Tests | Passed | Failed | Pass Rate | Green Features |
|---------|------|-------|--------|--------|-----------|----------------|
| apache-poi | W | 125 | 125 | 0 | 100% | 18/18 |
| excelize | W | 125 | 125 | 0 | 100% | 18/18 |

## Libraries Tested

- **apache-poi** v5.5.1 (java) - write; modify: No
- **excelize** vgo-helper (go) - write; modify: No

## Diagnostics Summary

No diagnostics recorded.

## Detailed Results

<a id="alignment-details"></a>
### alignment

**apache-poi** — Write: 🟢 3

**excelize** — Write: 🟢 3

<a id="background_colors-details"></a>
### background_colors

**apache-poi** — Write: 🟢 3

**excelize** — Write: 🟢 3

<a id="borders-details"></a>
### borders

**apache-poi** — Write: 🟢 3

**excelize** — Write: 🟢 3

<a id="cell_values-details"></a>
### cell_values

**apache-poi** — Write: 🟢 3

**excelize** — Write: 🟢 3

<a id="comments-details"></a>
### comments

**apache-poi** — Write: 🟢 3

**excelize** — Write: 🟢 3

<a id="conditional_formatting-details"></a>
### conditional_formatting

**apache-poi** — Write: 🟢 3

**excelize** — Write: 🟢 3

<a id="data_validation-details"></a>
### data_validation

**apache-poi** — Write: 🟢 3

**excelize** — Write: 🟢 3

<a id="dimensions-details"></a>
### dimensions

**apache-poi** — Write: 🟢 3

**excelize** — Write: 🟢 3

<a id="formulas-details"></a>
### formulas

**apache-poi** — Write: 🟢 3

**excelize** — Write: 🟢 3

<a id="freeze_panes-details"></a>
### freeze_panes

**apache-poi** — Write: 🟢 3

**excelize** — Write: 🟢 3

<a id="hyperlinks-details"></a>
### hyperlinks

**apache-poi** — Write: 🟢 3

**excelize** — Write: 🟢 3

<a id="images-details"></a>
### images

**apache-poi** — Write: 🟢 3

**excelize** — Write: 🟢 3

<a id="merged_cells-details"></a>
### merged_cells

**apache-poi** — Write: 🟢 3

**excelize** — Write: 🟢 3

<a id="multiple_sheets-details"></a>
### multiple_sheets

**apache-poi** — Write: 🟢 3

**excelize** — Write: 🟢 3

<a id="named_ranges-details"></a>
### named_ranges

**apache-poi** — Write: 🟢 3

**excelize** — Write: 🟢 3

<a id="number_formats-details"></a>
### number_formats

**apache-poi** — Write: 🟢 3

**excelize** — Write: 🟢 3

<a id="pivot_tables-details"></a>
### pivot_tables

**apache-poi**
- Notes: Unsupported on macOS without a Windows-generated pivot fixture (fixtures/excel/tier2/15_pivot_tables.xlsx).

**excelize**
- Notes: Unsupported on macOS without a Windows-generated pivot fixture (fixtures/excel/tier2/15_pivot_tables.xlsx).

<a id="tables-details"></a>
### tables

**apache-poi** — Write: 🟢 3

**excelize** — Write: 🟢 3

<a id="text_formatting-details"></a>
### text_formatting

**apache-poi** — Write: 🟢 3

**excelize** — Write: 🟢 3

---
*Benchmark version: 0.1.0*