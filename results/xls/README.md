# ExcelBench Results

*Generated: 2026-02-08 22:56 UTC*
*Profile: xls*
*Excel Version: xlwt*
*Platform: Darwin-arm64*

## Overview

> Condensed view — shows the **best score** across read/write for each library. See [Full Results Matrix](#full-results-matrix) for the complete R/W breakdown.

**Tier 0 — Basic Values**

| Feature | calamine | xlrd |
|---------|:-:|:-:|
| Cell Values | 🟢 | 🟢 |
| Sheets | 🟢 | 🟢 |

**Tier 1 — Formatting**

| Feature | calamine | xlrd |
|---------|:-:|:-:|
| Alignment | 🟠 | 🟢 |
| Dimensions | 🔴 | 🟢 |

## Library Tiers

> Libraries ranked by their best capability (max of read/write green features).

| Tier | Library | Caps | Green Features | Summary |
|:----:|---------|:----:|:--------------:|---------|
| **S** | xlrd | R | 4/4 | Legacy .xls reader — not applicable to .xlsx |
| **C** | python-calamine | R | 2/4 | Fast Rust-backed reader — cell values + sheet names only |

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

| Feature | python-calamine (R) | xlrd (R) |
|---------|------------|------------|
| [cell_values](#cell_values-details) | 🟢 3 | 🟢 3 |
| [multiple_sheets](#multiple_sheets-details) | 🟢 3 | 🟢 3 |

**Tier 1 — Formatting**

| Feature | python-calamine (R) | xlrd (R) |
|---------|------------|------------|
| [alignment](#alignment-details) | 🟠 1 | 🟢 3 |
| [dimensions](#dimensions-details) | 🔴 0 | 🟢 3 |

## Notes

- **alignment**: Known limitation: python-calamine alignment read is limited because its API does not expose style/alignment metadata.

## Statistics

| Library | Mode | Tests | Passed | Failed | Pass Rate | Green Features |
|---------|------|-------|--------|--------|-----------|----------------|
| python-calamine | R | 35 | 23 | 12 | 66% | 2/4 |
| xlrd | R | 35 | 35 | 0 | 100% | 4/4 |

## Libraries Tested

- **python-calamine** v0.6.1 (python) - read
- **xlrd** v2.0.2 (python) - read

## Diagnostics Summary

No diagnostics recorded.

## Detailed Results

<a id="alignment-details"></a>
### alignment

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

**python-calamine** — Read: 🟢 3

**xlrd** — Read: 🟢 3

<a id="dimensions-details"></a>
### dimensions

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

**python-calamine** — Read: 🟢 3

**xlrd** — Read: 🟢 3

---
*Benchmark version: 0.1.0*