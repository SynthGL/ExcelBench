# ExcelBench Performance Results

*Generated: 2026-02-08T23:00:06.370307+00:00*
*Profile: xlsx*
*Platform: Darwin-arm64*
*Python: 3.12.3*
*Commit: 83df991*
*Config: warmup=3 iters=25 breakdown=True*

## Notes

These numbers measure only the library under test. Write timings do NOT include oracle verification.

## Summary (p50 wall time)

**Tier 0 — Basic Values**

| Feature | openpyxl (R p50 ms) | openpyxl (W p50 ms) | openpyxl-readonly (R p50 ms) | pandas (R p50 ms) | pandas (W p50 ms) | polars (R p50 ms) | pyexcel (R p50 ms) | pyexcel (W p50 ms) | pylightxl (R p50 ms) | pylightxl (W p50 ms) | python-calamine (R p50 ms) | tablib (R p50 ms) | tablib (W p50 ms) | xlrd (R p50 ms) | xlsxwriter (W p50 ms) | xlsxwriter-constmem (W p50 ms) | xlwt (W p50 ms) |
|---------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|
| cell_values | 1.42 | 1.91 | 5.14 | 1.87 | 1.83 | 0.39 | 1.66 | 1.66 | 1.46 | 0.42 | 0.95 | 1.41 | 1.58 | — | 2.62 | 2.16 | 0.26 |
| formulas | 1.24 | 1.65 | 1.47 | 1.74 | 2.00 | 0.44 | 1.31 | 1.77 | 1.22 | 0.31 | 0.16 | 1.19 | 1.63 | — | 2.02 | 2.19 | 0.29 |
| multiple_sheets | 1.27 | 1.92 | 1.08 | 1.76 | 2.48 | 0.61 | 1.42 | 2.02 | 1.37 | 0.36 | 0.07 | 1.34 | 2.06 | — | 2.29 | 2.93 | 0.19 |

**Tier 1 — Formatting**

| Feature | openpyxl (R p50 ms) | openpyxl (W p50 ms) | openpyxl-readonly (R p50 ms) | pandas (R p50 ms) | pandas (W p50 ms) | polars (R p50 ms) | pyexcel (R p50 ms) | pyexcel (W p50 ms) | pylightxl (R p50 ms) | pylightxl (W p50 ms) | python-calamine (R p50 ms) | tablib (R p50 ms) | tablib (W p50 ms) | xlrd (R p50 ms) | xlsxwriter (W p50 ms) | xlsxwriter-constmem (W p50 ms) | xlwt (W p50 ms) |
|---------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|
| alignment | 1.25 | 1.55 | 1.01 | 1.49 | 1.66 | 0.25 | 1.35 | 1.41 | — | 0.33 | 0.06 | 1.32 | 1.56 | — | 2.11 | 2.09 | 0.25 |
| background_colors | 1.08 | 1.52 | 0.93 | 1.31 | 1.64 | 0.23 | 1.17 | 1.38 | 1.02 | 0.29 | 0.06 | 1.10 | 1.38 | — | 2.22 | 2.07 | 0.20 |
| borders | 2.08 | 2.46 | 1.31 | 1.98 | 1.70 | 0.29 | 1.86 | 1.46 | — | 0.43 | 0.07 | 1.77 | 1.48 | — | 2.46 | 2.48 | 0.44 |
| dimensions | 0.98 | 1.36 | 0.85 | 1.19 | 1.34 | 0.32 | 1.03 | 1.26 | 0.96 | 0.25 | 0.06 | 0.95 | 1.22 | — | 1.69 | 1.91 | 0.14 |
| number_formats | 1.14 | 1.48 | 0.90 | 1.27 | 1.54 | 0.26 | 1.16 | 1.41 | 1.02 | 0.31 | 0.07 | 1.08 | 1.35 | — | 1.87 | 2.16 | 0.21 |
| text_formatting | 1.83 | 2.11 | 1.55 | 2.01 | 1.77 | 0.35 | 1.81 | 1.49 | 1.64 | 0.42 | 0.09 | 1.80 | 1.55 | — | 2.49 | 2.45 | 0.36 |

**Tier 2 — Advanced**

| Feature | openpyxl (R p50 ms) | openpyxl (W p50 ms) | openpyxl-readonly (R p50 ms) | pandas (R p50 ms) | pandas (W p50 ms) | polars (R p50 ms) | pyexcel (R p50 ms) | pyexcel (W p50 ms) | pylightxl (R p50 ms) | pylightxl (W p50 ms) | python-calamine (R p50 ms) | tablib (R p50 ms) | tablib (W p50 ms) | xlrd (R p50 ms) | xlsxwriter (W p50 ms) | xlsxwriter-constmem (W p50 ms) | xlwt (W p50 ms) |
|---------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|
| comments | 1.14 | 1.55 | 0.91 | 1.20 | 1.35 | 0.24 | 1.21 | 1.31 | 1.21 | 0.24 | 0.06 | 1.02 | 1.24 | — | 2.24 | 1.84 | 0.14 |
| conditional_formatting | 1.51 | 1.95 | 0.95 | 1.68 | 1.87 | 0.37 | 1.69 | 1.69 | — | 0.29 | 0.06 | 1.40 | 1.54 | — | 2.22 | 2.30 | 0.16 |
| data_validation | 1.26 | 1.40 | 0.90 | 1.62 | 1.38 | 0.48 | 1.42 | 1.33 | — | 0.26 | 0.06 | 1.26 | 1.25 | — | 1.86 | 1.86 | 0.15 |
| freeze_panes | 1.39 | 2.22 | 1.00 | 1.77 | 2.34 | 0.42 | 1.50 | 2.18 | — | 0.31 | 0.06 | 1.35 | 2.11 | — | 2.41 | 2.62 | 0.20 |
| hyperlinks | 1.23 | 1.39 | 0.89 | 1.52 | 1.35 | 0.37 | 1.29 | 1.24 | — | 0.24 | 0.06 | 1.20 | 1.22 | — | 2.18 | 2.04 | 0.13 |
| images | 1.22 | 1.83 | 0.83 | 1.17 | 1.34 | 0.24 | 1.30 | 1.26 | — | 0.24 | 0.06 | 0.96 | 1.22 | — | 2.91 | 1.79 | 0.13 |
| merged_cells | 1.27 | 1.59 | 0.98 | 1.33 | 1.62 | 0.31 | 1.38 | 1.33 | 1.06 | 0.31 | 0.07 | 1.07 | 1.35 | — | 2.00 | 2.10 | 0.21 |
| pivot_tables | 0.89 | 1.26 | 0.80 | 1.11 | 1.37 | 0.24 | 0.94 | 1.27 | 0.86 | 0.25 | 0.06 | 0.87 | 1.22 | — | 1.74 | 1.80 | 0.13 |

## Run Issues

- alignment / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'
- alignment / xlrd: Read not applicable: xlrd does not support .xlsx input
- background_colors / xlrd: Read not applicable: xlrd does not support .xlsx input
- borders / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'
- borders / xlrd: Read not applicable: xlrd does not support .xlsx input
- cell_values / xlrd: Read not applicable: xlrd does not support .xlsx input
- comments / xlrd: Read not applicable: xlrd does not support .xlsx input
- conditional_formatting / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'
- conditional_formatting / xlrd: Read not applicable: xlrd does not support .xlsx input
- data_validation / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'
- data_validation / xlrd: Read not applicable: xlrd does not support .xlsx input
- dimensions / xlrd: Read not applicable: xlrd does not support .xlsx input
- formulas / xlrd: Read not applicable: xlrd does not support .xlsx input
- freeze_panes / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'
- freeze_panes / xlrd: Read not applicable: xlrd does not support .xlsx input
- hyperlinks / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'
- hyperlinks / xlrd: Read not applicable: xlrd does not support .xlsx input
- images / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'
- images / xlrd: Read not applicable: xlrd does not support .xlsx input
- merged_cells / xlrd: Read not applicable: xlrd does not support .xlsx input
- multiple_sheets / xlrd: Read not applicable: xlrd does not support .xlsx input
- number_formats / xlrd: Read not applicable: xlrd does not support .xlsx input
- pivot_tables / xlrd: Read not applicable: xlrd does not support .xlsx input
- text_formatting / xlrd: Read not applicable: xlrd does not support .xlsx input
