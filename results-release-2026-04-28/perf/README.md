# ExcelBench Performance Results

*Generated: 2026-04-29T03:12:22.197235+00:00*
*Profile: xlsx*
*Platform: Darwin-arm64*
*Python: 3.12.3*
*Commit: 5f98de9*
*Config: warmup=3 iters=25 iteration_policy=fixed breakdown=False*

## Notes

These numbers measure only the library under test. Write timings do NOT include oracle verification.

## Summary (p50 wall time)

**Tier 0 — Basic Values**

| Feature | openpyxl (R p50 ms) | openpyxl (W p50 ms) | openpyxl-readonly (R p50 ms) | pandas (R p50 ms) | pandas (W p50 ms) | polars (R p50 ms) | pyexcel (R p50 ms) | pyexcel (W p50 ms) | pylightxl (R p50 ms) | pylightxl (W p50 ms) | python-calamine (R p50 ms) | tablib (R p50 ms) | tablib (W p50 ms) | wolfxl (R p50 ms) | wolfxl (W p50 ms) | xlrd (R p50 ms) | xlsxwriter (W p50 ms) | xlsxwriter-constmem (W p50 ms) | xlwt (W p50 ms) |
|---------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|
| cell_values | 1.36 | 1.42 | 5.00 | 1.79 | 1.79 | 0.33 | 1.51 | 1.46 | 1.41 | 0.39 | 0.93 | 1.37 | 1.52 | 0.24 | 0.19 | — | 1.90 | 1.85 | 0.24 |
| formulas | 1.17 | 1.53 | 1.44 | 1.64 | 1.93 | 0.41 | 1.30 | 1.65 | 1.20 | 0.31 | 0.16 | 1.30 | 1.62 | 0.19 | 0.19 | — | 1.78 | 2.03 | 0.28 |
| multiple_sheets | 1.33 | 1.82 | 1.02 | 1.84 | 2.56 | 0.62 | 1.45 | 1.96 | 1.39 | 0.35 | 0.07 | 1.36 | 2.04 | 0.07 | 0.19 | — | 2.29 | 2.61 | 0.19 |

**Tier 1 — Formatting**

| Feature | openpyxl (R p50 ms) | openpyxl (W p50 ms) | openpyxl-readonly (R p50 ms) | pandas (R p50 ms) | pandas (W p50 ms) | polars (R p50 ms) | pyexcel (R p50 ms) | pyexcel (W p50 ms) | pylightxl (R p50 ms) | pylightxl (W p50 ms) | python-calamine (R p50 ms) | tablib (R p50 ms) | tablib (W p50 ms) | wolfxl (R p50 ms) | wolfxl (W p50 ms) | xlrd (R p50 ms) | xlsxwriter (W p50 ms) | xlsxwriter-constmem (W p50 ms) | xlwt (W p50 ms) |
|---------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|
| alignment | 1.28 | 1.50 | 0.96 | 1.49 | 1.70 | 0.27 | 1.43 | 1.50 | — | 0.34 | 0.06 | 1.30 | 1.37 | 0.29 | 0.19 | — | 2.18 | 1.87 | 0.24 |
| background_colors | 1.08 | 1.51 | 0.91 | 1.41 | 1.82 | 0.33 | 1.17 | 1.40 | 1.03 | 0.33 | 0.06 | 1.09 | 1.50 | 0.27 | 0.17 | — | 2.28 | 2.03 | 0.21 |
| borders | 1.85 | 2.45 | 1.30 | 2.08 | 1.83 | 0.29 | 1.83 | 1.54 | — | 0.42 | 0.07 | 1.78 | 1.56 | 0.35 | 0.27 | — | 2.36 | 2.55 | 0.43 |
| dimensions | 0.99 | 1.29 | 0.87 | 1.28 | 1.41 | 0.32 | 1.09 | 1.30 | 0.96 | 0.27 | 0.06 | 1.00 | 1.33 | 0.08 | 0.16 | — | 1.71 | 1.78 | 0.15 |
| number_formats | 1.18 | 1.70 | 0.91 | 1.51 | 1.98 | 0.26 | 1.25 | 1.46 | 1.07 | 0.46 | 0.07 | 1.08 | 1.35 | 0.32 | 0.18 | — | 2.09 | 2.08 | 0.22 |
| text_formatting | 1.74 | 2.03 | 1.50 | 2.11 | 1.77 | 0.34 | 1.93 | 1.44 | 1.55 | 0.41 | 0.09 | 1.91 | 1.72 | 0.36 | 0.23 | — | 2.39 | 2.28 | 0.35 |

**Tier 2 — Advanced**

| Feature | openpyxl (R p50 ms) | openpyxl (W p50 ms) | openpyxl-readonly (R p50 ms) | pandas (R p50 ms) | pandas (W p50 ms) | polars (R p50 ms) | pyexcel (R p50 ms) | pyexcel (W p50 ms) | pylightxl (R p50 ms) | pylightxl (W p50 ms) | python-calamine (R p50 ms) | tablib (R p50 ms) | tablib (W p50 ms) | wolfxl (R p50 ms) | wolfxl (W p50 ms) | xlrd (R p50 ms) | xlsxwriter (W p50 ms) | xlsxwriter-constmem (W p50 ms) | xlwt (W p50 ms) |
|---------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|
| comments | 1.15 | 1.59 | 0.92 | 1.36 | 1.61 | 0.25 | 1.28 | 1.39 | 1.20 | 0.26 | 0.07 | 1.06 | 1.34 | 0.21 | 0.20 | — | 2.26 | 1.81 | 0.13 |
| conditional_formatting | 1.42 | 1.84 | 0.96 | 1.80 | 2.16 | 0.40 | 1.53 | 1.56 | — | 0.31 | 0.06 | 1.42 | 1.51 | 0.22 | 0.19 | — | 1.87 | 2.03 | 0.16 |
| data_validation | 1.33 | 1.38 | 0.97 | 2.29 | 1.87 | 0.49 | 1.51 | 1.61 | — | 0.27 | 0.05 | 1.37 | 1.42 | 0.16 | 0.16 | — | 1.65 | 1.99 | 0.15 |
| freeze_panes | 1.45 | 2.32 | 1.02 | 2.00 | 2.46 | 0.51 | 1.74 | 2.49 | — | 0.32 | 0.06 | 1.45 | 2.58 | 0.24 | 0.20 | — | 2.29 | 2.56 | 0.19 |
| hyperlinks | 1.39 | 1.50 | 0.97 | 1.69 | 1.54 | 0.47 | 1.55 | 1.45 | — | 0.28 | 0.06 | 1.24 | 1.49 | 0.25 | 0.18 | — | 2.20 | 2.54 | 0.13 |
| images | 1.28 | 2.03 | 0.83 | 1.27 | 1.43 | 0.29 | 1.31 | 1.38 | — | 0.27 | 0.06 | 0.98 | 1.34 | 0.46 | 0.28 | — | 3.07 | 1.67 | 0.14 |
| merged_cells | 1.28 | 1.47 | 0.91 | 1.40 | 1.72 | 0.26 | 1.40 | 1.36 | 1.03 | 0.34 | 0.06 | 1.10 | 1.35 | 0.30 | 0.16 | — | 1.85 | 1.91 | 0.20 |
| named_ranges | 1.24 | 1.59 | 0.89 | 1.81 | 1.97 | 0.43 | 1.34 | 1.77 | — | 0.28 | 0.05 | 1.27 | 1.78 | 0.06 | 0.17 | — | 2.01 | 2.16 | 0.16 |
| pivot_tables | 0.90 | 1.29 | 0.82 | 1.24 | 1.47 | 0.25 | 0.97 | 1.37 | 0.86 | 0.28 | 0.06 | 0.95 | 1.26 | 0.06 | 0.15 | — | 1.72 | 1.73 | 0.14 |
| tables | 1.66 | 1.43 | 0.92 | 1.96 | 1.69 | 0.38 | 2.06 | 1.67 | — | 0.27 | 0.06 | 1.41 | 1.35 | 0.06 | 0.15 | — | 1.75 | 2.03 | 0.14 |

## Run Issues

- alignment / openpyxl-readonly: Write unsupported
- alignment / polars: Write unsupported
- alignment / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'
- alignment / python-calamine: Write unsupported
- alignment / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- alignment / xlsxwriter-constmem: Read unsupported
- alignment / xlsxwriter: Read unsupported
- alignment / xlwt: Read unsupported
- background_colors / openpyxl-readonly: Write unsupported
- background_colors / polars: Write unsupported
- background_colors / python-calamine: Write unsupported
- background_colors / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- background_colors / xlsxwriter-constmem: Read unsupported
- background_colors / xlsxwriter: Read unsupported
- background_colors / xlwt: Read unsupported
- borders / openpyxl-readonly: Write unsupported
- borders / polars: Write unsupported
- borders / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'
- borders / python-calamine: Write unsupported
- borders / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- borders / xlsxwriter-constmem: Read unsupported
- borders / xlsxwriter: Read unsupported
- borders / xlwt: Read unsupported
- cell_values / openpyxl-readonly: Write unsupported
- cell_values / polars: Write unsupported
- cell_values / python-calamine: Write unsupported
- cell_values / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- cell_values / xlsxwriter-constmem: Read unsupported
- cell_values / xlsxwriter: Read unsupported
- cell_values / xlwt: Read unsupported
- comments / openpyxl-readonly: Write unsupported
- comments / polars: Write unsupported
- comments / python-calamine: Write unsupported
- comments / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- comments / xlsxwriter-constmem: Read unsupported
- comments / xlsxwriter: Read unsupported
- comments / xlwt: Read unsupported
- conditional_formatting / openpyxl-readonly: Write unsupported
- conditional_formatting / polars: Write unsupported
- conditional_formatting / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'
- conditional_formatting / python-calamine: Write unsupported
- conditional_formatting / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- conditional_formatting / xlsxwriter-constmem: Read unsupported
- conditional_formatting / xlsxwriter: Read unsupported
- conditional_formatting / xlwt: Read unsupported
- data_validation / openpyxl-readonly: Write unsupported
- data_validation / polars: Write unsupported
- data_validation / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'
- data_validation / python-calamine: Write unsupported
- data_validation / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- data_validation / xlsxwriter-constmem: Read unsupported
- data_validation / xlsxwriter: Read unsupported
- data_validation / xlwt: Read unsupported
- dimensions / openpyxl-readonly: Write unsupported
- dimensions / polars: Write unsupported
- dimensions / python-calamine: Write unsupported
- dimensions / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- dimensions / xlsxwriter-constmem: Read unsupported
- dimensions / xlsxwriter: Read unsupported
- dimensions / xlwt: Read unsupported
- formulas / openpyxl-readonly: Write unsupported
- formulas / polars: Write unsupported
- formulas / python-calamine: Write unsupported
- formulas / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- formulas / xlsxwriter-constmem: Read unsupported
- formulas / xlsxwriter: Read unsupported
- formulas / xlwt: Read unsupported
- freeze_panes / openpyxl-readonly: Write unsupported
- freeze_panes / polars: Write unsupported
- freeze_panes / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'
- freeze_panes / python-calamine: Write unsupported
- freeze_panes / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- freeze_panes / xlsxwriter-constmem: Read unsupported
- freeze_panes / xlsxwriter: Read unsupported
- freeze_panes / xlwt: Read unsupported
- hyperlinks / openpyxl-readonly: Write unsupported
- hyperlinks / polars: Write unsupported
- hyperlinks / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'
- hyperlinks / python-calamine: Write unsupported
- hyperlinks / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- hyperlinks / xlsxwriter-constmem: Read unsupported
- hyperlinks / xlsxwriter: Read unsupported
- hyperlinks / xlwt: Read unsupported
- images / openpyxl-readonly: Write unsupported
- images / polars: Write unsupported
- images / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'
- images / python-calamine: Write unsupported
- images / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- images / xlsxwriter-constmem: Read unsupported
- images / xlsxwriter: Read unsupported
- images / xlwt: Read unsupported
- merged_cells / openpyxl-readonly: Write unsupported
- merged_cells / polars: Write unsupported
- merged_cells / python-calamine: Write unsupported
- merged_cells / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- merged_cells / xlsxwriter-constmem: Read unsupported
- merged_cells / xlsxwriter: Read unsupported
- merged_cells / xlwt: Read unsupported
- multiple_sheets / openpyxl-readonly: Write unsupported
- multiple_sheets / polars: Write unsupported
- multiple_sheets / python-calamine: Write unsupported
- multiple_sheets / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- multiple_sheets / xlsxwriter-constmem: Read unsupported
- multiple_sheets / xlsxwriter: Read unsupported
- multiple_sheets / xlwt: Read unsupported
- named_ranges / openpyxl-readonly: Write unsupported
- named_ranges / polars: Write unsupported
- named_ranges / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'
- named_ranges / python-calamine: Write unsupported
- named_ranges / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- named_ranges / xlsxwriter-constmem: Read unsupported
- named_ranges / xlsxwriter: Read unsupported
- named_ranges / xlwt: Read unsupported
- number_formats / openpyxl-readonly: Write unsupported
- number_formats / polars: Write unsupported
- number_formats / python-calamine: Write unsupported
- number_formats / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- number_formats / xlsxwriter-constmem: Read unsupported
- number_formats / xlsxwriter: Read unsupported
- number_formats / xlwt: Read unsupported
- pivot_tables / openpyxl-readonly: Write unsupported
- pivot_tables / polars: Write unsupported
- pivot_tables / python-calamine: Write unsupported
- pivot_tables / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- pivot_tables / xlsxwriter-constmem: Read unsupported
- pivot_tables / xlsxwriter: Read unsupported
- pivot_tables / xlwt: Read unsupported
- tables / openpyxl-readonly: Write unsupported
- tables / polars: Write unsupported
- tables / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'
- tables / python-calamine: Write unsupported
- tables / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- tables / xlsxwriter-constmem: Read unsupported
- tables / xlsxwriter: Read unsupported
- tables / xlwt: Read unsupported
- text_formatting / openpyxl-readonly: Write unsupported
- text_formatting / polars: Write unsupported
- text_formatting / python-calamine: Write unsupported
- text_formatting / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- text_formatting / xlsxwriter-constmem: Read unsupported
- text_formatting / xlsxwriter: Read unsupported
- text_formatting / xlwt: Read unsupported
