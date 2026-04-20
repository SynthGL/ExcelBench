# ExcelBench Performance Results

*Generated: 2026-04-20T08:46:38.157240+00:00*
*Profile: xlsx*
*Platform: Darwin-arm64*
*Python: 3.12.3*
*Commit: 0aeaf16*
*Config: warmup=3 iters=25 iteration_policy=fixed breakdown=False*

## Notes

These numbers measure only the library under test. Write timings do NOT include oracle verification.

## Summary (p50 wall time)

**Tier 0 — Basic Values**

| Feature | openpyxl (R p50 ms) | openpyxl (W p50 ms) | openpyxl-readonly (R p50 ms) | pandas (R p50 ms) | pandas (W p50 ms) | polars (R p50 ms) | pyexcel (R p50 ms) | pyexcel (W p50 ms) | pylightxl (R p50 ms) | pylightxl (W p50 ms) | python-calamine (R p50 ms) | rust_xlsxwriter (W p50 ms) | tablib (R p50 ms) | tablib (W p50 ms) | wolfxl (R p50 ms) | wolfxl (W p50 ms) | xlrd (R p50 ms) | xlsxwriter (W p50 ms) | xlsxwriter-constmem (W p50 ms) | xlwt (W p50 ms) |
|---------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|
| cell_values | 1.39 | 1.91 | 5.83 | 2.27 | 2.24 | 0.42 | 1.62 | 2.13 | 1.49 | 0.51 | 0.97 | 0.54 | 1.50 | 1.74 | 0.25 | 0.53 | — | 3.16 | 3.01 | 0.29 |
| formulas | 1.38 | 1.83 | 1.50 | 1.89 | 2.77 | 0.48 | 1.67 | 2.15 | 1.24 | 0.31 | 0.17 | 0.56 | 1.26 | 1.70 | 0.19 | 0.47 | — | 2.19 | 2.60 | 0.37 |
| multiple_sheets | 1.38 | 2.08 | 1.35 | 2.10 | 4.00 | 0.80 | 1.71 | 2.65 | 1.55 | 0.46 | 0.11 | 0.65 | 1.56 | 3.15 | 0.08 | 0.59 | — | 3.04 | 4.38 | 0.27 |

**Tier 1 — Formatting**

| Feature | openpyxl (R p50 ms) | openpyxl (W p50 ms) | openpyxl-readonly (R p50 ms) | pandas (R p50 ms) | pandas (W p50 ms) | polars (R p50 ms) | pyexcel (R p50 ms) | pyexcel (W p50 ms) | pylightxl (R p50 ms) | pylightxl (W p50 ms) | python-calamine (R p50 ms) | rust_xlsxwriter (W p50 ms) | tablib (R p50 ms) | tablib (W p50 ms) | wolfxl (R p50 ms) | wolfxl (W p50 ms) | xlrd (R p50 ms) | xlsxwriter (W p50 ms) | xlsxwriter-constmem (W p50 ms) | xlwt (W p50 ms) |
|---------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|
| alignment | 1.69 | 2.85 | 1.07 | 2.45 | 1.97 | 0.34 | 2.72 | 1.83 | — | 0.47 | 0.08 | 0.56 | 1.24 | 2.08 | 0.29 | 0.50 | — | 2.86 | 4.70 | 0.30 |
| background_colors | 1.18 | 1.52 | 0.96 | 1.44 | 1.88 | 0.31 | 1.19 | 1.48 | 1.04 | 0.31 | 0.06 | 0.40 | 1.22 | 1.47 | 0.27 | 0.39 | — | 2.46 | 2.40 | 0.24 |
| borders | 2.65 | 4.39 | 1.49 | 2.39 | 2.07 | 0.35 | 2.01 | 1.65 | — | 0.49 | 0.07 | 0.64 | 1.80 | 1.90 | 0.36 | 0.62 | — | 3.59 | 3.28 | 0.50 |
| dimensions | 1.07 | 1.59 | 0.91 | 1.36 | 1.57 | 0.34 | 1.14 | 1.59 | 1.00 | 0.28 | 0.09 | 0.39 | 1.05 | 1.56 | 0.09 | 0.33 | — | 2.23 | 2.35 | 0.18 |
| number_formats | 1.23 | 1.67 | 1.05 | 1.37 | 2.49 | 0.32 | 1.28 | 1.50 | 1.07 | 0.33 | 0.07 | 0.47 | 1.15 | 1.68 | 0.43 | 0.43 | — | 2.24 | 2.89 | 0.25 |
| text_formatting | 2.22 | 2.59 | 1.57 | 2.29 | 1.94 | 0.34 | 2.24 | 1.78 | 1.76 | 0.56 | 0.08 | 0.60 | 1.80 | 1.65 | 0.39 | 0.58 | — | 3.04 | 2.76 | 0.45 |

**Tier 2 — Advanced**

| Feature | openpyxl (R p50 ms) | openpyxl (W p50 ms) | openpyxl-readonly (R p50 ms) | pandas (R p50 ms) | pandas (W p50 ms) | polars (R p50 ms) | pyexcel (R p50 ms) | pyexcel (W p50 ms) | pylightxl (R p50 ms) | pylightxl (W p50 ms) | python-calamine (R p50 ms) | rust_xlsxwriter (W p50 ms) | tablib (R p50 ms) | tablib (W p50 ms) | wolfxl (R p50 ms) | wolfxl (W p50 ms) | xlrd (R p50 ms) | xlsxwriter (W p50 ms) | xlsxwriter-constmem (W p50 ms) | xlwt (W p50 ms) |
|---------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|
| comments | 1.26 | 1.97 | 0.94 | 1.67 | 1.65 | 0.28 | 1.38 | 1.65 | 1.32 | 0.43 | 0.07 | 0.53 | 1.02 | 1.59 | 0.21 | 0.55 | — | 3.10 | 2.22 | 0.17 |
| conditional_formatting | 1.94 | 2.72 | 1.25 | 2.75 | 2.49 | 0.69 | 1.71 | 2.40 | — | 0.31 | 0.06 | 0.52 | 1.69 | 2.00 | 0.35 | 0.56 | — | 3.16 | 2.73 | 0.19 |
| data_validation | 1.51 | 1.69 | 0.95 | 1.86 | 1.53 | 0.66 | 1.53 | 1.47 | — | 0.25 | 0.06 | 0.38 | 1.63 | 1.48 | 0.16 | 0.40 | — | 2.04 | 2.27 | 0.16 |
| freeze_panes | 1.52 | 2.75 | 1.14 | 2.05 | 3.53 | 0.64 | 1.67 | 3.15 | — | 0.52 | 0.06 | 1.36 | 1.71 | 3.82 | 0.24 | 1.38 | — | 3.30 | 4.41 | 0.27 |
| hyperlinks | 1.51 | 1.62 | 1.04 | 1.90 | 2.01 | 0.51 | 1.50 | 1.83 | — | 0.27 | 0.06 | 0.52 | 1.31 | 1.66 | 0.24 | 0.49 | — | 3.38 | 3.04 | 0.16 |
| images | 1.34 | 2.26 | 0.84 | 1.16 | 1.56 | 0.32 | 1.34 | 1.57 | — | 0.25 | 0.06 | 0.37 | 1.13 | 1.51 | 0.06 | 0.35 | — | 3.74 | 2.27 | 0.16 |
| merged_cells | 1.61 | 2.24 | 1.13 | 2.66 | 2.23 | 0.36 | 1.56 | 2.71 | 1.12 | 0.39 | 0.07 | 0.45 | 1.43 | 2.11 | 0.32 | 0.45 | — | 3.68 | 2.80 | 0.28 |
| named_ranges | 1.32 | 2.35 | 0.90 | 1.97 | 2.37 | 0.45 | 1.64 | 2.09 | — | 0.56 | 0.06 | 0.64 | 1.38 | 1.79 | 0.07 | 1.35 | — | 3.99 | 2.79 | 0.19 |
| pivot_tables | 0.92 | 1.44 | 0.84 | 1.22 | 1.70 | 0.33 | 1.02 | 1.41 | 0.90 | 0.27 | 0.06 | 0.47 | 0.93 | 1.50 | 0.06 | 0.42 | — | 2.41 | 2.37 | 0.17 |
| tables | 1.65 | 1.63 | 0.99 | 2.01 | 1.73 | 0.42 | 2.05 | 1.62 | — | 0.34 | 0.06 | 0.38 | 1.36 | 1.59 | 0.06 | 0.36 | — | 2.38 | 2.50 | 0.19 |

## Run Issues

- alignment / openpyxl-readonly: Write unsupported
- alignment / polars: Write unsupported
- alignment / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'
- alignment / python-calamine: Write unsupported
- alignment / rust_xlsxwriter: Read unsupported
- alignment / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- alignment / xlsxwriter-constmem: Read unsupported
- alignment / xlsxwriter: Read unsupported
- alignment / xlwt: Read unsupported
- background_colors / openpyxl-readonly: Write unsupported
- background_colors / polars: Write unsupported
- background_colors / python-calamine: Write unsupported
- background_colors / rust_xlsxwriter: Read unsupported
- background_colors / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- background_colors / xlsxwriter-constmem: Read unsupported
- background_colors / xlsxwriter: Read unsupported
- background_colors / xlwt: Read unsupported
- borders / openpyxl-readonly: Write unsupported
- borders / polars: Write unsupported
- borders / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'
- borders / python-calamine: Write unsupported
- borders / rust_xlsxwriter: Read unsupported
- borders / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- borders / xlsxwriter-constmem: Read unsupported
- borders / xlsxwriter: Read unsupported
- borders / xlwt: Read unsupported
- cell_values / openpyxl-readonly: Write unsupported
- cell_values / polars: Write unsupported
- cell_values / python-calamine: Write unsupported
- cell_values / rust_xlsxwriter: Read unsupported
- cell_values / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- cell_values / xlsxwriter-constmem: Read unsupported
- cell_values / xlsxwriter: Read unsupported
- cell_values / xlwt: Read unsupported
- comments / openpyxl-readonly: Write unsupported
- comments / polars: Write unsupported
- comments / python-calamine: Write unsupported
- comments / rust_xlsxwriter: Read unsupported
- comments / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- comments / xlsxwriter-constmem: Read unsupported
- comments / xlsxwriter: Read unsupported
- comments / xlwt: Read unsupported
- conditional_formatting / openpyxl-readonly: Write unsupported
- conditional_formatting / polars: Write unsupported
- conditional_formatting / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'
- conditional_formatting / python-calamine: Write unsupported
- conditional_formatting / rust_xlsxwriter: Read unsupported
- conditional_formatting / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- conditional_formatting / xlsxwriter-constmem: Read unsupported
- conditional_formatting / xlsxwriter: Read unsupported
- conditional_formatting / xlwt: Read unsupported
- data_validation / openpyxl-readonly: Write unsupported
- data_validation / polars: Write unsupported
- data_validation / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'
- data_validation / python-calamine: Write unsupported
- data_validation / rust_xlsxwriter: Read unsupported
- data_validation / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- data_validation / xlsxwriter-constmem: Read unsupported
- data_validation / xlsxwriter: Read unsupported
- data_validation / xlwt: Read unsupported
- dimensions / openpyxl-readonly: Write unsupported
- dimensions / polars: Write unsupported
- dimensions / python-calamine: Write unsupported
- dimensions / rust_xlsxwriter: Read unsupported
- dimensions / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- dimensions / xlsxwriter-constmem: Read unsupported
- dimensions / xlsxwriter: Read unsupported
- dimensions / xlwt: Read unsupported
- formulas / openpyxl-readonly: Write unsupported
- formulas / polars: Write unsupported
- formulas / python-calamine: Write unsupported
- formulas / rust_xlsxwriter: Read unsupported
- formulas / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- formulas / xlsxwriter-constmem: Read unsupported
- formulas / xlsxwriter: Read unsupported
- formulas / xlwt: Read unsupported
- freeze_panes / openpyxl-readonly: Write unsupported
- freeze_panes / polars: Write unsupported
- freeze_panes / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'
- freeze_panes / python-calamine: Write unsupported
- freeze_panes / rust_xlsxwriter: Read unsupported
- freeze_panes / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- freeze_panes / xlsxwriter-constmem: Read unsupported
- freeze_panes / xlsxwriter: Read unsupported
- freeze_panes / xlwt: Read unsupported
- hyperlinks / openpyxl-readonly: Write unsupported
- hyperlinks / polars: Write unsupported
- hyperlinks / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'
- hyperlinks / python-calamine: Write unsupported
- hyperlinks / rust_xlsxwriter: Read unsupported
- hyperlinks / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- hyperlinks / xlsxwriter-constmem: Read unsupported
- hyperlinks / xlsxwriter: Read unsupported
- hyperlinks / xlwt: Read unsupported
- images / openpyxl-readonly: Write unsupported
- images / polars: Write unsupported
- images / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'
- images / python-calamine: Write unsupported
- images / rust_xlsxwriter: Read unsupported
- images / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- images / xlsxwriter-constmem: Read unsupported
- images / xlsxwriter: Read unsupported
- images / xlwt: Read unsupported
- merged_cells / openpyxl-readonly: Write unsupported
- merged_cells / polars: Write unsupported
- merged_cells / python-calamine: Write unsupported
- merged_cells / rust_xlsxwriter: Read unsupported
- merged_cells / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- merged_cells / xlsxwriter-constmem: Read unsupported
- merged_cells / xlsxwriter: Read unsupported
- merged_cells / xlwt: Read unsupported
- multiple_sheets / openpyxl-readonly: Write unsupported
- multiple_sheets / polars: Write unsupported
- multiple_sheets / python-calamine: Write unsupported
- multiple_sheets / rust_xlsxwriter: Read unsupported
- multiple_sheets / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- multiple_sheets / xlsxwriter-constmem: Read unsupported
- multiple_sheets / xlsxwriter: Read unsupported
- multiple_sheets / xlwt: Read unsupported
- named_ranges / openpyxl-readonly: Write unsupported
- named_ranges / polars: Write unsupported
- named_ranges / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'
- named_ranges / python-calamine: Write unsupported
- named_ranges / rust_xlsxwriter: Read unsupported
- named_ranges / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- named_ranges / xlsxwriter-constmem: Read unsupported
- named_ranges / xlsxwriter: Read unsupported
- named_ranges / xlwt: Read unsupported
- number_formats / openpyxl-readonly: Write unsupported
- number_formats / polars: Write unsupported
- number_formats / python-calamine: Write unsupported
- number_formats / rust_xlsxwriter: Read unsupported
- number_formats / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- number_formats / xlsxwriter-constmem: Read unsupported
- number_formats / xlsxwriter: Read unsupported
- number_formats / xlwt: Read unsupported
- pivot_tables / openpyxl-readonly: Write unsupported
- pivot_tables / polars: Write unsupported
- pivot_tables / python-calamine: Write unsupported
- pivot_tables / rust_xlsxwriter: Read unsupported
- pivot_tables / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- pivot_tables / xlsxwriter-constmem: Read unsupported
- pivot_tables / xlsxwriter: Read unsupported
- pivot_tables / xlwt: Read unsupported
- tables / openpyxl-readonly: Write unsupported
- tables / polars: Write unsupported
- tables / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'
- tables / python-calamine: Write unsupported
- tables / rust_xlsxwriter: Read unsupported
- tables / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- tables / xlsxwriter-constmem: Read unsupported
- tables / xlsxwriter: Read unsupported
- tables / xlwt: Read unsupported
- text_formatting / openpyxl-readonly: Write unsupported
- text_formatting / polars: Write unsupported
- text_formatting / python-calamine: Write unsupported
- text_formatting / rust_xlsxwriter: Read unsupported
- text_formatting / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- text_formatting / xlsxwriter-constmem: Read unsupported
- text_formatting / xlsxwriter: Read unsupported
- text_formatting / xlwt: Read unsupported
