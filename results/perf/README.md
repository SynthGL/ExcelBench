# ExcelBench Performance Results

*Generated: 2026-02-08T23:42:28.484726+00:00*
*Profile: xlsx*
*Platform: Darwin-arm64*
*Python: 3.12.3*
*Commit: 2129abe*
*Config: warmup=1 iters=5 breakdown=True*

## Notes

These numbers measure only the library under test. Write timings do NOT include oracle verification.

## Summary (p50 wall time)

**Tier 2 — Advanced**

| Feature | openpyxl (R p50 ms) | openpyxl (W p50 ms) | openpyxl-readonly (R p50 ms) | pandas (R p50 ms) | pandas (W p50 ms) | polars (R p50 ms) | pyexcel (R p50 ms) | pyexcel (W p50 ms) | pylightxl (R p50 ms) | pylightxl (W p50 ms) | python-calamine (R p50 ms) | tablib (R p50 ms) | tablib (W p50 ms) | xlrd (R p50 ms) | xlsxwriter (W p50 ms) | xlsxwriter-constmem (W p50 ms) | xlwt (W p50 ms) |
|---------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|
| alignment_1k | 11.18 | 17.62 | 1.52 | 7.56 | 8.13 | 1.31 | 7.54 | 5.84 | — | 3.69 | 0.33 | 6.58 | 6.60 | — | 11.82 | 13.37 | 8.49 |
| background_colors_1k | 10.62 | 20.85 | 1.60 | 7.36 | 7.66 | 1.30 | 7.33 | 5.55 | — | 3.71 | 0.28 | 6.87 | 6.64 | — | 12.50 | 13.47 | 9.33 |
| borders_200 | 2.92 | 11.28 | 1.19 | 2.70 | 2.78 | 0.54 | 2.30 | 2.10 | — | 0.95 | 0.10 | 2.34 | 2.28 | — | 5.02 | 4.97 | 2.22 |
| cell_values_10k | 40.27 | 40.70 | 108559.13 | 88.60 | 47.48 | 38.29 | 162.56 | 32.67 | — | 32.14 | 38735.74 | 31.49 | 44.61 | — | 47.17 | 34.09 | 20.58 |
| cell_values_10k_bulk_read | 29.68 | — | 26.21 | 25.81 | — | 7.67 | — | — | — | — | — | 22.56 | — | — | — | — | — |
| cell_values_10k_bulk_write | — | 28.24 | — | — | 39.98 | — | — | — | — | — | — | — | 36.54 | — | 18.76 | 2.13 | — |
| cell_values_1k | 4.84 | 4.92 | 1459.52 | 10.43 | 6.38 | 2.79 | 8.49 | 4.42 | — | 3.48 | 395.17 | 4.41 | 5.76 | — | 4.43 | 4.80 | 2.19 |
| cell_values_1k_bulk_read | 3.68 | — | 3.78 | 4.17 | — | 1.43 | — | — | — | — | — | 3.58 | — | — | — | — | — |
| cell_values_1k_bulk_write | — | 4.04 | — | — | 5.38 | — | — | — | — | — | — | — | 4.72 | — | 3.78 | 2.01 | — |
| formulas_10k | 43.74 | 39.28 | — | 33.44 | 52.14 | — | 47.07 | 32.38 | — | 30.78 | 38940.36 | — | 48.92 | — | 81.69 | 85.79 | 177.36 |
| formulas_1k | 5.36 | 5.60 | 1739.55 | 4.67 | 6.89 | 2.49 | 5.53 | 4.40 | — | 3.18 | 410.70 | 4.61 | 5.77 | — | 9.98 | 9.98 | 15.06 |
| number_formats_1k | 8.23 | 7.71 | 1.76 | 4.63 | 6.56 | 1.49 | 4.79 | 4.54 | — | 3.67 | 0.30 | 4.18 | 6.24 | — | 11.46 | 11.73 | 8.14 |

## Throughput (derived from p50)

Computed as: op_count * 1000 / p50_wall_ms

| Scenario | op_count | op_unit | openpyxl (R units/s) | openpyxl (W units/s) | openpyxl-readonly (R units/s) | pandas (R units/s) | pandas (W units/s) | polars (R units/s) | pyexcel (R units/s) | pyexcel (W units/s) | pylightxl (R units/s) | pylightxl (W units/s) | python-calamine (R units/s) | tablib (R units/s) | tablib (W units/s) | xlrd (R units/s) | xlsxwriter (W units/s) | xlsxwriter-constmem (W units/s) | xlwt (W units/s) |
|----------|----------|---------|----------------|----------------|----------------|----------------|----------------|----------------|----------------|----------------|----------------|----------------|----------------|----------------|----------------|----------------|----------------|----------------|----------------|
| alignment_1k | 1000 | cells | 89.47K | 56.74K | 656.62K | 132.34K | 122.94K | 762.34K | 132.55K | 171.33K | — | 270.98K | 3.08M | 152.01K | 151.53K | — | 84.62K | 74.77K | 117.83K |
| background_colors_1k | 1000 | cells | 94.16K | 47.97K | 626.17K | 135.93K | 130.62K | 769.70K | 136.47K | 180.02K | — | 269.23K | 3.51M | 145.62K | 150.56K | — | 80.00K | 74.22K | 107.20K |
| borders_200 | 200 | cells | 68.38K | 17.72K | 168.78K | 74.21K | 72.02K | 367.70K | 87.11K | 95.08K | — | 210.45K | 2.00M | 85.57K | 87.89K | — | 39.81K | 40.23K | 89.96K |
| cell_values_10k | 10000 | cells | 248.35K | 245.72K | 92.12 | 112.86K | 210.61K | 261.17K | 61.52K | 306.09K | — | 311.14K | 258.16 | 317.58K | 224.15K | — | 212.02K | 293.34K | 485.97K |
| cell_values_10k_bulk_read | 10000 | cells | 336.94K | — | 381.49K | 387.38K | — | 1.30M | — | — | — | — | — | 443.17K | — | — | — | — | — |
| cell_values_10k_bulk_write | 10000 | cells | — | 354.15K | — | — | 250.15K | — | — | — | — | — | — | — | 273.69K | — | 532.95K | 4.69M | — |
| cell_values_1k | 1000 | cells | 206.81K | 203.26K | 685.16 | 95.89K | 156.69K | 358.78K | 117.74K | 226.27K | — | 287.29K | 2.53K | 226.71K | 173.66K | — | 225.88K | 208.34K | 457.18K |
| cell_values_1k_bulk_read | 1000 | cells | 271.82K | — | 264.34K | 239.66K | — | 698.28K | — | — | — | — | — | 278.99K | — | — | — | — | — |
| cell_values_1k_bulk_write | 1000 | cells | — | 247.72K | — | — | 185.96K | — | — | — | — | — | — | — | 211.95K | — | 264.42K | 496.86K | — |
| formulas_10k | 10000 | cells | 228.60K | 254.58K | — | 299.04K | 191.79K | — | 212.45K | 308.81K | — | 324.84K | 256.80 | — | 204.43K | — | 122.42K | 116.56K | 56.38K |
| formulas_1k | 1000 | cells | 186.72K | 178.62K | 574.86 | 214.15K | 145.04K | 401.38K | 180.72K | 227.28K | — | 314.08K | 2.43K | 216.86K | 173.23K | — | 100.25K | 100.23K | 66.40K |
| number_formats_1k | 1000 | cells | 121.45K | 129.78K | 569.57K | 215.89K | 152.52K | 672.27K | 208.87K | 220.25K | — | 272.60K | 3.36M | 239.06K | 160.19K | — | 87.26K | 85.26K | 122.89K |

## Run Issues

- alignment_1k / openpyxl-readonly: Write unsupported
- alignment_1k / polars: Write unsupported
- alignment_1k / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'
- alignment_1k / python-calamine: Write unsupported
- alignment_1k / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- alignment_1k / xlsxwriter-constmem: Read unsupported
- alignment_1k / xlsxwriter: Read unsupported
- alignment_1k / xlwt: Read unsupported
- background_colors_1k / openpyxl-readonly: Write unsupported
- background_colors_1k / polars: Write unsupported
- background_colors_1k / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'
- background_colors_1k / python-calamine: Write unsupported
- background_colors_1k / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- background_colors_1k / xlsxwriter-constmem: Read unsupported
- background_colors_1k / xlsxwriter: Read unsupported
- background_colors_1k / xlwt: Read unsupported
- borders_200 / openpyxl-readonly: Write unsupported
- borders_200 / polars: Write unsupported
- borders_200 / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'
- borders_200 / python-calamine: Write unsupported
- borders_200 / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- borders_200 / xlsxwriter-constmem: Read unsupported
- borders_200 / xlsxwriter: Read unsupported
- borders_200 / xlwt: Read unsupported
- cell_values_10k / openpyxl-readonly: Write unsupported
- cell_values_10k / polars: Write unsupported
- cell_values_10k / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'
- cell_values_10k / python-calamine: Write unsupported
- cell_values_10k / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- cell_values_10k / xlsxwriter-constmem: Read unsupported
- cell_values_10k / xlsxwriter: Read unsupported
- cell_values_10k / xlwt: Read unsupported
- cell_values_10k_bulk_read / pyexcel: Read failed: ValueError: Adapter does not support bulk sheet reads: pyexcel
- cell_values_10k_bulk_read / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'
- cell_values_10k_bulk_read / python-calamine: Read failed: ValueError: Adapter does not support bulk sheet reads: python-calamine
- cell_values_10k_bulk_read / xlrd: Read not applicable: xlrd does not support .xlsx input
- cell_values_10k_bulk_read / xlsxwriter-constmem: Read unsupported
- cell_values_10k_bulk_read / xlsxwriter: Read unsupported
- cell_values_10k_bulk_read / xlwt: Read unsupported
- cell_values_10k_bulk_write / openpyxl-readonly: Write unsupported
- cell_values_10k_bulk_write / polars: Write unsupported
- cell_values_10k_bulk_write / pyexcel: Write failed: ValueError: Adapter does not support bulk sheet writes: pyexcel
- cell_values_10k_bulk_write / pylightxl: Write failed: ValueError: Adapter does not support bulk sheet writes: pylightxl
- cell_values_10k_bulk_write / python-calamine: Write unsupported
- cell_values_10k_bulk_write / xlrd: Write unsupported
- cell_values_10k_bulk_write / xlwt: Write failed: ValueError: Adapter does not support bulk sheet writes: xlwt
- cell_values_1k / openpyxl-readonly: Write unsupported
- cell_values_1k / polars: Write unsupported
- cell_values_1k / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'
- cell_values_1k / python-calamine: Write unsupported
- cell_values_1k / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- cell_values_1k / xlsxwriter-constmem: Read unsupported
- cell_values_1k / xlsxwriter: Read unsupported
- cell_values_1k / xlwt: Read unsupported
- cell_values_1k_bulk_read / pyexcel: Read failed: ValueError: Adapter does not support bulk sheet reads: pyexcel
- cell_values_1k_bulk_read / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'
- cell_values_1k_bulk_read / python-calamine: Read failed: ValueError: Adapter does not support bulk sheet reads: python-calamine
- cell_values_1k_bulk_read / xlrd: Read not applicable: xlrd does not support .xlsx input
- cell_values_1k_bulk_read / xlsxwriter-constmem: Read unsupported
- cell_values_1k_bulk_read / xlsxwriter: Read unsupported
- cell_values_1k_bulk_read / xlwt: Read unsupported
- cell_values_1k_bulk_write / openpyxl-readonly: Write unsupported
- cell_values_1k_bulk_write / polars: Write unsupported
- cell_values_1k_bulk_write / pyexcel: Write failed: ValueError: Adapter does not support bulk sheet writes: pyexcel
- cell_values_1k_bulk_write / pylightxl: Write failed: ValueError: Adapter does not support bulk sheet writes: pylightxl
- cell_values_1k_bulk_write / python-calamine: Write unsupported
- cell_values_1k_bulk_write / xlrd: Write unsupported
- cell_values_1k_bulk_write / xlwt: Write failed: ValueError: Adapter does not support bulk sheet writes: xlwt
- formulas_10k / openpyxl-readonly: Write unsupported; Read failed: EOFError: 
- formulas_10k / polars: Write unsupported; Read failed: CalamineError: calamine error: Xlsx error: Zip error: invalid Zip archive: Could not find EOCD
Context:
    0: Could not open workbook at test_files/throughput_xlsx/tier0/00_formulas_10k.xlsx
    1: could not load excel file at test_files/throughput_xlsx/tier0/00_formulas_10k.xlsx

- formulas_10k / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'
- formulas_10k / python-calamine: Write unsupported
- formulas_10k / tablib: Read failed: BadZipFile: File is not a zip file
- formulas_10k / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- formulas_10k / xlsxwriter-constmem: Read unsupported
- formulas_10k / xlsxwriter: Read unsupported
- formulas_10k / xlwt: Read unsupported
- formulas_1k / openpyxl-readonly: Write unsupported
- formulas_1k / polars: Write unsupported
- formulas_1k / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'
- formulas_1k / python-calamine: Write unsupported
- formulas_1k / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- formulas_1k / xlsxwriter-constmem: Read unsupported
- formulas_1k / xlsxwriter: Read unsupported
- formulas_1k / xlwt: Read unsupported
- number_formats_1k / openpyxl-readonly: Write unsupported
- number_formats_1k / polars: Write unsupported
- number_formats_1k / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'
- number_formats_1k / python-calamine: Write unsupported
- number_formats_1k / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- number_formats_1k / xlsxwriter-constmem: Read unsupported
- number_formats_1k / xlsxwriter: Read unsupported
- number_formats_1k / xlwt: Read unsupported
