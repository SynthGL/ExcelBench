# ExcelBench Performance Results

*Generated: 2026-10-28T01:54:17.715991+00:00*
*Profile: xlsx*
*Platform: Darwin-arm64*
*Python: 3.13.9*
*CPU: arm*
*Cores: 18 | Memory MB: 65536.0*
*Commit: 5f71713*
*Config: warmup=3 iters=25 iteration_policy=fixed breakdown=False*

## Run receipt

Snapshot label: 2026-10-27 local date (America/Los_Angeles); measured 2026-10-28T01:54:17Z to 01:54:31Z. Native arm64 CPython 3.13.9 on Apple M5 Pro, 18 cores, 64 GiB RAM, macOS 26.5.1 (25F80). Source commit: `5f71713413c43f3733e303ba8a4eed6a756760cd`. WolfXL Community 2.0.21 uses its public API and the verified public PyPI wheel. See the [full receipt](../README.md), [wheel identity](../receipt.json), [package pins](../requirements.txt), [command](../run.sh), and [raw output](results.json).

All 13 adapters and the 19 features match the previous perf command, including unsuccessful operations reported below. Three warmups and 25 recorded iterations per library, feature, and operation; fixed policy, breakdown off, memory mode `getrusage`. These are warm in-process timings, not cold start. No per-operation timeout; the complete calc-plus-perf driver had a 600-second outer Python timeout. This perf invocation completed in 14.63 seconds, exit 0. One scored attempt, none discarded.

This is a local shared-machine trial, not an isolated version comparison: the 2026-10-03 snapshot used Apple M4 Pro and Python 3.12.3. Do not attribute cross-date changes to WolfXL alone. The new raw results include all failed or unsupported operations. The 1-minute load was 9.05 before perf and 7.94 after; no load gate or exclusive-machine guarantee was applied.

### Per-adapter totals

Totals are sums of per-feature p50 wall milliseconds, derived from `results.json`; `n` counts successfully measured features. Ratios are openpyxl's 19-feature total divided by each other 19-feature total within this run. Partial feature sets are not comparable and receive no ratio.

| Adapter | Read total ms | Read n | Read vs openpyxl | Write total ms | Write n | Write vs openpyxl |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| openpyxl | 22.13 | 19 | 1.00x | 28.09 | 19 | 1.00x |
| wolfxl | 14.54 | 19 | 1.52x | 13.73 | 19 | 2.05x |
| python-calamine | 1.75 | 19 | 12.67x | -- | 0 | -- |
| polars | 7.24 | 19 | 3.06x | -- | 0 | -- |
| pandas | 31.28 | 19 | 0.71x | 32.80 | 19 | 0.86x |
| xlsxwriter | -- | 0 | -- | 36.28 | 19 | 0.77x |
| xlsxwriter-constmem | -- | 0 | -- | 35.29 | 19 | 0.80x |
| openpyxl-readonly | 21.58 | 19 | 1.03x | -- | 0 | -- |
| tablib | 22.28 | 19 | 0.99x | 25.52 | 19 | 1.10x |
| pyexcel | 25.58 | 19 | 0.87x | 7.84 | 5 | -- |
| pylightxl | 9.67 | 10 | -- | 1.26 | 5 | -- |
| xlwt | -- | 0 | -- | 2.88 | 14 | -- |
| xlrd | -- | 0 | -- | -- | 0 | -- |

These **19 small feature workbooks mostly measure per-call overhead**, not large-file throughput. **python-calamine leads this small-file read comparison.** For a different workload, see WolfXL's [large-file ecosystem results](https://github.com/SynthGL/wolfxl-oss/tree/main/benchmarks/results), including the [1.6M-cell comparison](https://github.com/SynthGL/wolfxl-oss/blob/main/benchmarks/results/2026-08-18-ecosystem-full-linux-epyc.md); its APIs, hosts, versions, and ratios are not interchangeable with this lane.


## Notes

These numbers measure only the library under test. Write timings do NOT include oracle verification.
Confidence note: treat deltas under ~5% as noise unless stable across multiple runs.

## Summary (p50 wall time)

**Tier 0 -- Basic Values**

| Feature | openpyxl (R p50/p95 ms) | openpyxl (W p50/p95 ms) | openpyxl-readonly (R p50/p95 ms) | pandas (R p50/p95 ms) | pandas (W p50/p95 ms) | polars (R p50/p95 ms) | pyexcel (R p50/p95 ms) | pyexcel (W p50/p95 ms) | pylightxl (R p50/p95 ms) | pylightxl (W p50/p95 ms) | python-calamine (R p50/p95 ms) | tablib (R p50/p95 ms) | tablib (W p50/p95 ms) | wolfxl (R p50/p95 ms) | wolfxl (W p50/p95 ms) | xlrd (R p50/p95 ms) | xlsxwriter (W p50/p95 ms) | xlsxwriter-constmem (W p50/p95 ms) | xlwt (W p50/p95 ms) |
|---------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|
| cell_values | 1.22/1.48 | 1.34/1.45 | 5.33/5.82 | 2.18/2.71 | 1.88/2.14 | 0.40/0.52 | 1.44/1.72 | -- | 1.29/1.49 | -- | 0.63/0.67 | 1.42/1.75 | 1.42/1.65 | 0.83/1.06 | 1.59/3.56 | -- | 2.80/3.71 | 1.80/2.57 | 0.23/0.28 |
| formulas | 1.15/1.32 | 1.47/1.69 | 1.38/1.69 | 1.72/2.08 | 1.87/2.14 | 0.42/0.57 | 1.24/1.46 | 1.62/2.92 | 1.06/1.24 | 0.27/0.39 | 0.13/0.13 | 1.17/1.42 | 1.40/3.06 | 0.64/0.82 | 0.69/0.91 | -- | 1.67/2.85 | 1.95/3.33 | 0.24/0.27 |
| multiple_sheets | 1.23/1.42 | 1.73/2.09 | 0.98/2.73 | 2.01/2.40 | 2.52/3.01 | 0.64/0.75 | 1.36/1.51 | 2.09/2.60 | 1.15/1.42 | 0.31/0.34 | 0.07/0.10 | 1.25/1.62 | 1.81/2.18 | 0.49/0.65 | 0.63/0.78 | -- | 2.07/13.70 | 2.20/3.53 | 0.20/0.27 |

**Tier 1 -- Formatting**

| Feature | openpyxl (R p50/p95 ms) | openpyxl (W p50/p95 ms) | openpyxl-readonly (R p50/p95 ms) | pandas (R p50/p95 ms) | pandas (W p50/p95 ms) | polars (R p50/p95 ms) | pyexcel (R p50/p95 ms) | pyexcel (W p50/p95 ms) | pylightxl (R p50/p95 ms) | pylightxl (W p50/p95 ms) | python-calamine (R p50/p95 ms) | tablib (R p50/p95 ms) | tablib (W p50/p95 ms) | wolfxl (R p50/p95 ms) | wolfxl (W p50/p95 ms) | xlrd (R p50/p95 ms) | xlsxwriter (W p50/p95 ms) | xlsxwriter-constmem (W p50/p95 ms) | xlwt (W p50/p95 ms) |
|---------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|
| alignment | 1.11/1.31 | 1.30/1.49 | 0.98/1.28 | 1.46/1.71 | 1.75/2.02 | 0.33/0.45 | 1.27/1.50 | -- | -- | -- | 0.05/0.06 | 1.14/1.40 | 1.27/1.50 | 0.67/0.85 | 0.65/0.83 | -- | 1.89/2.68 | 1.92/3.00 | 0.23/0.28 |
| background_colors | 0.88/1.16 | 1.23/1.35 | 0.80/1.03 | 1.30/1.73 | 1.78/2.38 | 0.28/0.39 | 1.17/1.45 | -- | 0.80/1.03 | -- | 0.06/0.08 | 1.00/1.28 | 1.33/5.32 | 0.71/0.94 | 0.63/0.89 | -- | 1.68/3.75 | 1.89/2.63 | 0.19/0.28 |
| borders | 1.63/1.81 | 2.01/2.18 | 1.27/1.59 | 2.09/2.71 | 1.75/2.02 | 0.29/0.37 | 1.94/2.75 | -- | -- | -- | 0.06/0.07 | 1.75/1.99 | 1.41/1.96 | 1.02/1.27 | 0.85/1.03 | -- | 2.14/3.97 | 2.29/4.39 | 0.40/0.70 |
| dimensions | 0.87/1.00 | 1.17/1.42 | 0.74/0.99 | 1.16/1.48 | 1.34/1.73 | 0.39/0.53 | 1.06/1.25 | -- | 0.74/0.99 | -- | 0.06/0.06 | 0.90/1.13 | 1.17/1.36 | 0.57/0.68 | 0.58/0.74 | -- | 1.43/2.12 | 1.66/3.32 | 0.14/0.15 |
| number_formats | 0.93/1.08 | 1.46/2.76 | 0.79/1.01 | 1.32/1.70 | 1.56/1.89 | 0.26/0.36 | 1.00/1.21 | -- | 0.82/1.02 | -- | 0.06/0.07 | 0.91/1.14 | 1.10/1.28 | 0.69/0.85 | 0.59/0.69 | -- | 1.73/3.53 | 1.74/2.10 | 0.18/0.20 |
| text_formatting | 1.59/2.04 | 1.67/2.04 | 1.32/1.53 | 2.05/2.35 | 1.66/1.90 | 0.32/0.46 | 1.65/1.91 | -- | 1.34/1.49 | -- | 0.08/0.08 | 1.68/1.88 | 1.33/1.44 | 0.89/1.12 | 0.73/0.86 | -- | 1.97/3.41 | 1.93/2.47 | 0.30/0.37 |

**Tier 2 -- Advanced**

| Feature | openpyxl (R p50/p95 ms) | openpyxl (W p50/p95 ms) | openpyxl-readonly (R p50/p95 ms) | pandas (R p50/p95 ms) | pandas (W p50/p95 ms) | polars (R p50/p95 ms) | pyexcel (R p50/p95 ms) | pyexcel (W p50/p95 ms) | pylightxl (R p50/p95 ms) | pylightxl (W p50/p95 ms) | python-calamine (R p50/p95 ms) | tablib (R p50/p95 ms) | tablib (W p50/p95 ms) | wolfxl (R p50/p95 ms) | wolfxl (W p50/p95 ms) | xlrd (R p50/p95 ms) | xlsxwriter (W p50/p95 ms) | xlsxwriter-constmem (W p50/p95 ms) | xlwt (W p50/p95 ms) |
|---------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|
| comments | 1.08/1.27 | 1.60/2.04 | 0.76/0.94 | 1.35/1.94 | 1.52/1.98 | 0.31/0.41 | 1.21/1.50 | -- | 0.96/1.34 | -- | 0.06/0.07 | 0.94/1.17 | 1.17/1.51 | 1.22/1.46 | 0.63/0.97 | -- | 2.05/3.21 | 1.58/2.98 | -- |
| conditional_formatting | 1.26/1.53 | 1.64/2.01 | 0.89/1.10 | 1.83/2.07 | 1.73/2.09 | 0.44/0.52 | 1.46/1.61 | -- | -- | -- | 0.05/0.08 | 1.31/1.55 | 1.49/1.90 | 1.03/1.25 | 0.69/0.88 | -- | 1.96/3.83 | 1.76/3.17 | -- |
| data_validation | 1.12/1.34 | 1.18/1.58 | 0.90/1.38 | 1.71/1.96 | 1.49/1.75 | 0.55/0.72 | 1.30/1.62 | -- | -- | -- | 0.05/0.06 | 1.22/1.48 | 1.20/1.40 | 0.62/0.77 | 0.59/0.77 | -- | 1.51/1.86 | 1.67/2.33 | -- |
| freeze_panes | 1.30/1.59 | 1.90/2.27 | 0.86/1.26 | 1.90/2.40 | 2.43/2.91 | 0.49/0.61 | 1.71/2.11 | -- | -- | -- | 0.06/0.09 | 1.37/1.68 | 2.04/2.40 | 0.72/0.87 | 0.74/0.96 | -- | 2.19/2.46 | 2.45/3.96 | 0.18/0.23 |
| hyperlinks | 1.18/1.45 | 1.47/1.66 | 0.80/1.03 | 1.80/2.35 | 1.55/1.89 | 0.47/0.57 | 1.33/1.78 | -- | -- | -- | 0.05/0.05 | 1.13/1.49 | 1.13/1.40 | 1.20/1.65 | 0.74/0.96 | -- | 2.03/3.47 | 1.84/2.71 | -- |
| images | 1.17/1.39 | 1.83/2.42 | 0.69/0.96 | 1.45/1.74 | 1.49/1.93 | 0.27/0.42 | 1.32/2.17 | -- | -- | -- | 0.05/0.05 | 0.86/1.21 | 1.16/1.36 | 0.88/1.19 | 0.79/1.46 | -- | 2.61/4.29 | 1.77/2.36 | -- |
| merged_cells | 1.25/1.36 | 1.32/1.57 | 0.81/1.07 | 1.35/1.59 | 1.72/2.02 | 0.25/0.37 | 1.20/1.44 | -- | 0.82/1.22 | -- | 0.06/0.06 | 0.96/1.32 | 1.18/1.38 | 0.94/2.79 | 0.84/1.22 | -- | 1.70/2.41 | 1.82/3.10 | 0.18/0.21 |
| named_ranges | 1.02/1.25 | 1.41/1.70 | 0.78/1.08 | 1.68/2.32 | 1.77/2.19 | 0.43/0.52 | 1.22/1.58 | 1.50/1.86 | -- | 0.23/0.29 | 0.06/0.07 | 1.20/1.43 | 1.52/1.73 | 0.46/0.70 | 0.60/0.74 | -- | 1.60/2.13 | 1.85/3.55 | 0.14/0.15 |
| pivot_tables | 0.81/1.03 | 1.14/1.52 | 0.67/0.98 | 1.29/1.61 | 1.59/2.01 | 0.23/0.35 | 0.94/1.24 | 1.41/1.69 | 0.70/1.03 | 0.23/0.31 | 0.06/0.07 | 0.81/1.03 | 1.25/1.71 | 0.47/0.52 | 0.62/0.71 | -- | 1.55/1.87 | 1.62/2.18 | 0.12/0.13 |
| tables | 1.34/1.59 | 1.23/1.48 | 0.83/1.10 | 1.64/1.91 | 1.39/1.73 | 0.46/0.63 | 1.76/2.28 | 1.22/1.42 | -- | 0.21/0.23 | 0.05/0.06 | 1.26/1.54 | 1.15/1.30 | 0.49/0.56 | 0.56/0.92 | -- | 1.70/2.49 | 1.55/2.46 | 0.14/0.19 |

## Run Issues

- alignment / openpyxl-readonly: Write unsupported
- alignment / polars: Write unsupported
- alignment / pyexcel: Write failed: UnsupportedAdapterOperationError: pyexcel does not support write_cell_format: pyexcel exposes value-only cells
- alignment / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'; Write failed: UnsupportedAdapterOperationError: pylightxl does not support write_cell_format: pylightxl does not implement this feature
- alignment / python-calamine: Write unsupported
- alignment / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- alignment / xlsxwriter-constmem: Read unsupported
- alignment / xlsxwriter: Read unsupported
- alignment / xlwt: Read unsupported
- background_colors / openpyxl-readonly: Write unsupported
- background_colors / polars: Write unsupported
- background_colors / pyexcel: Write failed: UnsupportedAdapterOperationError: pyexcel does not support write_cell_format: pyexcel exposes value-only cells
- background_colors / pylightxl: Write failed: UnsupportedAdapterOperationError: pylightxl does not support write_cell_format: pylightxl does not implement this feature
- background_colors / python-calamine: Write unsupported
- background_colors / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- background_colors / xlsxwriter-constmem: Read unsupported
- background_colors / xlsxwriter: Read unsupported
- background_colors / xlwt: Read unsupported
- borders / openpyxl-readonly: Write unsupported
- borders / polars: Write unsupported
- borders / pyexcel: Write failed: UnsupportedAdapterOperationError: pyexcel does not support write_cell_border: pyexcel exposes value-only cells
- borders / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'; Write failed: UnsupportedAdapterOperationError: pylightxl does not support write_cell_border: pylightxl does not implement this feature
- borders / python-calamine: Write unsupported
- borders / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- borders / xlsxwriter-constmem: Read unsupported
- borders / xlsxwriter: Read unsupported
- borders / xlwt: Read unsupported
- cell_values / openpyxl-readonly: Write unsupported
- cell_values / polars: Write unsupported
- cell_values / pyexcel: Write failed: UnsupportedAdapterOperationError: pyexcel does not support write_cell_format: pyexcel exposes value-only cells
- cell_values / pylightxl: Write failed: UnsupportedAdapterOperationError: pylightxl does not support write_cell_format: pylightxl does not implement this feature
- cell_values / python-calamine: Write unsupported
- cell_values / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- cell_values / xlsxwriter-constmem: Read unsupported
- cell_values / xlsxwriter: Read unsupported
- cell_values / xlwt: Read unsupported
- comments / openpyxl-readonly: Write unsupported
- comments / polars: Write unsupported
- comments / pyexcel: Write failed: UnsupportedAdapterOperationError: pyexcel does not support add_comment: pyexcel does not expose this worksheet feature
- comments / pylightxl: Write failed: UnsupportedAdapterOperationError: pylightxl does not support add_comment: pylightxl does not implement this feature
- comments / python-calamine: Write unsupported
- comments / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- comments / xlsxwriter-constmem: Read unsupported
- comments / xlsxwriter: Read unsupported
- comments / xlwt: Read unsupported; Write failed: UnsupportedAdapterOperationError: xlwt does not support add_comment: xlwt cannot author comments
- conditional_formatting / openpyxl-readonly: Write unsupported
- conditional_formatting / polars: Write unsupported
- conditional_formatting / pyexcel: Write failed: UnsupportedAdapterOperationError: pyexcel does not support add_conditional_format: pyexcel does not expose this worksheet feature
- conditional_formatting / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'; Write failed: UnsupportedAdapterOperationError: pylightxl does not support add_conditional_format: pylightxl does not implement this feature
- conditional_formatting / python-calamine: Write unsupported
- conditional_formatting / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- conditional_formatting / xlsxwriter-constmem: Read unsupported
- conditional_formatting / xlsxwriter: Read unsupported
- conditional_formatting / xlwt: Read unsupported; Write failed: UnsupportedAdapterOperationError: xlwt does not support add_conditional_format: xlwt cannot author conditional formatting
- data_validation / openpyxl-readonly: Write unsupported
- data_validation / polars: Write unsupported
- data_validation / pyexcel: Write failed: UnsupportedAdapterOperationError: pyexcel does not support add_data_validation: pyexcel does not expose this worksheet feature
- data_validation / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'; Write failed: UnsupportedAdapterOperationError: pylightxl does not support add_data_validation: pylightxl does not implement this feature
- data_validation / python-calamine: Write unsupported
- data_validation / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- data_validation / xlsxwriter-constmem: Read unsupported
- data_validation / xlsxwriter: Read unsupported
- data_validation / xlwt: Read unsupported; Write failed: UnsupportedAdapterOperationError: xlwt does not support add_data_validation: xlwt cannot author data validations
- dimensions / openpyxl-readonly: Write unsupported
- dimensions / polars: Write unsupported
- dimensions / pyexcel: Write failed: UnsupportedAdapterOperationError: pyexcel does not support set_row_height: pyexcel does not expose row dimension writing
- dimensions / pylightxl: Write failed: UnsupportedAdapterOperationError: pylightxl does not support set_row_height: pylightxl does not implement this feature
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
- freeze_panes / pyexcel: Write failed: UnsupportedAdapterOperationError: pyexcel does not support set_freeze_panes: pyexcel does not expose this worksheet feature
- freeze_panes / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'; Write failed: UnsupportedAdapterOperationError: pylightxl does not support set_freeze_panes: pylightxl does not implement this feature
- freeze_panes / python-calamine: Write unsupported
- freeze_panes / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- freeze_panes / xlsxwriter-constmem: Read unsupported
- freeze_panes / xlsxwriter: Read unsupported
- freeze_panes / xlwt: Read unsupported
- hyperlinks / openpyxl-readonly: Write unsupported
- hyperlinks / polars: Write unsupported
- hyperlinks / pyexcel: Write failed: UnsupportedAdapterOperationError: pyexcel does not support add_hyperlink: pyexcel does not expose this worksheet feature
- hyperlinks / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'; Write failed: UnsupportedAdapterOperationError: pylightxl does not support add_hyperlink: pylightxl does not implement this feature
- hyperlinks / python-calamine: Write unsupported
- hyperlinks / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- hyperlinks / xlsxwriter-constmem: Read unsupported
- hyperlinks / xlsxwriter: Read unsupported
- hyperlinks / xlwt: Read unsupported; Write failed: UnsupportedAdapterOperationError: xlwt does not support add_hyperlink: xlwt hyperlink metadata is not supported in this adapter
- images / openpyxl-readonly: Write unsupported
- images / polars: Write unsupported
- images / pyexcel: Write failed: UnsupportedAdapterOperationError: pyexcel does not support add_image: pyexcel does not expose this worksheet feature
- images / pylightxl: Read failed: TypeError: expected string or bytes-like object, got 'NoneType'; Write failed: UnsupportedAdapterOperationError: pylightxl does not support add_image: pylightxl does not implement this feature
- images / python-calamine: Write unsupported
- images / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- images / xlsxwriter-constmem: Read unsupported
- images / xlsxwriter: Read unsupported
- images / xlwt: Read unsupported; Write failed: UnsupportedAdapterOperationError: xlwt does not support add_image: xlwt cannot embed images in this adapter
- merged_cells / openpyxl-readonly: Write unsupported
- merged_cells / polars: Write unsupported
- merged_cells / pyexcel: Write failed: UnsupportedAdapterOperationError: pyexcel does not support merge_cells: pyexcel does not expose this worksheet feature
- merged_cells / pylightxl: Write failed: UnsupportedAdapterOperationError: pylightxl does not support merge_cells: pylightxl does not implement this feature
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
- number_formats / pyexcel: Write failed: UnsupportedAdapterOperationError: pyexcel does not support write_cell_format: pyexcel exposes value-only cells
- number_formats / pylightxl: Write failed: UnsupportedAdapterOperationError: pylightxl does not support write_cell_format: pylightxl does not implement this feature
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
- text_formatting / pyexcel: Write failed: UnsupportedAdapterOperationError: pyexcel does not support write_cell_format: pyexcel exposes value-only cells
- text_formatting / pylightxl: Write failed: UnsupportedAdapterOperationError: pylightxl does not support write_cell_format: pylightxl does not implement this feature
- text_formatting / python-calamine: Write unsupported
- text_formatting / xlrd: Write unsupported; Read not applicable: xlrd does not support .xlsx input
- text_formatting / xlsxwriter-constmem: Read unsupported
- text_formatting / xlsxwriter: Read unsupported
- text_formatting / xlwt: Read unsupported
