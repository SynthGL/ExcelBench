# ExcelBench Dashboard

*Profile: xlsx | Generated: 2026-02-14T00:04:45.572737+00:00*

> Combined fidelity and performance view. Fidelity shows correctness;
> throughput shows speed. Use this to find the right library for your needs.

## Library Comparison

| Library | Caps | Green Features | Pass Rate | Read (cells/s) | Write (cells/s) | Best For |
|---------|:----:|:--------------:|:---------:|:--------------:|:---------------:|----------|
| openpyxl | R+W | 18/18 | 100% | 337K | 354K | Full-fidelity read + write |
| xlsxwriter | W | 16/18 | 90% | — | 533K | High-fidelity write-only workflows |
| pyumya | R+W | 14/18 | 77% | — | — | General use |
| umya-spreadsheet | R+W | 14/18 | 90% | — | — | General use |
| xlsxwriter-constmem | W | 13/18 | 85% | — | 4.7M | Large file writes with memory limits |
| rust_xlsxwriter | W | 8/18 | 68% | — | — | General use |
| xlwt | W | 4/18 | 58% | — | 486K | Legacy .xls file writes |
| openpyxl-readonly | R | 3/18 | 22% | 381K | — | Streaming reads when formatting isn't needed |
| pandas | R+W | 3/18 | 19% | 387K | 250K | Data analysis pipelines (accept NaN coercion) |
| pyexcel | R+W | 3/18 | 20% | 62K | 306K | Multi-format compatibility layer |
| tablib | R+W | 3/18 | 20% | 443K | 274K | Dataset export/import workflows |
| pylightxl | R+W | 2/18 | 18% | — | 311K | Lightweight value extraction |
| calamine | R | 1/18 | 18% | — | — | General use |
| python-calamine | R | 1/18 | 16% | 1.6M | — | Fast bulk value reads |
| polars | R | 0/18 | 14% | 1.3M | — | High-performance DataFrames (values only) |

## Key Insights

- **Fidelity leaders**: openpyxl (18/18 green features)
- **Fastest reader**: python-calamine (1.6M cells/s on cell_values)
- **Fastest writer**: xlsxwriter-constmem (4.7M cells/s on cell_values)
- **Abstraction cost**: pandas wraps openpyxl but drops from 18 to 3 green features due to DataFrame coercion
- **Optimization cost**: xlsxwriter constant_memory mode loses 3 green features for lower memory usage
