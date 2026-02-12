# ExcelBench Dashboard

*Profile: xlsx | Generated: 2026-02-08T22:56:26.041377+00:00*

> Combined fidelity and performance view. Fidelity shows correctness;
> throughput shows speed. Use this to find the right library for your needs.

## Library Comparison

| Library | Caps | Green Features | Pass Rate | Read (cells/s) | Write (cells/s) | Best For |
|---------|:----:|:--------------:|:---------:|:--------------:|:---------------:|----------|
| openpyxl | R+W | 16/16 | 100% | 337K | 354K | Full-fidelity read + write |
| xlsxwriter | W | 16/16 | 100% | — | 533K | High-fidelity write-only workflows |
| xlsxwriter-constmem | W | 13/16 | 94% | — | 4.7M | Large file writes with memory limits |
| xlwt | W | 4/16 | 64% | — | 486K | Legacy .xls file writes |
| openpyxl-readonly | R | 3/16 | 24% | 381K | — | Streaming reads when formatting isn't needed |
| pandas | R+W | 3/16 | 21% | 387K | 250K | Data analysis pipelines (accept NaN coercion) |
| pyexcel | R+W | 3/16 | 23% | 62K | 306K | Multi-format compatibility layer |
| tablib | R+W | 3/16 | 23% | 443K | 274K | Dataset export/import workflows |
| pylightxl | R+W | 2/16 | 20% | — | 311K | Lightweight value extraction |
| python-calamine | R | 1/16 | 18% | 1.6M | — | Fast bulk value reads |
| polars | R | 0/16 | 16% | 1.3M | — | High-performance DataFrames (values only) |

## Key Insights

- **Fidelity leaders**: openpyxl, xlsxwriter (16/16 green features)
- **Fastest reader**: python-calamine (1.6M cells/s on cell_values)
- **Fastest writer**: xlsxwriter-constmem (4.7M cells/s on cell_values)
- **Abstraction cost**: pandas wraps openpyxl but drops from 16 to 3 green features due to DataFrame coercion
- **Optimization cost**: xlsxwriter constant_memory mode loses 3 green features for lower memory usage
