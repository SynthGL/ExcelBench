# ExcelBench Dashboard

*Profile: xlsx | Generated: 2026-04-29T03:12:20.382948+00:00*

> Combined fidelity and performance view. Fidelity shows correctness;
> throughput shows speed. Use this to find the right library for your needs.

## Library Comparison

| Library | Caps | Modify | Green Features | Pass Rate | Best For |
|---------|:----:|:------:|:--------------:|:---------:|----------|
| openpyxl | R+W | Rewrite | 18/18 | 100% | Full-fidelity read + write |
| wolfxl | R+W | Patch | 18/18 | 100% | General use |
| xlsxwriter | W | No | 15/18 | 90% | High-fidelity write-only workflows |
| xlsxwriter-constmem | W | No | 12/18 | 84% | Large file writes with memory limits |
| xlwt | W | No | 4/18 | 58% | Legacy .xls file writes |
| openpyxl-readonly | R | No | 3/18 | 21% | Streaming reads when formatting isn't needed |
| pandas | R+W | Rebuild | 3/18 | 18% | Data analysis pipelines (accept NaN coercion) |
| pyexcel | R+W | Rebuild | 3/18 | 19% | Multi-format compatibility layer |
| tablib | R+W | Rebuild | 3/18 | 19% | Dataset export/import workflows |
| pylightxl | R+W | Rebuild | 2/18 | 18% | Lightweight value extraction |
| python-calamine | R | No | 1/18 | 15% | Fast bulk value reads |
| polars | R | No | 0/18 | 14% | High-performance DataFrames (values only) |

## Key Insights

- **Fidelity leaders**: openpyxl, wolfxl (18/18 green features)
- **Abstraction cost**: pandas wraps openpyxl but drops from 18 to 3 green features due to DataFrame coercion
- **Optimization cost**: xlsxwriter constant_memory mode loses 3 green features for lower memory usage

## Best Adapter by Workload Profile

| Workload Size | Best Read Adapter | Best Write Adapter |
|---------------|-------------------|--------------------|
| small | — | — |
| medium | — | — |
| large | — | — |
