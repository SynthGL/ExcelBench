# Performance Benchmark Runs - Log

Created: 02/08/2026 05:07 AM PST (via pst-timestamp)

Use this to record baseline perf runs (machine conditions, command, output paths, quick notes).

## Run Template

```text
### MM/DD/YYYY HH:MM AM/PM PST/PDT (via pst-timestamp)

Git:
- commit: <short sha>
- branch: <name>

Environment:
- OS: <...>
- CPU: <...>
- RAM: <...>
- Python (uv): <...>
- Power: <AC/Battery>
- Notes: <background load, thermal state, etc>

Command:
- `uv run excelbench perf ...`

Outputs:
- JSON: <path>
- Markdown: <path>
- CSV: <path>
- History: <path>

Observations:
- <outliers, regressions, variance>
```

## Runs

### 02/08/2026 05:08 AM PST (via pst-timestamp)

Git:
- commit: c2cab42
- branch: master

Environment:
- OS: macOS 26.2 (25C5031i)
- CPU: Apple M4 Pro
- RAM: 24 GB
- Python (uv): 3.12.3
- Power: (not recorded)
- Notes: Tier 1 only; breakdown enabled.

Command:
- `uv run excelbench perf --tests fixtures/excel --output results_dev_perf_baseline_tier1 --feature cell_values --feature formulas --feature text_formatting --feature background_colors --feature number_formats --feature alignment --feature borders --feature dimensions --feature multiple_sheets --warmup 2 --iters 10 --breakdown`

Outputs:
- JSON: `results_dev_perf_baseline_tier1/perf/results.json`
- Markdown: `results_dev_perf_baseline_tier1/perf/README.md`
- CSV: `results_dev_perf_baseline_tier1/perf/matrix.csv`
- History: `results_dev_perf_baseline_tier1/perf/history.jsonl`

Observations:
- 10 libraries x 9 features = 90 result rows.

### 02/08/2026 01:50 PM PST (via pst-timestamp)

Git:
- commit: c2cab42
- branch: master

Environment:
- OS: macOS 26.2 (25C5031i)
- CPU: Apple M4 Pro
- RAM: 24 GB
- Python (uv): 3.12.3
- Power: (not recorded)
- Notes: Full xlsx feature set; breakdown enabled; warmup=1 iters=5.

Command:
- `uv run excelbench perf --tests fixtures/excel --output results_dev_perf_baseline_full_xlsx --warmup 1 --iters 5 --breakdown`

Outputs:
- JSON: `results_dev_perf_baseline_full_xlsx/perf/results.json`
- Markdown: `results_dev_perf_baseline_full_xlsx/perf/README.md`
- CSV: `results_dev_perf_baseline_full_xlsx/perf/matrix.csv`
- History: `results_dev_perf_baseline_full_xlsx/perf/history.jsonl`

Observations:
- 10 libraries x 17 features = 170 result rows.

### 02/08/2026 02:00 PM PST (via pst-timestamp)

Git:
- commit: c2cab42
- branch: master

Environment:
- OS: macOS 26.2 (25C5031i)
- CPU: Apple M4 Pro
- RAM: 24 GB
- Python (uv): 3.12.3
- Power: (not recorded)
- Notes: xls profile fixtures (4 features); breakdown enabled; warmup=1 iters=5; adapters restricted to xls-relevant set.

Command:
- `uv run excelbench perf --profile xls --tests fixtures/excel_xls --output results_dev_perf_baseline_full_xls --warmup 1 --iters 5 --breakdown --adapter xlrd --adapter python-calamine --adapter xlwt`

Outputs:
- JSON: `results_dev_perf_baseline_full_xls/perf/results.json`
- Markdown: `results_dev_perf_baseline_full_xls/perf/README.md`
- CSV: `results_dev_perf_baseline_full_xls/perf/matrix.csv`
- History: `results_dev_perf_baseline_full_xls/perf/history.jsonl`

Observations:
- 3 libraries x 4 features = 12 result rows.

### 02/08/2026 02:08 PM PST (via pst-timestamp)

Git:
- commit: c2cab42
- branch: master

Environment:
- OS: macOS 26.2 (25C5031i)
- CPU: Apple M4 Pro
- RAM: 24 GB
- Python (uv): 3.12.3
- Power: (not recorded)
- Notes: Throughput fixture (cell_values_10k); adapters restricted to openpyxl; breakdown enabled; warmup=1 iters=3.

Command:
- `uv run excelbench perf --tests test_files/throughput_xlsx --output results_dev_perf_throughput_openpyxl --warmup 1 --iters 3 --breakdown --adapter openpyxl`

Outputs:
- JSON: `results_dev_perf_throughput_openpyxl/perf/results.json`
- Markdown: `results_dev_perf_throughput_openpyxl/perf/README.md`
- CSV: `results_dev_perf_throughput_openpyxl/perf/matrix.csv`
- History: `results_dev_perf_throughput_openpyxl/perf/history.jsonl`

Observations:
- Workload op_count: 10,000 cells (read + write).

### 02/08/2026 02:16 PM PST (via pst-timestamp)

Git:
- commit: c2cab42
- branch: master

Environment:
- OS: macOS 26.2 (25C5031i)
- CPU: Apple M4 Pro
- RAM: 24 GB
- Python (uv): 3.12.3
- Power: (not recorded)
- Notes: Throughput fixtures (cell_values_10k + formulas_10k); adapters restricted to openpyxl; breakdown enabled; warmup=1 iters=2.

Command:
- `uv run excelbench perf --tests test_files/throughput_xlsx --output results_dev_perf_throughput_openpyxl --warmup 1 --iters 2 --breakdown --adapter openpyxl`

Outputs:
- JSON: `results_dev_perf_throughput_openpyxl/perf/results.json`
- Markdown: `results_dev_perf_throughput_openpyxl/perf/README.md`
- CSV: `results_dev_perf_throughput_openpyxl/perf/matrix.csv`
- History: `results_dev_perf_throughput_openpyxl/perf/history.jsonl`

Observations:
- Workload op_count: 10,000 cells (read + write) for both scenarios.

### 02/08/2026 02:18 PM PST (via pst-timestamp)

Git:
- commit: c2cab42
- branch: master

Environment:
- OS: macOS 26.2 (25C5031i)
- CPU: Apple M4 Pro
- RAM: 24 GB
- Python (uv): 3.12.3
- Power: (not recorded)
- Notes: Throughput fixtures (cell_values_10k + formulas_10k); batch run with common adapters; warmup=0 iters=1.

Command:
- `uv run excelbench perf --tests test_files/throughput_xlsx --output results_dev_perf_throughput_batch1 --warmup 0 --iters 1 --adapter openpyxl --adapter xlsxwriter --adapter pylightxl --adapter pyexcel --adapter python-calamine`

Outputs:
- JSON: `results_dev_perf_throughput_batch1/perf/results.json`
- Markdown: `results_dev_perf_throughput_batch1/perf/README.md`
- CSV: `results_dev_perf_throughput_batch1/perf/matrix.csv`
- History: `results_dev_perf_throughput_batch1/perf/history.jsonl`

Observations:
- python-calamine per-cell reads were extremely slow on these 10k scenarios (~39s p50 for read).

### 02/08/2026 02:21 PM PST (via pst-timestamp)

Git:
- commit: c2cab42
- branch: master

Environment:
- OS: macOS 26.2 (25C5031i)
- CPU: Apple M4 Pro
- RAM: 24 GB
- Python (uv): 3.12.3
- Power: (not recorded)
- Notes: Throughput fixtures updated to include `background_colors_1k`; openpyxl only; warmup=0 iters=1.

Command:
- `uv run excelbench perf --tests test_files/throughput_xlsx --output results_dev_perf_throughput_openpyxl --warmup 0 --iters 1 --adapter openpyxl`

Outputs:
- JSON: `results_dev_perf_throughput_openpyxl/perf/results.json`
- Markdown: `results_dev_perf_throughput_openpyxl/perf/README.md`
- CSV: `results_dev_perf_throughput_openpyxl/perf/matrix.csv`
- History: `results_dev_perf_throughput_openpyxl/perf/history.jsonl`

Observations:
- Workload op_count: 10,000 cells for `cell_values_10k` and `formulas_10k`; 1,000 cells for `background_colors_1k`.

### 02/08/2026 02:23 PM PST (via pst-timestamp)

Git:
- commit: c2cab42
- branch: master

Environment:
- OS: macOS 26.2 (25C5031i)
- CPU: Apple M4 Pro
- RAM: 24 GB
- Python (uv): 3.12.3
- Power: (not recorded)
- Notes: Throughput fixtures expanded (adds `number_formats_1k` and `alignment_1k`); openpyxl only; warmup=0 iters=1.

Command:
- `uv run python scripts/generate_throughput_fixtures.py`
- `uv run excelbench perf --tests test_files/throughput_xlsx --output results_dev_perf_throughput_openpyxl --warmup 0 --iters 1 --adapter openpyxl`

Outputs:
- JSON: `results_dev_perf_throughput_openpyxl/perf/results.json`
- Markdown: `results_dev_perf_throughput_openpyxl/perf/README.md`
- CSV: `results_dev_perf_throughput_openpyxl/perf/matrix.csv`
- History: `results_dev_perf_throughput_openpyxl/perf/history.jsonl`

### 02/08/2026 02:36 PM PST (via pst-timestamp)

Git:
- commit: 4704676
- branch: master

Environment:
- OS: macOS 26.2 (25C5031i)
- CPU: Apple M4 Pro
- RAM: 24 GB
- Python (uv): 3.12.3
- Power: (not recorded)
- Notes: Throughput fixtures expanded with `borders_200`; openpyxl only; warmup=0 iters=1.

Command:
- `uv run python scripts/generate_throughput_fixtures.py`
- `uv run excelbench perf --tests test_files/throughput_xlsx --output results_dev_perf_throughput_openpyxl --warmup 0 --iters 1 --adapter openpyxl`

Outputs:
- JSON: `results_dev_perf_throughput_openpyxl/perf/results.json`
- Markdown: `results_dev_perf_throughput_openpyxl/perf/README.md`
- CSV: `results_dev_perf_throughput_openpyxl/perf/matrix.csv`
- History: `results_dev_perf_throughput_openpyxl/perf/history.jsonl`

Observations:
- Workload op_count now includes `borders_200` (200 border ops) in addition to existing scenarios.

### 02/08/2026 02:39 PM PST (via pst-timestamp)

Git:
- commit: 4704676
- branch: master

Environment:
- OS: macOS 26.2 (25C5031i)
- CPU: Apple M4 Pro
- RAM: 24 GB
- Python (uv): 3.12.3
- Power: (not recorded)
- Notes: Throughput fixtures expanded with small variants (`cell_values_1k`, `formulas_1k`); python-calamine only; warmup=0 iters=1.

Command:
- `uv run python scripts/generate_throughput_fixtures.py`
- `uv run excelbench perf --tests test_files/throughput_xlsx --output results_dev_perf_throughput_small_calamine --warmup 0 --iters 1 --adapter python-calamine --feature cell_values_1k --feature formulas_1k`

Outputs:
- JSON: `results_dev_perf_throughput_small_calamine/perf/results.json`
- Markdown: `results_dev_perf_throughput_small_calamine/perf/README.md`
- CSV: `results_dev_perf_throughput_small_calamine/perf/matrix.csv`
- History: `results_dev_perf_throughput_small_calamine/perf/history.jsonl`

Observations:
- python-calamine per-cell reads: ~400ms for 1k cells (~2.5K cells/sec).

### 02/08/2026 02:40 PM PST (via pst-timestamp)

Git:
- commit: 4704676
- branch: master

Environment:
- OS: macOS 26.2 (25C5031i)
- CPU: Apple M4 Pro
- RAM: 24 GB
- Python (uv): 3.12.3
- Power: (not recorded)
- Notes: Throughput fixtures (8 scenarios including 1k + 10k variants); fast batch excluding python-calamine; warmup=0 iters=1.

Command:
- `uv run python scripts/generate_throughput_fixtures.py`
- `uv run excelbench perf --tests test_files/throughput_xlsx --output results_dev_perf_throughput_batch_fast --warmup 0 --iters 1 --adapter openpyxl --adapter xlsxwriter --adapter pylightxl --adapter pyexcel`

Outputs:
- JSON: `results_dev_perf_throughput_batch_fast/perf/results.json`
- Markdown: `results_dev_perf_throughput_batch_fast/perf/README.md`
- CSV: `results_dev_perf_throughput_batch_fast/perf/matrix.csv`
- History: `results_dev_perf_throughput_batch_fast/perf/history.jsonl`

### 02/08/2026 03:03 PM PST (via pst-timestamp)

Git:
- commit: 4704676
- branch: master

Environment:
- OS: macOS 26.2 (25C5031i)
- CPU: Apple M4 Pro
- RAM: 24 GB
- Python (uv): 3.12.3
- Power: (not recorded)
- Notes: Bulk read workload op (`bulk_sheet_values`) added; openpyxl only; warmup=0 iters=1.

Command:
- `uv run python scripts/generate_throughput_fixtures.py`
- `uv run excelbench perf --tests test_files/throughput_xlsx --output results_dev_perf_throughput_bulk --warmup 0 --iters 1 --adapter openpyxl --feature cell_values_10k_bulk_read --feature cell_values_1k_bulk_read`

Outputs:
- JSON: `results_dev_perf_throughput_bulk/perf/results.json`
- Markdown: `results_dev_perf_throughput_bulk/perf/README.md`
- CSV: `results_dev_perf_throughput_bulk/perf/matrix.csv`
- History: `results_dev_perf_throughput_bulk/perf/history.jsonl`

Observations:
- Workload is read-only (`operations: [read]`); write is skipped.

### 02/08/2026 03:08 PM PST (via pst-timestamp)

Git:
- commit: 4704676
- branch: master

Environment:
- OS: macOS 26.2 (25C5031i)
- CPU: Apple M4 Pro
- RAM: 24 GB
- Python (uv): 3.12.3
- Power: (not recorded)
- Notes: Bulk read scenarios executed across openpyxl + dataframe/dataset adapters; warmup=0 iters=1.

Command:
- `uv run excelbench perf --tests test_files/throughput_xlsx --output results_dev_perf_throughput_bulk_multi --warmup 0 --iters 1 --adapter openpyxl --adapter pandas --adapter polars --adapter tablib --adapter openpyxl-readonly --feature cell_values_10k_bulk_read --feature cell_values_1k_bulk_read`

Outputs:
- JSON: `results_dev_perf_throughput_bulk_multi/perf/results.json`
- Markdown: `results_dev_perf_throughput_bulk_multi/perf/README.md`
- CSV: `results_dev_perf_throughput_bulk_multi/perf/matrix.csv`
- History: `results_dev_perf_throughput_bulk_multi/perf/history.jsonl`

### 02/08/2026 03:13 PM PST (via pst-timestamp)

Git:
- commit: 4704676
- branch: master

Environment:
- OS: macOS 26.2 (25C5031i)
- CPU: Apple M4 Pro
- RAM: 24 GB
- Python (uv): 3.12.3
- Power: (not recorded)
- Notes: Bulk write workload op (`bulk_write_grid`) added; openpyxl + xlsxwriter only; warmup=0 iters=1.

Command:
- `uv run python scripts/generate_throughput_fixtures.py`
- `uv run excelbench perf --tests test_files/throughput_xlsx --output results_dev_perf_throughput_bulk_write --warmup 0 --iters 1 --adapter openpyxl --adapter xlsxwriter --feature cell_values_1k_bulk_write --feature cell_values_10k_bulk_write`

Outputs:
- JSON: `results_dev_perf_throughput_bulk_write/perf/results.json`
- Markdown: `results_dev_perf_throughput_bulk_write/perf/README.md`
- CSV: `results_dev_perf_throughput_bulk_write/perf/matrix.csv`
- History: `results_dev_perf_throughput_bulk_write/perf/history.jsonl`

### 02/08/2026 03:16 PM PST (via pst-timestamp)

Git:
- commit: 4704676
- branch: master

Environment:
- OS: macOS 26.2 (25C5031i)
- CPU: Apple M4 Pro
- RAM: 24 GB
- Python (uv): 3.12.3
- Power: (not recorded)
- Notes: Bulk write scenarios executed across multiple adapters; warmup=0 iters=1.

Command:
- `uv run excelbench perf --tests test_files/throughput_xlsx --output results_dev_perf_throughput_bulk_write_multi --warmup 0 --iters 1 --adapter openpyxl --adapter xlsxwriter --adapter pandas --adapter tablib --feature cell_values_1k_bulk_write --feature cell_values_10k_bulk_write`

Outputs:
- JSON: `results_dev_perf_throughput_bulk_write_multi/perf/results.json`
- Markdown: `results_dev_perf_throughput_bulk_write_multi/perf/README.md`
- CSV: `results_dev_perf_throughput_bulk_write_multi/perf/matrix.csv`
- History: `results_dev_perf_throughput_bulk_write_multi/perf/history.jsonl`
