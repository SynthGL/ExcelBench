# Throughput/Scale Fixtures (Performance)

These fixtures are for *performance benchmarking* (speed + best-effort memory).

They are intentionally separate from the canonical Excel-generated fixtures in `fixtures/excel/`.
The throughput fixtures use a compact `expected.workload` spec so we can describe large workloads
without writing 10k/100k test cases into `manifest.json`.

Implementation note:

- We generate these `.xlsx` files with `xlsxwriter` (not openpyxl) because some readers (notably
  pylightxl) can choke on openpyxl-emitted namespace placement in `xl/workbook.xml`.

Generate locally (default output is gitignored under `test_files/`):

```bash
uv run python scripts/generate_throughput_fixtures.py
uv run excelbench perf --tests test_files/throughput_xlsx --output results_dev_perf_throughput --warmup 1 --iters 5 --breakdown
```

Run the standard dashboard batches:

```bash
uv run python scripts/run_throughput_dashboard.py --warmup 0 --iters 1
```

Include python-calamine per-cell scenarios (1k only; bulk reads run by default):

```bash
uv run python scripts/run_throughput_dashboard.py --warmup 0 --iters 1 --include-slow
```

Currently generated scenarios:

- `cell_values_1k`
- `cell_values_1k_bulk_read`
- `cell_values_1k_bulk_write`
- `cell_values_10k`
- `cell_values_10k_bulk_read`
- `cell_values_10k_bulk_write`
- `cell_values_10k_sparse_1pct_bulk_write`
- `cell_values_10k_1000x10_bulk_read`
- `cell_values_10k_1000x10_bulk_write`
- `cell_values_10k_10x1000_bulk_read`
- `cell_values_10k_10x1000_bulk_write`
- `formulas_1k`
- `formulas_1k_bulk_read`
- `formulas_10k`
- `formulas_10k_bulk_read`
- `strings_unique_1k_bulk_read`
- `strings_unique_1k_bulk_write`
- `strings_unique_10k_bulk_read`
- `strings_unique_10k_bulk_write`
- `strings_repeated_10k_bulk_read`
- `strings_repeated_10k_bulk_write`
- `strings_unique_1k_len64_bulk_read`
- `strings_unique_1k_len64_bulk_write`
- `strings_unique_1k_len256_bulk_read`
- `strings_unique_1k_len256_bulk_write`
- `strings_repeated_1k_len256_bulk_read`
- `strings_repeated_1k_len256_bulk_write`
- `background_colors_1k`
- `number_formats_1k`
- `alignment_1k`
- `borders_200`

Optional: include ~100k cell fixture (slower to generate):

```bash
uv run python scripts/generate_throughput_fixtures.py --include-100k
```

When `--include-100k` is enabled, the manifest also includes:

- `cell_values_100k`
- `cell_values_100k_bulk_read`
- `cell_values_100k_bulk_write`
