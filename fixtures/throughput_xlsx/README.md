# Throughput/Scale Fixtures (Performance)

These fixtures are for *performance benchmarking* (speed + best-effort memory).

They are intentionally separate from the canonical Excel-generated fixtures in `fixtures/excel/`.
The throughput fixtures use a compact `expected.workload` spec so we can describe large workloads
without writing 10k/100k test cases into `manifest.json`.

Generate locally (default output is gitignored under `test_files/`):

```bash
uv run python scripts/generate_throughput_fixtures.py
uv run excelbench perf --tests test_files/throughput_xlsx --output results_dev_perf_throughput --warmup 1 --iters 5 --breakdown
```

Currently generated scenarios:

- `cell_values_1k`
- `cell_values_1k_bulk_read`
- `cell_values_1k_bulk_write`
- `cell_values_10k`
- `cell_values_10k_bulk_read`
- `cell_values_10k_bulk_write`
- `formulas_1k`
- `formulas_10k`
- `background_colors_1k`
- `number_formats_1k`
- `alignment_1k`
- `borders_200`

Optional: include ~100k cell fixture (slower to generate):

```bash
uv run python scripts/generate_throughput_fixtures.py --include-100k
```
