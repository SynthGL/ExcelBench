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

## Data-shape matrix (Sprint 2 — `feat/perf-data-shape`)

In addition to the legacy scenarios above, the generator emits a parametric
`(dtype × tier)` matrix used by `excelbench perf-shape`. Each (dtype, tier)
pair produces one fixture file plus two manifest entries (`_bulk_read` and
`_bulk_write`).

**Tiers**: `1k` (40×25), `10k` (100×100), `100k` (316×316), `1m` (1000×1000).
The `1m` tier is gated behind `--include-1m` because xlsxwriter generation
takes ~5 min for the full 1M slice. Default invocation emits 1k/10k/100k.

**Dtypes** (10 total):

- `int` — sequential integers
- `float` — sequential integers × 1.5 (forces float type)
- `string_short` — padded to 16 chars
- `string_long` — padded to 512 chars
- `boolean` — alternating True/False
- `date` — `2020-01-01 + N days`
- `datetime` — `2020-01-01 00:00:00 + N seconds`
- `formula_simple` — `=SUM(A{r}:B{r})` (per row)
- `formula_cross_sheet` — `=Sheet2!A{r}` (Sheet2 pre-populated)
- `mixed_realistic` — 60/30/5/3/2 mix (short string / int / date / formula /
  blank), calibrated from a 50-file survey; see
  `fixtures/synthetic_calibration/sample_set.md` and DEC-019 for provenance.

To emit only the shape matrix (skipping legacy scenarios) for fast iteration:

```bash
uv run python scripts/generate_throughput_fixtures.py --shape-only --include-1m
```

The 40 shape feature names follow the pattern
`data_shape_<dtype>_<tier>_bulk_<read|write>`, e.g.
`data_shape_formula_cross_sheet_1m_bulk_read`.
