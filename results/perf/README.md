# ExcelBench Performance Results

*Generated: 2026-10-03T00:00:03.786680+00:00*
*Profile: xlsx*
*Platform: Darwin-arm64*
*Python: 3.12.3*
*CPU: arm*
*Cores: 14 | Memory MB: 24576.0*
*Config: warmup=3 iters=25 iteration_policy=fixed breakdown=False*

## Run Receipt

Rerun of the 2026-10-02 perf snapshot (`results-2026-10-02/perf/`) with the WolfXL 2.0.8
PyPI wheel. Method, fixtures, adapter set, iteration counts, host, interpreter and every
non-WolfXL package version are unchanged from 2026-10-02. WolfXL is measured through its
public API (`wolfxl.load_workbook`, `wolfxl.Workbook`), like every other library.

### Source identity

- ExcelBench commit: `b4533551150b3d4932f914088855f43b63493532` (`origin/master`, published),
  tree `31a5f7e6237f43029120df6f89387129b5fdbc31`. `metadata.commit` in `results.json` is
  `null` because the harness reads `git rev-parse` and the staged tree is a git archive, not a
  checkout; the staging receipt below is the source identity.
- The perf harness (`src/excelbench/perf/`), the adapters (`src/excelbench/harness/adapters/`),
  `src/excelbench/cli.py` and `fixtures/excel/` are byte-identical to the 2026-10-02 run's
  source: `git diff --stat <base> b453355 -- src/excelbench/perf src/excelbench/harness/adapters src/excelbench/cli.py fixtures/excel`
  prints nothing for `<base>` = `a9b223c` (the commit the 2026-10-02 run recorded) and for
  `8db6afa` (its published equivalent).
- Staged from a clean `git worktree` at that commit (no dirty or untracked overlay).
- Staging receipt (`remote-source-sync run`):

```json
{"action": "run", "schema_version": "remote-source-sync.v1", "mode": "git-archive-tar-gzip-1+dirty-overlay",
 "head": "b4533551150b3d4932f914088855f43b63493532", "source_key": "b4533551150b3d4932f914088855f43b63493532",
 "tracked_files": 563, "dirty_files": 0, "deleted_files": 0, "excluded_generated_files": 0,
 "allowed_sensitive_files": 0, "target": "m4", "remote_dir": "~/wolf-lane/officelibs-perf-2026-10-03/src",
 "replaces_remote_directory": true, "transfer_ok": true}
```

- Fixture manifest `fixtures/excel/manifest.json` sha256
  `984073b67669fb0566b55e35a05447fbaaeb9f0da2d6de1a335baf61de9b382f`, the same as 2026-10-02.

### Host

- Same host as 2026-10-02: Apple Mac16,8, Apple M4 Pro, 14 cores, 24 GB RAM; macOS 26.6.2
  (25G83), Darwin 25.6.0. Headless worker with sleep disabled.
- Resident non-benchmark load: the same long-running background agent process as on 2026-10-02,
  at about one core for the whole window (not ours; not stopped). No other job of ours ran during
  the measurement.
- Gates in `run.sh`: wait until `date -u +%F` prints 2026-10-03, then wait until load1 <= 4.0
  (10 s poll, 10 min cap).
  - Date gate passed at 2026-10-03T00:00:02Z; the load gate passed on its first check.
  - Load before: `{ 2.75 3.27 3.08 }` at 2026-10-03T00:00:02Z.
  - Load after: `{ 2.93 3.29 3.08 }` at 2026-10-03T00:00:17Z.
  - This is the only scored attempt; nothing was discarded.

### Interpreter and packages

- CPython 3.12.3 (python-build-standalone via uv, Clang 17.0.6). Fresh `uv venv` (uv 0.12.5),
  default PyPI index, no index overrides.
- Pinned adapter packages (`pins.txt`): wolfxl 2.0.8, openpyxl 3.1.5, xlsxwriter 3.2.9,
  python-calamine 0.8.2, pylightxl 1.61, xlrd 2.0.2, pyexcel 0.7.6, pyexcel-xlsx 0.6.1,
  xlwt 1.3.0, pandas 3.0.6, polars 1.44.2, fastexcel 0.21.0, tablib 3.10.0. Each is the PyPI
  latest on 2026-10-03 (checked against the PyPI JSON API before the run); no pin needed a bump.
  `pip-freeze.txt` differs from 2026-10-02 only in the `wolfxl` line.
- WolfXL: `wolfxl.__version__ == "2.0.8"`, wheel `wolfxl-2.0.8-cp312-cp312-macosx_11_0_arm64.whl`
  (maturin 1.14.1, uploaded 2026-10-02T23:48:12Z), sha256
  `7000537dde2101880671d83f010b6d3f519d3c553c5ad18b9ad56dc87add5ab7` (equal to the PyPI digest;
  re-downloaded and hashed on the host). The installed `wolfxl/_rust.cpython-312-darwin.so`
  sha256 `b967a8009824badcb982eca9e4a727aa2a4d9b2acc59b8dc55c63b07da5ee643` equals the copy
  inside that wheel. Source: SynthGL/wolfxl-oss tag `v2.0.8`, commit
  `396be5a9f5bb82c2892c9236e3571515315c4bfb` (tag resolved through the GitHub API).

### Command

Run from the staged source root by `run.sh` (log: `run-log.txt`; wall 14.5 s, exit 0). The
command is identical to 2026-10-02:

```bash
excelbench perf --tests fixtures/excel --output ../out --profile xlsx \
  --warmup 3 --iters 25 --iteration-policy fixed --memory-mode getrusage \
  --feature cell_values --feature formulas --feature text_formatting --feature background_colors \
  --feature number_formats --feature alignment --feature borders --feature dimensions \
  --feature multiple_sheets --feature merged_cells --feature conditional_formatting \
  --feature data_validation --feature hyperlinks --feature images --feature pivot_tables \
  --feature comments --feature freeze_panes --feature named_ranges --feature tables \
  --adapter openpyxl --adapter xlsxwriter --adapter python-calamine --adapter wolfxl \
  --adapter pylightxl --adapter xlrd --adapter pyexcel --adapter xlwt --adapter pandas \
  --adapter xlsxwriter-constmem --adapter openpyxl-readonly --adapter polars --adapter tablib
```

- Warmup 3, recorded iterations 25, policy `fixed`, breakdown off, memory mode `getrusage`.
  The perf harness has no per-operation timeout.
- 19 features x 13 adapters, the same set and order as 2026-10-02 and 2026-04-29. Every adapter
  gets the same fixtures, warmup and iteration count; nothing is WolfXL-specific.
- Cold vs warm: the harness runs every (feature, adapter, operation) in one process, discards
  the first 3 iterations as warmup, and reports min/p50/p95 over the next 25. All numbers are
  warm in-process timings; cold start is not measured by this lane.

### Differences from 2026-10-02

- **WolfXL 2.0.5 to 2.0.8.** This is the only package change.
- **Non-WolfXL adapters:** no version, method or coverage change. The Run Issues list below is
  identical to 2026-10-02 (163 lines). Their p50 totals moved by -4% to +13% (most +1% to +8%)
  with unchanged code; this is run-to-run and host variance, and it is why the ratios against
  openpyxl below are taken within a single run.
- **WolfXL write.** The write total fell from 22.28 ms to 11.26 ms, and every one of the 19
  write features dropped (for example `cell_values` 1.18 to 0.59 ms). Read is unchanged within
  noise (13.46 to 13.90 ms, the same drift as the other adapters).
- **Version-isolation diagnostic** (`diagnostics/`, same host and harness, openpyxl as control,
  run back to back right after the scored run at 00:00:39Z to 00:00:57Z): WolfXL 2.0.5 read
  14.30 ms / write 26.01 ms vs 2.0.8 read 14.44 ms / write 12.23 ms, with openpyxl at
  24.21/30.67 and 24.24/29.38 ms. The write change follows the WolfXL version, not the host.
  The 2.0.5 wheel installed for the diagnostic carries the same `_rust` extension as the
  2026-10-02 run (RECORD sha256 `O-2nMvPe8nVwzwcWifeoaRoIGNCjRk4HX8Gxg8DNHcA`).

### Artifacts

- `results.json`, `README.md` (this file; summary tables below are harness output),
  `matrix.csv`, `history.jsonl`: harness output.
- `run.sh`, `run-log.txt`: the exact driver and its log (home prefix replaced with `~`, one
  unrelated process path redacted; `.txt` because the repo ignores `*.log`).
- `pins.txt`, `pip-freeze.txt`: requested pins and the resolved environment.
- `diagnostics/wolfxl-2.0.5-control.results.json`, `diagnostics/wolfxl-2.0.8-control.results.json`:
  the version-isolation diagnostic above (openpyxl + wolfxl only). Not part of the snapshot.

### Per-adapter p50 totals (sum of per-feature p50 wall ms; n = features measured)

Ratios against openpyxl are openpyxl's total divided by the library's total, within the same run.

| Adapter | 2026-10-02 version | 2026-10-03 version | Op | 2026-10-02 total (n) | 2026-10-03 total (n) | 2026-10-02 vs openpyxl | 2026-10-03 vs openpyxl |
|---|---|---|---|---|---|---|---|
| openpyxl | 3.1.5 | 3.1.5 | read | 22.79 (19) | 24.01 (19) | 1.00x | 1.00x |
| openpyxl | 3.1.5 | 3.1.5 | write | 26.70 (19) | 27.59 (19) | 1.00x | 1.00x |
| xlsxwriter | 3.2.9 | 3.2.9 | write | 31.12 (19) | 33.63 (19) | 0.86x | 0.82x |
| python-calamine | 0.8.2 | 0.8.2 | read | 1.63 (19) | 1.73 (19) | 14.02x | 13.92x |
| wolfxl | 2.0.5 | 2.0.8 | read | 13.46 (19) | 13.90 (19) | 1.69x | 1.73x |
| wolfxl | 2.0.5 | 2.0.8 | write | 22.28 (19) | 11.26 (19) | 1.20x | 2.45x |
| pylightxl | 1.61 | 1.61 | read | 10.55 (10) | 10.92 (10) | 2.16x | 2.20x |
| pylightxl | 1.61 | 1.61 | write | 1.21 (5) | 1.23 (5) | not comparable (5 features) | not comparable (5 features) |
| pyexcel | 0.7.6 | 0.7.6 | read | 24.92 (19) | 25.37 (19) | 0.91x | 0.95x |
| pyexcel | 0.7.6 | 0.7.6 | write | 7.28 (5) | 7.00 (5) | not comparable (5 features) | not comparable (5 features) |
| xlwt | 1.3.0 | 1.3.0 | write | 2.90 (14) | 3.28 (14) | not comparable (14 features) | not comparable (14 features) |
| pandas | 3.0.6 | 3.0.6 | read | 28.43 (19) | 29.33 (19) | 0.80x | 0.82x |
| pandas | 3.0.6 | 3.0.6 | write | 26.66 (19) | 28.57 (19) | 1.00x | 0.97x |
| xlsxwriter-constmem | 3.2.9 | 3.2.9 | write | 31.17 (19) | 33.69 (19) | 0.86x | 0.82x |
| openpyxl-readonly | 3.1.5 | 3.1.5 | read | 22.75 (19) | 23.52 (19) | 1.00x | 1.02x |
| polars | 1.44.2 | 1.44.2 | read | 6.13 (19) | 6.21 (19) | 3.72x | 3.87x |
| tablib | 3.10.0 | 3.10.0 | read | 23.09 (19) | 23.74 (19) | 0.99x | 1.01x |
| tablib | 3.10.0 | 3.10.0 | write | 23.58 (19) | 24.61 (19) | 1.13x | 1.12x |

xlrd has no timings (it does not read `.xlsx`). Leaders on this workload: python-calamine for
read (13.92x openpyxl); for write over all 19 features, WolfXL 2.0.8 (2.45x openpyxl).
pylightxl, pyexcel and xlwt write only a subset of features, so their write totals do not rank
against the 19-feature totals.

## Notes

These numbers measure only the library under test. Write timings do NOT include oracle verification.
Confidence note: treat deltas under ~5% as noise unless stable across multiple runs.

## Summary (p50 wall time)

**Tier 0 — Basic Values**

| Feature | openpyxl (R p50/p95 ms) | openpyxl (W p50/p95 ms) | openpyxl-readonly (R p50/p95 ms) | pandas (R p50/p95 ms) | pandas (W p50/p95 ms) | polars (R p50/p95 ms) | pyexcel (R p50/p95 ms) | pyexcel (W p50/p95 ms) | pylightxl (R p50/p95 ms) | pylightxl (W p50/p95 ms) | python-calamine (R p50/p95 ms) | tablib (R p50/p95 ms) | tablib (W p50/p95 ms) | wolfxl (R p50/p95 ms) | wolfxl (W p50/p95 ms) | xlrd (R p50/p95 ms) | xlsxwriter (W p50/p95 ms) | xlsxwriter-constmem (W p50/p95 ms) | xlwt (W p50/p95 ms) |
|---------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|
| cell_values | 1.30/1.62 | 1.29/1.51 | 5.92/6.17 | 1.77/1.93 | 1.57/1.85 | 0.45/0.72 | 1.49/1.52 | — | 1.32/1.33 | — | 0.65/0.66 | 1.43/1.69 | 1.35/1.45 | 0.69/0.74 | 0.59/0.67 | — | 1.70/2.00 | 1.73/2.03 | 0.28/0.33 |
| formulas | 1.14/1.19 | 1.41/1.59 | 1.49/1.56 | 1.59/1.72 | 1.67/1.76 | 0.39/0.47 | 1.26/1.31 | 1.45/1.52 | 1.13/1.19 | 0.26/0.28 | 0.14/0.14 | 1.19/1.25 | 1.43/1.66 | 0.63/0.70 | 0.66/0.73 | — | 1.61/2.07 | 1.77/2.18 | 0.28/0.30 |
| multiple_sheets | 1.28/1.33 | 1.75/1.95 | 1.01/1.04 | 1.79/1.87 | 2.08/2.27 | 0.51/0.62 | 1.40/1.54 | 1.72/1.96 | 1.31/1.32 | 0.31/0.41 | 0.06/0.06 | 1.35/1.47 | 1.69/1.89 | 0.52/0.58 | 0.57/0.62 | — | 1.97/2.71 | 2.07/2.54 | 0.20/0.21 |

**Tier 1 — Formatting**

| Feature | openpyxl (R p50/p95 ms) | openpyxl (W p50/p95 ms) | openpyxl-readonly (R p50/p95 ms) | pandas (R p50/p95 ms) | pandas (W p50/p95 ms) | polars (R p50/p95 ms) | pyexcel (R p50/p95 ms) | pyexcel (W p50/p95 ms) | pylightxl (R p50/p95 ms) | pylightxl (W p50/p95 ms) | python-calamine (R p50/p95 ms) | tablib (R p50/p95 ms) | tablib (W p50/p95 ms) | wolfxl (R p50/p95 ms) | wolfxl (W p50/p95 ms) | xlrd (R p50/p95 ms) | xlsxwriter (W p50/p95 ms) | xlsxwriter-constmem (W p50/p95 ms) | xlwt (W p50/p95 ms) |
|---------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|
| alignment | 1.21/1.25 | 1.36/1.42 | 0.97/1.01 | 1.44/1.51 | 1.44/1.59 | 0.25/0.31 | 1.27/1.29 | — | — | — | 0.05/0.05 | 1.24/1.29 | 1.22/1.36 | 0.67/0.73 | 0.55/0.69 | — | 1.80/2.17 | 1.72/2.16 | 0.24/0.36 |
| background_colors | 1.01/1.02 | 1.30/1.43 | 0.88/0.94 | 1.26/1.36 | 1.44/1.62 | 0.23/0.27 | 1.08/1.11 | — | 0.94/0.95 | — | 0.05/0.06 | 1.03/1.05 | 1.19/1.23 | 0.59/0.65 | 0.53/0.60 | — | 1.76/2.27 | 1.67/1.97 | 0.23/0.36 |
| borders | 1.85/2.00 | 2.15/2.44 | 1.31/1.41 | 2.05/2.32 | 1.59/1.79 | 0.26/0.32 | 1.79/2.27 | — | — | — | 0.06/0.07 | 1.80/2.16 | 1.37/1.55 | 1.03/1.06 | 0.82/0.87 | — | 2.19/2.62 | 2.32/2.62 | 0.43/0.58 |
| dimensions | 0.95/0.97 | 1.16/1.29 | 0.81/0.94 | 1.25/1.32 | 1.30/1.43 | 0.33/0.46 | 1.05/1.07 | — | 0.90/0.91 | — | 0.06/0.06 | 0.98/1.01 | 1.16/1.28 | 0.55/0.60 | 0.52/0.71 | — | 1.41/1.82 | 1.71/2.02 | 0.16/0.36 |
| number_formats | 1.03/1.06 | 1.27/1.41 | 0.87/0.92 | 1.27/1.30 | 1.39/1.45 | 0.24/0.30 | 1.09/1.19 | — | 0.94/1.02 | — | 0.06/0.06 | 1.05/1.11 | 1.18/1.35 | 0.61/0.67 | 0.53/0.57 | — | 1.58/1.89 | 1.70/2.06 | 0.22/0.28 |
| text_formatting | 1.72/1.80 | 1.70/1.82 | 1.46/1.48 | 2.01/2.12 | 1.57/1.84 | 0.25/0.30 | 1.78/1.82 | — | 1.49/1.53 | — | 0.08/0.08 | 1.78/1.86 | 1.33/1.43 | 0.85/0.91 | 0.66/0.74 | — | 1.99/2.34 | 2.00/2.22 | 0.36/0.39 |

**Tier 2 — Advanced**

| Feature | openpyxl (R p50/p95 ms) | openpyxl (W p50/p95 ms) | openpyxl-readonly (R p50/p95 ms) | pandas (R p50/p95 ms) | pandas (W p50/p95 ms) | polars (R p50/p95 ms) | pyexcel (R p50/p95 ms) | pyexcel (W p50/p95 ms) | pylightxl (R p50/p95 ms) | pylightxl (W p50/p95 ms) | python-calamine (R p50/p95 ms) | tablib (R p50/p95 ms) | tablib (W p50/p95 ms) | wolfxl (R p50/p95 ms) | wolfxl (W p50/p95 ms) | xlrd (R p50/p95 ms) | xlsxwriter (W p50/p95 ms) | xlsxwriter-constmem (W p50/p95 ms) | xlwt (W p50/p95 ms) |
|---------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|
| comments | 1.13/1.21 | 1.53/1.75 | 0.87/0.90 | 1.26/1.37 | 1.24/1.40 | 0.24/0.32 | 1.15/1.18 | — | 1.12/1.14 | — | 0.06/0.07 | 1.01/1.07 | 1.12/1.17 | 1.14/1.24 | 0.61/0.67 | — | 2.08/2.55 | 1.50/1.83 | — |
| conditional_formatting | 1.38/1.42 | 1.67/1.81 | 0.96/0.98 | 1.76/1.79 | 1.54/1.71 | 0.37/0.52 | 1.46/1.50 | — | — | — | 0.05/0.05 | 1.45/1.63 | 1.40/1.47 | 1.02/1.08 | 0.61/0.66 | — | 1.68/2.06 | 2.08/2.48 | — |
| data_validation | 1.21/1.36 | 1.22/1.28 | 0.90/0.94 | 1.64/1.76 | 1.27/1.42 | 0.48/0.63 | 1.32/1.34 | — | — | — | 0.05/0.05 | 1.29/1.32 | 1.13/1.23 | 0.59/0.65 | 0.53/0.59 | — | 1.51/1.87 | 1.60/1.92 | — |
| freeze_panes | 1.37/1.80 | 1.91/2.11 | 0.98/1.00 | 1.80/1.90 | 2.07/2.30 | 0.45/0.59 | 1.48/1.86 | — | — | — | 0.05/0.06 | 1.40/1.56 | 1.91/2.11 | 0.70/0.75 | 0.74/0.97 | — | 1.88/2.28 | 2.10/2.61 | 0.21/0.22 |
| hyperlinks | 1.19/1.21 | 1.26/1.41 | 0.88/0.98 | 1.55/1.56 | 1.28/1.44 | 0.37/0.48 | 1.28/1.39 | — | — | — | 0.05/0.05 | 1.22/1.26 | 1.14/1.25 | 1.15/1.21 | 0.54/0.58 | — | 1.75/1.94 | 1.82/2.13 | — |
| images | 1.63/2.13 | 1.62/1.82 | 0.79/0.82 | 1.16/1.26 | 1.26/1.45 | 0.22/0.26 | 1.21/1.26 | — | — | — | 0.05/0.05 | 0.94/1.02 | 1.14/1.19 | 0.87/0.91 | 0.69/0.73 | — | 2.49/2.89 | 1.47/2.01 | — |
| merged_cells | 1.24/1.27 | 1.33/1.59 | 0.88/0.92 | 1.32/1.34 | 1.43/1.61 | 0.23/0.29 | 1.31/1.36 | — | 0.95/0.97 | — | 0.06/0.06 | 1.07/1.09 | 1.16/1.23 | 0.82/0.87 | 0.54/0.57 | — | 1.65/2.10 | 1.60/2.03 | 0.20/0.36 |
| named_ranges | 1.11/1.14 | 1.41/1.47 | 0.87/0.90 | 1.57/2.06 | 1.78/2.04 | 0.37/0.46 | 1.24/1.34 | 1.49/4.76 | — | 0.23/0.25 | 0.05/0.05 | 1.19/1.21 | 1.36/3.56 | 0.49/0.53 | 0.54/0.57 | — | 1.57/1.99 | 1.68/2.17 | 0.16/0.17 |
| pivot_tables | 0.81/0.85 | 1.13/1.30 | 0.74/0.81 | 1.16/1.39 | 1.41/1.81 | 0.23/0.27 | 0.90/0.96 | 1.21/1.26 | 0.82/0.84 | 0.22/0.23 | 0.05/0.05 | 0.90/0.93 | 1.19/1.26 | 0.47/0.52 | 0.50/0.55 | — | 1.58/2.05 | 1.69/2.19 | 0.18/0.19 |
| tables | 1.44/1.85 | 1.13/1.30 | 0.93/0.95 | 1.68/1.78 | 1.23/1.30 | 0.37/0.54 | 1.81/2.23 | 1.13/1.20 | — | 0.21/0.22 | 0.05/0.06 | 1.40/1.42 | 1.14/1.20 | 0.52/0.56 | 0.49/0.52 | — | 1.45/1.83 | 1.46/1.79 | 0.14/0.15 |

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
