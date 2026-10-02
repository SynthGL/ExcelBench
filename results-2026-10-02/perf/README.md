# ExcelBench Performance Results

*Generated: 2026-10-02T08:33:41.795781+00:00*
*Profile: xlsx*
*Platform: Darwin-arm64*
*Python: 3.12.3*
*CPU: arm*
*Cores: 14 | Memory MB: 24576.0*
*Config: warmup=3 iters=25 iteration_policy=fixed breakdown=False*

## Run Receipt

Fresh rerun of the 2026-04-29 perf scenario (`results-release-2026-04-28/perf/`) with
the WolfXL 2.0.5 PyPI wheel, the same build every other 2026-10-02 lane uses. WolfXL
2.0.6 was released to PyPI at 2026-10-02T07:55Z and is **not** measured here.

### Adapter path change

The 2026-04-28 release snapshot (and the 2026-04-20 `results/perf/` snapshot) measured
WolfXL through private backend objects in `wolfxl._rust`, while every other library was
measured through its public API. Those WolfXL numbers are not comparable with the other
libraries' numbers and are superseded by this run, which measures WolfXL 2.0.5 through its
public API (`wolfxl.load_workbook`, `wolfxl.Workbook`) like every other library.

### Source identity

- ExcelBench commit: `a9b223c268ad2c68a5314bee51810619550607d4` (local branch `officelibs-m2`),
  tree `edc3fe0316aea9d6f38de6505ae30ae1fcb1fe0b`. Local history was squashed before
  publication, so that commit id is not on GitHub. Its published equivalent is
  `8db6afaebc3a5aef515da933608838780491a244`, which has the identical tree:
  `git rev-parse 8db6afa^{tree}` prints `edc3fe0316aea9d6f38de6505ae30ae1fcb1fe0b`.
  `metadata.commit` in `results.json` is `null` because the harness reads
  `git rev-parse` and the staged tree is a git archive, not a checkout; the staging
  receipt below is the source identity.
- Staged from a clean detached `git worktree` at that commit (no dirty or untracked overlay).
- Staging receipt (`remote-source-sync run`):

```json
{"action": "run", "schema_version": "remote-source-sync.v1", "mode": "git-archive-tar-gzip-1+dirty-overlay",
 "head": "a9b223c268ad2c68a5314bee51810619550607d4", "source_key": "a9b223c268ad2c68a5314bee51810619550607d4",
 "tracked_files": 551, "dirty_files": 0, "deleted_files": 0, "excluded_generated_files": 0,
 "allowed_sensitive_files": 0, "target": "m4", "remote_dir": "~/wolf-lane/officelibs-perf/src",
 "replaces_remote_directory": true, "transfer_ok": true}
```

- Fixture manifest `fixtures/excel/manifest.json` sha256
  `984073b67669fb0566b55e35a05447fbaaeb9f0da2d6de1a335baf61de9b382f` (identical locally and on
  the host). The 19 measured fixture files are byte-identical to those in 0bf80c9, the commit
  that checked in April's snapshot (April's recorded commit `5f98de9` was squashed and is not
  in history); the manifest only gained the three Tier 4 entries, which the `--feature`
  filter excludes.

### Host

- Apple Mac16,8, Apple M4 Pro, 14 cores, 24 GB RAM; macOS 26.6.2 (25G83), Darwin 25.6.0.
  Headless worker with sleep disabled. (`run_environment.cpu_model` reads `arm`
  because the harness uses `platform.processor()`.)
- Resident non-benchmark load: one long-running background agent process at ~97% of one
  core for the whole window (not stopped; it is not ours). No other job of ours ran during
  the measurement.
- 1-minute load gate: wait until load1 <= 4.0 (10 s poll, 10 min cap).
  - Load before: `{ 3.81 3.75 2.87 }` at 2026-10-02T08:33:41Z (gate passed after 70 s; load was
    7.59 at 08:32:30Z).
  - Load after: `{ 3.62 3.72 2.87 }` at 2026-10-02T08:33:54Z.
  - An earlier attempt at 08:31:59Z started at load1 9.63 (Spotlight `mds` at ~60% and an
    unrelated agent server at ~58% CPU) without the gate; it was discarded unread and is not
    part of this snapshot.

### Interpreter and packages

- CPython 3.12.3 (python-build-standalone via uv, Clang 17.0.6), matching April's 3.12.3.
  Fresh `uv venv` (uv 0.12.5), default PyPI index, no index overrides.
- Pinned adapter packages (`pins.txt`): wolfxl 2.0.5, openpyxl 3.1.5, xlsxwriter 3.2.9,
  python-calamine 0.8.2, pylightxl 1.61, xlrd 2.0.2, pyexcel 0.7.6, pyexcel-xlsx 0.6.1,
  xlwt 1.3.0, pandas 3.0.6, polars 1.44.2, fastexcel 0.21.0, tablib 3.10.0 (each the PyPI
  latest on 2026-10-02, except wolfxl). Full environment: `pip-freeze.txt`.
- WolfXL: `wolfxl.__version__ == "2.0.5"`, wheel `wolfxl-2.0.5-cp312-cp312-macosx_11_0_arm64.whl`
  (maturin 1.14.1), sha256 `599b6ac753b1d32e2e68eb54c37d060c66f6a65f8378d725cda4b15b2b4f5e2b`
  (equal to the PyPI digest; re-downloaded and hashed on the host). The installed
  `wolfxl/_rust.cpython-312-darwin.so` sha256
  `3beda732f3def27570cf071689f7a8691a0818d0a3464e075fc1b183c0cd1dc0` equals the copy inside
  that wheel.

### Command

Run from the staged source root by `run.sh` (log: `run-log.txt`; wall 13.1 s, exit 0):

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

- Warmup 3, recorded iterations 25, policy `fixed`, breakdown off, memory mode `getrusage`
  (all harness defaults, same as April). The perf harness has no per-operation timeout.
- The explicit feature and adapter lists reproduce April's 19 features and 13 adapters in
  April's order (the current manifest has 22 features). Every adapter gets the same
  fixtures, warmup, and iteration count; nothing is WolfXL-specific.
- Cold vs warm: the harness runs every (feature, adapter, operation) in one process,
  discards the first 3 iterations as warmup, and reports min/p50/p95 over the next 25.
  All reported numbers are warm in-process timings. Cold-start cost (interpreter start,
  imports, first-open) is not measured by this lane.

### Scenario differences from April

- **WolfXL adapter path.** April's adapter drove `wolfxl._rust` backend objects directly
  (`CalamineStyledBook` for reads, a native writer book for writes; version string
  `cal=2.0.0+rxw=2.0.0`). Since 2026-09-07 (5524cac) the adapter uses WolfXL's public
  openpyxl-compatible API (`wolfxl.load_workbook`, `wolfxl.Workbook`), the same kind of
  surface every other library is measured through. The published 2.0.0 and 2.0.5 wheels do
  not expose `CalamineStyledBook`, so April's adapter cannot run against them.
  Diagnostic (`diagnostics/`, same host, same harness, openpyxl as control, run back to
  back at 08:35Z): p50 totals over the 19 features are WolfXL 2.0.0 read 13.57 ms / write
  22.13 ms vs 2.0.5 read 13.82 ms / write 22.86 ms (openpyxl 23.16/26.57 vs 22.96/26.64).
  The change from April's WolfXL numbers comes from the adapter path, not a 2.0.0 to 2.0.5
  regression.
- **Unsupported writes are now failures, not timed no-ops** (a7ef0b8). pyexcel and pylightxl
  write timings now cover 5 features (April: 19) and xlwt 14 (April: 19); the rest appear
  under Run Issues as `UnsupportedAdapterOperationError`. Their write totals are not
  comparable to April's without restricting to common features.
- **Library versions** moved for python-calamine (0.6.1 to 0.8.2), pyexcel (0.7.4 to 0.7.6),
  pandas (3.0.0 to 3.0.6), polars (1.38.1 to 1.44.2), and tablib (3.9.0 to 3.10.0).
- **Host.** April's results record only `Darwin-arm64`; the April host model is not recorded,
  so cross-snapshot absolute deltas mix host and version effects.
- `excel_version` 16.105.3 in `results.json` is the fixture manifest's generator metadata,
  not software used by this run.

### Artifacts

- `results.json`, `README.md` (this file; summary tables below are harness output),
  `matrix.csv`, `history.jsonl`: harness output.
- `run.sh`, `run-log.txt`: the exact driver and its log (home prefix replaced with `~`, one
  unrelated process path redacted; `.txt` because the repo ignores `*.log`).
- `pins.txt`, `pip-freeze.txt`: requested pins and the resolved environment.
- `diagnostics/wolfxl-2.0.0-current-adapter.results.json`,
  `diagnostics/wolfxl-2.0.5-current-adapter.results.json`: the WolfXL version-isolation
  diagnostic above (openpyxl + wolfxl only). Not part of the snapshot.

### Per-adapter p50 totals vs April (sum of per-feature p50 wall ms; n = features measured)

| Adapter | April version | Oct version | Op | April total (n) | Oct total (n) | Oct / April on common features |
|---|---|---|---|---|---|---|
| openpyxl | 3.1.5 | 3.1.5 | read | 25.07 (19) | 22.79 (19) | 0.91x |
| openpyxl | 3.1.5 | 3.1.5 | write | 31.68 (19) | 26.70 (19) | 0.84x |
| xlsxwriter | 3.2.9 | 3.2.9 | write | 39.64 (19) | 31.12 (19) | 0.79x |
| python-calamine | 0.6.1 | 0.8.2 | read | 2.16 (19) | 1.63 (19) | 0.75x |
| wolfxl | cal=2.0.0+rxw=2.0.0 | 2.0.5 | read | 4.19 (19) | 13.46 (19) | 3.21x (adapter path changed) |
| wolfxl | cal=2.0.0+rxw=2.0.0 | 2.0.5 | write | 3.64 (19) | 22.28 (19) | 6.12x (adapter path changed) |
| pylightxl | 1.61 | 1.61 | read | 11.69 (10) | 10.55 (10) | 0.90x |
| pylightxl | 1.61 | 1.61 | write | 6.14 (19) | 1.21 (5) | 0.81x (5 common) |
| pyexcel | 0.7.4 | 0.7.6 | read | 27.66 (19) | 24.92 (19) | 0.90x |
| pyexcel | 0.7.4 | 0.7.6 | write | 29.77 (19) | 7.28 (5) | 0.86x (5 common) |
| xlwt | 1.3.0 | 1.3.0 | write | 3.84 (19) | 2.90 (14) | 0.93x (14 common) |
| pandas | 3.0.0 | 3.0.6 | read | 31.96 (19) | 28.43 (19) | 0.89x |
| pandas | 3.0.0 | 3.0.6 | write | 34.71 (19) | 26.66 (19) | 0.77x |
| xlsxwriter-constmem | 3.2.9 | 3.2.9 | write | 39.51 (19) | 31.17 (19) | 0.79x |
| openpyxl-readonly | 3.1.5 | 3.1.5 | read | 23.11 (19) | 22.75 (19) | 0.98x |
| polars | 1.38.1 | 1.44.2 | read | 6.88 (19) | 6.13 (19) | 0.89x |
| tablib | 3.9.0 | 3.10.0 | read | 24.45 (19) | 23.09 (19) | 0.94x |
| tablib | 3.9.0 | 3.10.0 | write | 29.45 (19) | 23.58 (19) | 0.80x |

xlrd has no timings in either snapshot (it does not read `.xlsx`).

## Notes

These numbers measure only the library under test. Write timings do NOT include oracle verification.
Confidence note: treat deltas under ~5% as noise unless stable across multiple runs.

## Summary (p50 wall time)

**Tier 0 — Basic Values**

| Feature | openpyxl (R p50/p95 ms) | openpyxl (W p50/p95 ms) | openpyxl-readonly (R p50/p95 ms) | pandas (R p50/p95 ms) | pandas (W p50/p95 ms) | polars (R p50/p95 ms) | pyexcel (R p50/p95 ms) | pyexcel (W p50/p95 ms) | pylightxl (R p50/p95 ms) | pylightxl (W p50/p95 ms) | python-calamine (R p50/p95 ms) | tablib (R p50/p95 ms) | tablib (W p50/p95 ms) | wolfxl (R p50/p95 ms) | wolfxl (W p50/p95 ms) | xlrd (R p50/p95 ms) | xlsxwriter (W p50/p95 ms) | xlsxwriter-constmem (W p50/p95 ms) | xlwt (W p50/p95 ms) |
|---------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|
| cell_values | 1.26/1.55 | 1.22/1.29 | 5.69/5.94 | 1.71/1.82 | 1.48/1.55 | 0.32/0.42 | 1.40/1.44 | — | 1.26/1.28 | — | 0.62/0.63 | 1.37/1.40 | 1.30/1.38 | 0.66/0.73 | 1.18/1.28 | — | 1.59/2.03 | 1.59/2.00 | 0.22/0.28 |
| formulas | 1.11/1.13 | 1.35/1.41 | 1.43/1.67 | 1.52/1.55 | 1.60/1.86 | 0.41/0.54 | 1.19/1.24 | 1.39/1.47 | 1.12/1.14 | 0.26/0.29 | 0.12/0.12 | 1.16/1.19 | 1.37/1.43 | 0.57/0.63 | 1.17/1.33 | — | 1.53/1.99 | 1.71/2.17 | 0.26/0.34 |
| multiple_sheets | 1.31/1.41 | 1.74/1.93 | 0.97/1.01 | 1.76/1.84 | 2.02/2.37 | 0.52/0.69 | 1.65/2.73 | 2.31/2.69 | 1.29/1.44 | 0.32/0.38 | 0.06/0.06 | 1.31/1.35 | 1.71/2.52 | 0.54/0.58 | 1.30/1.50 | — | 1.95/2.41 | 2.00/2.45 | 0.18/0.18 |

**Tier 1 — Formatting**

| Feature | openpyxl (R p50/p95 ms) | openpyxl (W p50/p95 ms) | openpyxl-readonly (R p50/p95 ms) | pandas (R p50/p95 ms) | pandas (W p50/p95 ms) | polars (R p50/p95 ms) | pyexcel (R p50/p95 ms) | pyexcel (W p50/p95 ms) | pylightxl (R p50/p95 ms) | pylightxl (W p50/p95 ms) | python-calamine (R p50/p95 ms) | tablib (R p50/p95 ms) | tablib (W p50/p95 ms) | wolfxl (R p50/p95 ms) | wolfxl (W p50/p95 ms) | xlrd (R p50/p95 ms) | xlsxwriter (W p50/p95 ms) | xlsxwriter-constmem (W p50/p95 ms) | xlwt (W p50/p95 ms) |
|---------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|
| alignment | 1.18/1.22 | 1.32/1.40 | 0.92/0.96 | 1.42/1.54 | 1.35/1.41 | 0.22/0.27 | 1.23/1.27 | — | — | — | 0.05/0.05 | 1.21/1.38 | 1.16/1.22 | 0.65/0.69 | 1.10/1.18 | — | 1.58/1.85 | 1.63/1.99 | 0.22/0.23 |
| background_colors | 0.98/1.03 | 1.20/1.25 | 0.84/0.87 | 1.20/1.30 | 1.30/1.36 | 0.24/0.32 | 1.04/1.08 | — | 0.90/0.91 | — | 0.05/0.06 | 0.99/1.02 | 1.11/1.19 | 0.58/0.65 | 1.12/1.23 | — | 1.57/1.97 | 1.55/1.97 | 0.18/0.19 |
| borders | 1.68/1.80 | 2.06/2.32 | 1.27/1.31 | 2.00/2.27 | 1.48/1.54 | 0.26/0.32 | 1.76/1.80 | — | — | — | 0.06/0.07 | 1.76/2.10 | 1.32/1.48 | 0.98/1.01 | 1.40/1.50 | — | 2.01/2.46 | 1.95/2.32 | 0.39/0.45 |
| dimensions | 0.94/1.04 | 1.24/1.30 | 0.84/0.93 | 1.46/1.63 | 1.29/1.43 | 0.36/0.45 | 1.01/1.05 | — | 0.87/0.89 | — | 0.05/0.05 | 1.03/1.18 | 1.24/1.37 | 0.52/0.56 | 1.04/1.18 | — | 1.35/1.85 | 1.64/2.24 | 0.16/0.22 |
| number_formats | 1.01/1.06 | 1.17/1.20 | 0.86/0.88 | 1.23/1.28 | 1.32/1.71 | 0.24/0.29 | 1.06/1.12 | — | 0.93/0.95 | — | 0.06/0.06 | 1.01/1.11 | 1.12/1.17 | 0.59/0.63 | 1.08/1.34 | — | 1.39/1.88 | 1.74/1.97 | 0.19/0.24 |
| text_formatting | 1.67/1.70 | 1.79/1.85 | 1.42/1.46 | 1.94/1.96 | 1.44/1.54 | 0.25/0.30 | 1.73/1.79 | — | 1.42/1.46 | — | 0.07/0.08 | 1.73/1.85 | 1.25/1.32 | 0.84/0.90 | 1.22/1.30 | — | 1.82/2.23 | 1.92/2.22 | 0.34/0.63 |

**Tier 2 — Advanced**

| Feature | openpyxl (R p50/p95 ms) | openpyxl (W p50/p95 ms) | openpyxl-readonly (R p50/p95 ms) | pandas (R p50/p95 ms) | pandas (W p50/p95 ms) | polars (R p50/p95 ms) | pyexcel (R p50/p95 ms) | pyexcel (W p50/p95 ms) | pylightxl (R p50/p95 ms) | pylightxl (W p50/p95 ms) | python-calamine (R p50/p95 ms) | tablib (R p50/p95 ms) | tablib (W p50/p95 ms) | wolfxl (R p50/p95 ms) | wolfxl (W p50/p95 ms) | xlrd (R p50/p95 ms) | xlsxwriter (W p50/p95 ms) | xlsxwriter-constmem (W p50/p95 ms) | xlwt (W p50/p95 ms) |
|---------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|--------------|
| comments | 1.07/1.09 | 1.35/1.45 | 0.84/0.88 | 1.17/1.22 | 1.22/1.28 | 0.23/0.27 | 1.13/1.27 | — | 1.08/1.10 | — | 0.05/0.06 | 0.97/1.02 | 1.07/1.27 | 1.11/1.19 | 1.21/1.27 | — | 1.77/2.27 | 1.36/1.62 | — |
| conditional_formatting | 1.35/1.40 | 1.59/1.76 | 0.91/0.95 | 1.71/1.73 | 1.43/1.56 | 0.36/0.50 | 1.45/1.60 | — | — | — | 0.05/0.05 | 1.38/1.41 | 1.32/1.44 | 0.99/1.04 | 1.22/1.28 | — | 1.62/2.08 | 1.72/2.14 | — |
| data_validation | 1.20/1.24 | 1.19/1.34 | 0.85/0.92 | 1.61/1.75 | 1.18/1.25 | 0.51/0.69 | 1.31/1.62 | — | — | — | 0.05/0.05 | 1.32/1.48 | 1.05/1.10 | 0.58/0.64 | 1.08/1.11 | — | 1.41/1.77 | 1.46/1.85 | — |
| freeze_panes | 1.33/1.36 | 1.85/2.03 | 0.95/0.99 | 1.73/1.88 | 2.00/2.28 | 0.44/0.52 | 1.45/1.52 | — | — | — | 0.05/0.05 | 1.36/1.41 | 1.81/2.01 | 0.67/0.71 | 1.30/1.49 | — | 1.75/2.27 | 2.01/2.60 | 0.18/0.20 |
| hyperlinks | 1.14/1.17 | 1.18/1.21 | 0.89/1.02 | 1.51/1.55 | 1.19/1.32 | 0.36/0.48 | 1.26/1.35 | — | — | — | 0.05/0.05 | 1.21/1.28 | 1.08/1.15 | 1.12/1.17 | 1.13/1.28 | — | 1.68/2.12 | 1.68/1.89 | — |
| images | 1.10/1.17 | 1.65/1.76 | 0.77/0.81 | 1.11/1.15 | 1.17/1.24 | 0.22/0.27 | 1.17/1.24 | — | — | — | 0.05/0.05 | 0.93/0.95 | 1.06/1.13 | 0.85/0.90 | 1.42/1.55 | — | 2.32/2.80 | 1.39/1.82 | — |
| merged_cells | 1.19/1.23 | 1.28/1.32 | 0.86/0.88 | 1.26/1.28 | 1.35/1.43 | 0.23/0.27 | 1.27/1.30 | — | 0.93/0.95 | — | 0.05/0.06 | 1.05/1.06 | 1.13/1.22 | 0.80/0.87 | 1.10/1.37 | — | 1.49/2.04 | 1.49/1.94 | 0.18/0.19 |
| named_ranges | 1.09/1.13 | 1.37/1.48 | 0.85/0.87 | 1.45/1.51 | 1.46/1.51 | 0.41/0.50 | 1.20/1.43 | 1.39/1.51 | — | 0.23/0.23 | 0.05/0.05 | 1.15/1.17 | 1.33/1.44 | 0.47/0.52 | 1.13/1.19 | — | 1.65/2.20 | 1.57/1.81 | 0.15/0.16 |
| pivot_tables | 0.81/0.84 | 1.08/1.13 | 0.69/0.75 | 1.03/1.08 | 1.18/1.27 | 0.20/0.25 | 0.88/0.92 | 1.08/1.13 | 0.77/0.83 | 0.20/0.21 | 0.05/0.05 | 0.82/0.85 | 1.05/1.12 | 0.47/0.51 | 1.03/1.12 | — | 1.31/1.69 | 1.39/1.91 | 0.12/0.13 |
| tables | 1.40/1.44 | 1.08/1.13 | 0.90/1.05 | 1.61/1.68 | 1.20/1.25 | 0.37/0.48 | 1.72/2.02 | 1.11/1.73 | — | 0.20/0.22 | 0.05/0.05 | 1.33/1.46 | 1.08/1.12 | 0.51/0.53 | 1.04/1.07 | — | 1.33/1.70 | 1.38/1.79 | 0.13/0.13 |

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
