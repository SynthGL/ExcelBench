# ExcelBench — Architecture & Design Decisions

> Purpose: Master log of significant design and architecture decisions.
> Reverse chronological (newest first). Every session that makes a material decision MUST add an
> entry here.
>
> Note: Entries up through DEC-012 were backfilled from git history and existing design docs on
> 2026-02-15. If the intent differs from what is written, edit the decision to match reality.

---

## How to Add a Decision

```markdown
### DEC-NNN — Short descriptive title (YYYY-MM-DD)

**Context**: What situation or problem prompted this decision?

**Decision**: What was decided? Be specific about the choice made.

**Alternatives considered**: What other options were evaluated? Why were they rejected?

**Consequences**: What follows from this decision? Any tradeoffs accepted?

**Commit(s)**: `abc1234` (optional)
```

When to log a decision:

- Architecture boundaries, dependency direction, new layers/modules
- Methodology/scoring changes that affect results comparability
- Fixture/oracle strategy changes
- Introducing new output formats or publishing/deploy workflows
- Public/private boundary shifts for adapters/backends
- Major performance methodology changes (workloads, measurement strategy)

Skip logging for routine bug fixes, refactors, or incremental test additions.

---

## Decisions

### DEC-021 — Optional external oracle subprocess contract (2026-04-28)

**Context**: WolfXL's pre-release parity pass is now green against the existing
openpyxl-centered matrix, but openpyxl does not construct or validate every
advanced OOXML structure that matters for a production-grade Excel library.
The next audit needs fixtures from tools such as Excelize, LibreOffice, Apache
POI, and ClosedXML without forcing Go/Java/.NET/LibreOffice onto every
ExcelBench install or CI job.

**Decision**:

- Add `src/excelbench/harness/external_oracles.py` as a subprocess-only JSON
  contract for optional non-Python helpers.
- External helpers receive one JSON request on stdin and return one JSON
  payload on stdout. Missing helper commands produce structured skips rather
  than failures.
- Keep external oracles out of `get_all_adapters()` until each helper and
  fixture pack is deterministic, audited, and ready for normal benchmark
  flows.
- Track the rollout in `docs/trackers/external-oracle-expansion.md`; initial
  reserved helpers are Excelize, LibreOffice, Apache POI, and ClosedXML.

**Alternatives considered**:

1. **Register each external tool as a normal adapter immediately** — rejected.
   That would make local runtimes and helper build systems part of ordinary
   benchmark discovery before the fixture semantics are stable.
2. **Use ad hoc scripts outside the package** — rejected. The contract needs
   unit tests, versioned semantics, and stable skip/failure behavior.
3. **Depend directly on Go/Java/.NET packages from Python** — rejected. It adds
   packaging complexity and hides process/runtime failures that should be
   visible in diagnostics.

**Consequences**:

- Core ExcelBench tests remain hermetic when optional helpers are missing.
- The first oracle sprint can focus on helper binaries/scripts and fixture
  truth-passing without changing benchmark scoring.
- Promotion into public results requires a later explicit decision once the
  generated workbooks are deterministic and manually audited.

**Commit(s)**: this commit.

### DEC-020 — File-shape parametric scenarios + n_sheets fan-out (2026-04-27)

**Context**: Sprint 2 (DEC-019) shipped the dtype × tier matrix that answers
*"how does this library handle int vs string vs formula at 1M cells?"*. The
orthogonal axis the dashboard still couldn't address is *layout cost* —
"does this library handle 100k cells the same way regardless of whether
they're 1000×100 (square), 10×10000 (wide), 100000×1 (tall), 1M-grid with
90% blanks (sparse), or split across 100 sheets?". Real libraries diverge
sharply on these: openpyxl loads each sheet on demand vs wolfxl/calamine
load everything upfront, sparse libs differ by 10× on blank-handling
strategies, and tall/wide expose row-iterator vs column-iterator code
paths. Sprint 3 closes the layout gap.

**Decision**:

- **5 shape categories × 3 tiers = 12 scenarios** (degenerate combos
  skipped). Categories: `wide` (many cols, few rows), `tall` (few cols,
  many rows), `sparse` (90% blank via `sparse_every=10`), `many_sheets`
  (same total cell count fanned across N sheets). Tiers are 10k / 100k /
  1M based on **total** cell count (not per-sheet).

- **Single dtype (`int`) held constant** across all shapes. File-shape
  cost is mostly dtype-independent, so cross-producting with the 10
  dtypes from S2 (= 60 scenarios) is ~5× the work for negligible
  additional signal. Holds the dtype axis steady so the file-shape
  dashboard heatmap stays clean.

- **`n_sheets` + `sheet_pattern` workload fields** added to the existing
  `bulk_write_grid` and `bulk_sheet_values` ops in
  `_run_workload_write` / `_run_workload_read`. Default `n_sheets=1`
  preserves single-sheet callers from S1/S2 unchanged. The pre-existing
  `sparse_every` covers the sparse case without runner additions.

- **`add_sheet` phase fans out**: `_measure_write_workload_iteration`
  pre-creates Sheet1..SheetN before `_run_workload_write` runs, so the
  many-sheets scenarios work with adapters whose `create_workbook()`
  starts empty (e.g. openpyxl).

- **`perf-file-shape` CLI subcommand** mirrors `perf-shape` exactly —
  separate command rather than overloading `perf-shape` with a
  `--shape-axis=file` flag, because *shape* is already overloaded
  between dtype-shape and file-shape and cramming both into one command
  hurts readability more than it saves wiring.

- **Dashboard tab** renders one heatmap per direction (read, write) with
  rows = library sorted by overall median, columns = 4 shape categories.
  Cell = ms / 100k cells at the largest tier each (lib, category) was
  run at; tooltip shows the per-tier curve. Color is log-scale,
  normalized **per category-column** so many-sheets per-sheet XML
  overhead doesn't wash out wide/tall/sparse columns.

- **op_count semantics**: for `n_sheets > 1`, op_count =
  per-sheet-cells × n_sheets (i.e. total cells touched across the fan-
  out). For `sparse_every > 1`, op_count = filled cells only (matches
  S2 convention). Both multipliers compose, so a many-sheets sparse
  scenario would correctly count `(filled_cells_per_sheet × n_sheets)` —
  not currently used in S3 but the math is composable.

**Alternatives considered**:

1. **Cross-product file-shape × all 10 dtypes (= 60 scenarios)** —
   rejected. ~5× the runtime + storage for marginal signal; file-shape
   cost is mostly orthogonal to dtype because adapters use the same
   row/col iterators regardless of cell content. If a future adapter
   shows dtype-dependent layout cost, that scenario can be added
   point-wise without re-running the full matrix.

2. **Extend `perf-shape` with `--shape-axis={dtype,file}` flag** —
   rejected. The word *shape* is already overloaded between dtype-shape
   and file-shape; cramming both behind one command makes the help text
   harder to read and the `--types`/`--shapes` flag separation harder
   to validate. Two commands with a shared helper extraction is cleaner.

3. **New op `bulk_write_grid_multi_sheet`** instead of extending
   `bulk_write_grid` with `n_sheets`** — rejected. Would duplicate the
   ~70-line value-generation block (10 `value_type` branches) for the
   sole benefit of dispatch separation. Single op with optional
   `n_sheets` field keeps the dispatch table small and means future
   value_types added in S4+ automatically work for many-sheets without
   per-feature plumbing.

4. **Sparse scenarios at 1k tier** — rejected. 1k grid × 10% sparse =
   100 filled cells, dominated by allocator setup costs and not
   representative of how libraries handle sparse storage at scale. 10k
   minimum keeps the signal honest.

5. **Streaming-read shapes (e.g. read-without-materialize)** — deferred
   to a future sprint. Requires opt-in adapter support beyond
   `read_sheet_values`, and openpyxl's `read_only=True` mode is the
   only adapter that meaningfully supports it today.

**Consequences**:

- 12 fixtures + ~150MB scratch on first cold run (1M tier dominates).
  Default (no `--include-1m`) emits 7 scenarios in <30s.
- Many-sheets scenarios stress the per-sheet XML overhead path. Early
  smoke runs on openpyxl show 100-sheet scenarios are ~3× slower per
  cell than the equivalent single-sheet 100k case — confirming the
  axis is informative.
- The 100x10k and 1000x1k many-sheets scenarios are individually 1M
  total cells but parsed differently than the 1000×1000 square case
  (the 1000-sheet scenario opens 1000 workbook entries). This is
  expected and is the headline data point S3 surfaces.
- `n_sheets` is a generic runner extension, not file-shape specific.
  Future sprints (S4 high-cost ops) can use it for "modify one cell
  across N sheets" workloads without further runner changes.
- New `file_shape_*` feature names are additive — schema unchanged.
  The dashboard's File Shape tab only renders when at least one
  `file_shape_*` entry is present in the loaded results.

**Commit(s)**: Sprint 3, branch `feat/perf-file-shape`.

---

### DEC-019 — Data-shape parametric scenarios + mixed-realistic ratio (2026-04-27)

**Context**: Sprint 1 of the 7-Dimension Extension shipped honest memory
measurement, but the perf manifest still grouped everything under a single
"feature" axis (cell_values, formulas, ...). That axis is too coarse to answer
the most actionable question users have when picking a library: *"how does
this library handle int vs string vs formula loads at 1M cells?"*. Real
libraries diverge by an order of magnitude across dtypes — openpyxl can be 5×
slower on `string_long` than `int`; wolfxl wins biggest on `formula_*` and
strings. Without dtype-axis data, the dashboard hides where each library is
actually weakest. Sprint 2 closes that gap.

The original sprint paragraph assumed Sprint 2 builds a new generator from
scratch. Discovery contradicted that: ExcelBench already has a parallel
"throughput fixtures" pipeline (`scripts/generate_throughput_fixtures.py` +
`fixtures/throughput_xlsx/manifest.json` + `_run_workload_write`/`_run_workload_read`
in the perf runner) that already supports parameterized cell counts and
`value_type ∈ {number, string}`. Sprint 2 became an *extension* of those
existing seams rather than a new infrastructure track.

**Decision**:

- **Matrix shape**: 10 dtypes × 4 cell-count tiers = 40 fixtures, each with
  one bulk-read and one bulk-write workload (80 manifest rows total). Tiers
  are 1k / 10k / 100k / 1M. Dtypes are int, float, string_short (≤16c),
  string_long (≤512c), boolean, date, datetime, formula_simple
  (`=SUM(A{r}:B{r})`), formula_cross_sheet (`=Sheet2!A{r}`), and
  mixed_realistic.

- **`mixed_realistic` ratio = 60/30/5/3/2** (short string / int / date /
  formula / blank), calibrated against a 50-file public xlsx survey
  documented at `fixtures/synthetic_calibration/sample_set.md`. The ratio is
  rounded from observed class-weighted means (58-63% / 27-32% / 4-7% /
  2-5% / 1-3%) and is deliberately deterministic per-cell-index so runs
  reproduce across libraries.

- **1M tier gated behind `--include-1m`** so `python scripts/generate_throughput_fixtures.py`
  default runs stay under 30s. Full 1M generation is ~5 min on the bench
  machine.

- **Generator extension, not rewrite**: new function
  `generate_data_shape_scenarios()` reuses the existing `_xlsx_workbook` /
  `_coord_to_cell` helpers; the runner extension adds 7 new branches
  (`float`, `date`, `datetime`, `boolean`, `formula_simple`,
  `formula_cross_sheet`, `mixed_realistic`) inside the existing `value_type`
  dispatch in `_run_workload_write`. `string_short` and `string_long` fold
  into the existing `string` op via `string_length=16` / `string_length=512`.

- **`perf-shape` CLI subcommand** is a thin wrapper: it computes the feature
  filter from `--rows` (largest tier) and `--types`, regenerates fixtures
  on-demand if the manifest is stale, and delegates to `run_perf` —
  inheriting Sprint 1's `--memory-mode` plumbing for free.

- **Dashboard tab** renders one heatmap per direction (read, write) with
  rows = library (sorted by overall median ms/100k), columns = 10 dtypes,
  cell = ms/100k cells at the largest tier the run has data for. Color is
  log-scale green→red, **normalized per dtype-column** so a slow column
  (formula_cross_sheet) doesn't wash out fast columns (int).

**Alternatives considered**:

1. **`value_type=any` mode that randomizes per-cell** — rejected. Cross-run
   reproducibility is more valuable than realism here; libraries running on
   different inputs can't be compared apples-to-apples. The deterministic
   `mixed_realistic` ratio gives the same per-cell-index distribution every
   time.

2. **Generate fixtures via openpyxl `write_only` mode** — rejected.
   `fixtures/throughput_xlsx/README.md` already documents that pylightxl
   chokes on openpyxl's namespace placement in `xl/workbook.xml`. Switching
   would silently break a downstream adapter.

3. **Sample real workbooks from EDGAR / public sources directly** —
   deferred. Licensing complexity (mixed sources, some scraping involved)
   plus manifest stability concerns (real files churn as upstreams update)
   would slow the sprint without proportional accuracy gain. The 50-file
   calibration is a reasonable proxy.

4. **One scenario per (dtype × tier) pair without separate read/write
   features** — rejected. The runner's `_workload_operations` already
   distinguishes by feature; collapsing them would lose the read-vs-write
   divergence (which is itself a key signal — wolfxl's read path and write
   path have very different cost profiles).

**Consequences**:

- 40 fixtures + ~250MB scratch on first cold run. Subsequent runs reuse
  cached fixtures keyed by manifest mtime vs generator-script mtime.
- Full read+write matrix at 1M tier completes in <30 min on the bench
  machine across the 16+ adapter set.
- New `data_shape_*` feature names are additive — existing perf consumers
  (history JSONL, perf README, perf CSV) accept arbitrary feature strings,
  so no schema changes are required. The dashboard only renders the new
  tab when at least one `data_shape_*` entry is present.
- The `formula_cross_sheet` dtype may surface adapter-specific quirks where
  some libs auto-evaluate formula values on write (returning numbers
  instead of formula strings). If that happens during full-run collection,
  the affected adapters can be skipped via `notes_parts` without re-planning
  this sprint.
- TODO (deferred): re-calibrate `mixed_realistic` against a >500-file
  corpus once a stable, well-licensed source is identified. The current 50-
  file sample is rounded conservatively; a larger survey could shift any
  ratio by 5-10 points but is not blocking.

**Commit(s)**: Sprint 2, branch `feat/perf-data-shape`.

---

### DEC-018 — Three coexisting memory-measurement modes (2026-04-27)

**Context**: Until Sprint 1 of the 7-Dimension Extension, the perf runner reported a single
memory number — peak RSS via `resource.getrusage(RUSAGE_SELF).ru_maxrss`. Two problems with
that single number: (1) `getrusage` returns the **process-lifetime peak**, so once the first
heavy iteration has allocated, subsequent iterations report the same sticky max even if they
allocated less — making per-iteration comparisons misleading; (2) ad-hoc 100k → 1M cell
benchmarks against openpyxl from the wolfxl 1.0 work showed the `getrusage` peak diverges
from `/usr/bin/time -l` peaks by 30-300% on large workloads, depending on Rust allocator
release behavior. There is no single right number — different libraries pay memory cost in
different places (Python heap vs Rust heap vs OS pages) and the perf dashboard needs to be
honest about that asymmetry rather than pretending a single column tells the truth.

**Decision**: Coexist three modes via a new `--memory-mode` flag, all populating the existing
`PerfOpResult` dataclass with separate fields:

- `getrusage` (default, cheap): in-process `RUSAGE_SELF.ru_maxrss`. Preserved as the hot-path
  default — fast, but documented as lifetime-peak-sticky.
- `tracemalloc`: in-process `tracemalloc.get_traced_memory()`. Adds the Python heap peak.
  Misleading for Rust-backed adapters (wolfxl, python-calamine, rust_xlsxwriter) because it
  cannot see native allocations; honest for pure-Python adapters.
- `time`: spawn each iteration under `/usr/bin/time -l` and parse peak RSS from stderr.
  Honest about Rust allocations because the OS reports it. Slow (subprocess startup +
  adapter import per iteration); quarterly deep-dive only.
- `all`: composite — every iteration runs in-process (capturing getrusage + tracemalloc) AND
  in a fresh subprocess (capturing time-l RSS). Used for the memory-deep-dive bench, not CI.

The dashboard renders `RSS (MB) — getrusage / time -l` as a dual cell with a tooltip
explaining the divergence whenever any entry has the time-l field populated.

**Alternatives considered**: (1) Replace `getrusage` with `psutil.Process().memory_info().rss`
polling — rejected: still in-process, still subject to allocator-release lag, and adds a
mandatory third-party dependency to the runner. (2) Drop `getrusage` once `time -l` is
available — rejected: `time -l` is 50-500x slower per iteration, breaking the CI hot path.
(3) Single `multi_mode` field instead of three separate fields — rejected: makes the JSON
schema lossy (can't tell which number came from which mode in a composite run).

**Consequences**: `PerfOpResult` JSON now carries `rss_kb_via_time` and `python_heap_peak_kb`
in addition to the existing `rss_peak_mb`. Downstream dashboards must accept these as
optional fields. The `time` and `all` modes spawn one subprocess per iteration via the new
private `excelbench.perf._iter_subprocess` module; on Windows where `/usr/bin/time` does not
exist, `rss_kb_via_time` is silently `None`. Comparisons across past results remain valid
because the existing `rss_peak_mb` field is unchanged.

**Commit(s)**: Sprint 1, branch `feat/perf-mem-honesty`.

---

### DEC-017 — Do not inject Excel alignment defaults in benchmark comparisons (2026-02-17)

**Context**: Several value-focused adapters return an empty `CellFormat()` for alignment reads/writes.
The harness previously injected Excel defaults (`h_align=general`, `v_align=bottom`) during
comparison, which created false-positive passes (notably `v_bottom`) even when no alignment
transformation happened.

**Decision**: Remove default alignment injection from the harness comparison path. Alignment checks now
use only values explicitly surfaced by the adapter/oracle. This prevents unsupported adapters from
earning non-zero alignment credit via implicit defaults.

**Alternatives considered**: (1) Keep default injection and only change fixtures (rejected: fragile,
still allows accidental matches). (2) Keep injection but add per-adapter exemptions (rejected:
complex, brittle, and hard to reason about). (3) Remove the `v_bottom` case entirely (rejected: this
remains a useful explicit-read test for adapters that truly report bottom alignment).

**Consequences**: Some previously non-zero alignment results drop to zero where support was not real.
Scores are stricter but more semantically accurate and less susceptible to default-value artifacts.

### DEC-016 — Extract WolfXL to standalone GitHub repo + PyPI (2026-02-15)

**Context**: WolfXL was embedded inside ExcelBench across `packages/wolfxl/` (Python wrapper) and
`rust/excelbench_rust/` (Rust backend). This made it unusable for anyone not building ExcelBench
from source. For adoption, WolfXL needs to be `pip install wolfxl`.

**Decision**: Extract WolfXL to `SynthGL/wolfxl` on GitHub. Publish the calamine fork as
`calamine-styles` crate on crates.io (required because cargo disallows git deps in published
crates). The standalone repo includes only the 3 core backends (calamine-styled, rust_xlsxwriter,
wolfxl patcher) — umya and basic calamine stay in ExcelBench. ExcelBench's `[project.optional-dependencies] rust`
now points to `wolfxl>=0.1.0` from PyPI instead of `maturin`.

**Alternatives considered**: (1) Keep WolfXL in ExcelBench and add maturin wheel CI (rejected:
couples product and benchmark releases). (2) Include all 5 backends in standalone (rejected: umya
and basic calamine are ExcelBench-only benchmarking tools). (3) Publish calamine fork under
original name (rejected: name collision on crates.io).

**Consequences**: `pip install wolfxl` provides pre-built wheels for Linux/macOS/Windows. ExcelBench
CI no longer needs Rust toolchain for WolfXL (just `pip install wolfxl`). The `excelbench_rust_shim`
package can be deprecated. Calamine fork must be maintained as `calamine-styles` on crates.io.

### DEC-015 — Publish WolfXL as standalone + `wolfxl._rust` with shim compatibility (2026-02-15)

**Context**: WolfXL started as an in-repo compatibility layer and used the native module name
`excelbench_rust`. To publish WolfXL independently and make branding/dependency boundaries clear,
the native module needed a WolfXL namespace while existing integrations still required compatibility.

**Decision**: Split WolfXL into standalone packages under `packages/` and brand the native module as
`wolfxl._rust` (`wolfxl-rust` distribution). Keep ExcelBench compatible by adding an
`excelbench-rust` shim distribution that re-exports `wolfxl._rust`, and keep runtime fallback import
logic in adapter utilities.

**Alternatives considered**: (1) Keep shipping WolfXL inside the `excelbench` package (rejected:
couples product and benchmark release cycles). (2) Keep native module name `excelbench_rust`
(rejected: mismatched branding for external users). (3) Break compatibility and require all callers
to migrate immediately (rejected: unnecessary migration friction).

**Consequences**: WolfXL can be released independently, with clearer product identity and dependency
boundaries. ExcelBench continues to function during transition via compatibility shim/fallback.
Documentation and error messages must consistently reference `wolfxl._rust` as primary and
`excelbench_rust` as legacy compatibility.

### DEC-014 — WolfXL: Surgical ZIP patcher for read-modify-write mode (2026-02-15)

**Context**: WolfXL's hybrid architecture (calamine read + rust_xlsxwriter write) cannot modify
existing files in place. The `load → modify → save` workflow is one of openpyxl's most common use
cases. Using umya-spreadsheet for this would match openpyxl's speed (both parse the full DOM),
defeating WolfXL's value proposition of Rust-backed speed.

**Decision**: Build WolfXL (`XlsxPatcher`) — a streaming XML patcher that treats .xlsx as a ZIP of
XML files. On save, it only parses and rewrites the worksheet XMLs that have dirty cells, patches
styles.xml only if formats changed, and copies other ZIP entries through the rewriter unchanged at
the file-content level (compressed bytes may differ). Uses inline strings (`t="str"`) to avoid
touching sharedStrings.xml entirely.

**Alternatives considered**: (1) Use umya-spreadsheet for R/W (rejected: parses full DOM, no faster
than openpyxl). (2) Full rewrite via calamine read + rust_xlsxwriter write (rejected: loses charts,
images, macros, VBA — destructive). (3) Python ZIP patcher with ElementTree (rejected: slower,
more memory). (4) Wait for calamine upstream R/W support (rejected: no timeline).

**Consequences**: WolfXL now has three modes: read-only (calamine), write-only (rust_xlsxwriter),
and modify (XlsxPatcher). Modify mode is 10-14x faster than openpyxl across file sizes (38KB→651KB).
Preserves images, hyperlinks, charts, comments, and other ZIP entries unchanged. Uses inline
strings for new values, which slightly increases file size vs shared strings but avoids SST mutation.

**Commit(s)**: `b64b497`, `ffc5cbd`, `15b1c18`, `266086e`

### DEC-013 — Separate pycalumya compat package in src/pycalumya/ (2026-02-15)

> **Note**: The package was originally created as `pycalumya` (`src/pycalumya/`) and later renamed
> to `wolfxl`, now located at `packages/wolfxl/src/wolfxl/`. References to `excelbench_rust` in
> this entry are historical and now map to `wolfxl._rust`.

**Context**: WolfXL has proven 3–12x faster than openpyxl with 17/18 feature fidelity. To drive
adoption, it needs an openpyxl-compatible API so users can switch with minimal code changes.

**Decision**: Create a separate `wolfxl` package namespace (not inside `excelbench`).
Dual-mode Workbook: `load_workbook()` wraps CalamineStyledBook for reading, `Workbook()` wraps
RustXlsxWriterBook for writing. Style dataclasses (Font, PatternFill, Border, Alignment) match
openpyxl's public names. No Rust changes needed — uses `excelbench_rust` directly.

**Alternatives considered**: (1) Embed inside `excelbench.compat` (rejected: circular import risk
and harder to publish standalone on PyPI). (2) Full openpyxl shim with read-modify-write (rejected:
calamine is read-only and rust_xlsxwriter is write-only — no shared state model).

**Consequences**: Future standalone PyPI publishing is trivial (`wolfxl` is already a
self-contained package). Users get `wb['Sheet1']['A1'].value` interface backed by Rust. Trade-off:
no read-modify-write support (fundamental limitation of the hybrid approach).

### DEC-012 — Memory profiling uses subprocess isolation (2026-02-14)

**Context**: In-process memory measurements are noisy and can cross-contaminate between adapters,
features, and iterations (allocator reuse, module caches, lingering objects).

**Decision**: Measure memory using subprocess isolation for each (adapter, operation, fixture)
execution, and report best-effort RSS + tracemalloc metrics as a complement to wall/cpu timings.

**Alternatives considered**: (1) In-process RSS snapshots (rejected: too noisy). (2) External
profilers only (rejected: not reproducible or easy to automate).

**Consequences**: Memory profiling runs are slower but more comparable and safer (one adapter cannot
poison another's memory baseline).

**Commit(s)**: `7b94655`, `0ecab5f`

### DEC-011 — Ship a hybrid "best-of-breed" Rust adapter (pycalumya, now WolfXL) (2026-02-14)

**Context**: No single library achieved the desired read and write fidelity/performance across all
scored features. Some libraries are excellent readers but limited writers (or vice versa).

**Decision**: Provide a hybrid adapter (`wolfxl`, originally named `pycalumya`) that composes the
fastest/highest-fidelity read backend with the best write backend, so users can benchmark a
realistic "production pairing".

**Alternatives considered**: (1) Require a single library per adapter (rejected: leaves a large gap
in the realistic Pareto frontier). (2) Keep hybrid logic out of ExcelBench (rejected: the benchmark
should represent practical configurations).

**Consequences**: Adds a composite adapter to the registry and requires careful version reporting and
capability labeling.

**Commit(s)**: `e5e78fd`, `f9d8b92`, `f809f97`

### DEC-010 — Results are published via a single-file HTML dashboard (2026-02-12)

**Context**: Markdown/CSV tables are useful but make it hard to explore multi-axis results (tiers,
read vs write, perf vs fidelity) and share them externally.

**Decision**: Generate a self-contained interactive HTML dashboard from results JSON, and provide an
auto-deploy workflow to publish updates.

**Alternatives considered**: (1) Only markdown reports (rejected: limited exploration). (2) A full
webapp with a backend (rejected: too heavy for a benchmark repo).

**Consequences**: The HTML output becomes a stable interface; schema changes to results JSON must be
backwards-compatible or carefully migrated.

**Commit(s)**: `f01758b`, `054193a`, `1be44ee`

### DEC-009 — Make structured diagnostics a first-class benchmark output (2026-02-10)

**Context**: A single numeric score does not explain failures. Adapter authors and users need fast,
reproducible insight into what mismatched (type vs value vs formatting) and where.

**Decision**: Store structured diagnostics in benchmark outputs (category/severity/test-case)
alongside scores, and render them in reports.

**Alternatives considered**: (1) Only log text output (rejected: not machine-aggregatable). (2)
Store only per-feature pass/fail (rejected: insufficient for debugging).

**Consequences**: Results JSON becomes more verbose but enables deterministic triage, filtering, and
trend tracking.

**Commit(s)**: `ebafaec`

### DEC-008 — Separate fidelity and performance tracks (2026-02-08)

**Context**: Fidelity runs require oracle verification (Excel/openpyxl) and are correctness-focused.
That overhead contaminates timing measurements and makes performance comparisons misleading.

**Decision**: Implement a separate `excelbench perf` track that reuses the same adapter surface area
but excludes oracle verification. Add scale/throughput fixtures to measure throughput where
correctness fixtures are too small and dominated by fixed overhead.

**Alternatives considered**: (1) Add perf timing to fidelity benchmark (rejected: oracle dominates
and mixes concerns). (2) Only microbenchmarks (rejected: not representative of end-to-end usage).

**Consequences**: Two result schemas/tracks must stay aligned in terminology but are intentionally
independent. Perf results are comparable within a machine, not across machines.

**Commit(s)**: `04656a3`, `68f397e`, `9b71b33`

### DEC-007 — Rust backends integrate via an optional PyO3 extension (2026-02-08)

**Context**: Rust libraries (calamine, rust_xlsxwriter, umya-spreadsheet) provide different
capabilities and performance characteristics. Maintaining a second harness would duplicate scoring,
fixtures, and reporting logic.

**Decision**: Keep ExcelBench's primary harness in Python and integrate Rust libraries via an
optional PyO3 extension module (`excelbench_rust`). Python adapters call into Rust and translate
results into the shared model contracts.

**Alternatives considered**: (1) Separate Rust benchmark runner (rejected: duplicated methodology).
(2) Replatform the whole project to maturin (rejected: increases packaging complexity and raises the
barrier to entry).

**Consequences**: Rust is a local optional extra; CI/headless users can still run the pure-Python
bench. Rust adapter contracts must remain stable and explicitly versioned.

**Commit(s)**: `b8a8eb4`

### DEC-006 — Canonical fixtures are Excel-generated and committed (2026-02-06)

**Context**: Using a library to generate its own test fixtures creates circular validation, and
re-generating fixtures in CI is fragile (requires Excel).

**Decision**: Generate fixtures by driving real Excel (xlwings) and commit the resulting fixtures
and manifest as the canonical ground truth used by CI and all benchmark runs.

**Alternatives considered**: (1) Generate fixtures with openpyxl/xlsxwriter (rejected: not ground
truth). (2) Generate in CI (rejected: Excel not available).

**Consequences**: Fixture generation is a special workflow requiring Excel installed. Updating
fixtures should be treated as a deliberate change with visible diffs.

**Commit(s)**: `d9d80bd`

### DEC-005 — Split xlsx and xls benchmark profiles (2026-02-06)

**Context**: `.xls` and `.xlsx` have fundamentally different formats, library support, and edge
cases. Mixing them in one run confuses scoring and capability reporting.

**Decision**: Provide separate benchmark profiles for xlsx and xls, including a dedicated `.xls`
fixture lane and adapter set.

**Alternatives considered**: (1) A single combined profile (rejected: hides format-specific gaps).
(2) Ignore `.xls` (rejected: it remains common in legacy workflows).

**Consequences**: Results are comparable within a profile; cross-profile comparisons should be
explicit.

**Commit(s)**: `4a6bfe0`

### DEC-004 — Results JSON is the source of truth (2026-02-04)

**Context**: ExcelBench needs to support multiple views (markdown, CSV, plots, dashboards) without
re-running benchmarks.

**Decision**: Store benchmark output as JSON as the source of truth and generate all other formats
from it.

**Alternatives considered**: (1) Render-only markdown tables (rejected: inflexible). (2) Multiple
independent output formats (rejected: drift and duplication).

**Consequences**: Result schema stability matters. New outputs should extend the JSON schema rather
than inventing parallel data stores.

**Commit(s)**: `7e8306a`

### DEC-003 — Fidelity scoring uses a 0-3 scale with tiered feature coverage (2026-02-04)

**Context**: Binary "supported/unsupported" is too coarse and does not reflect the reality of Excel
feature support (partial fidelity, edge-case gaps, read vs write asymmetry).

**Decision**: Score each feature on a 0-3 fidelity scale, with separate read and write scoring where
relevant. Organize features into tiers to prioritize common pain points first.

**Alternatives considered**: (1) Binary scoring (rejected: loses nuance). (2) A continuous numeric
metric only (rejected: hard to interpret and justify).

**Consequences**: Score changes must be accompanied by explicit rubric/fixture updates to preserve
reproducibility.

**Commit(s)**: `769cab2`, `7e8306a`

### DEC-002 — Use a unified adapter interface with capability-aware harness logic (2026-02-04)

**Context**: Excel libraries differ widely: some are read-only, some write-only, and some support a
subset of formatting/features.

**Decision**: Standardize on a single adapter interface with capability flags (read/write) and keep
feature normalization/scoring in the harness, not per adapter.

**Alternatives considered**: (1) Per-library custom harness logic (rejected: not scalable). (2)
Separate read and write harnesses (rejected: duplicated logic).

**Consequences**: Adapters remain thin shims; adding a new adapter is mostly mapping and optional
imports.

**Commit(s)**: `7e8306a`

### DEC-001 — Use real Excel as the ground truth fixture generator (2026-02-04)

**Context**: Testing Excel libraries requires an authoritative reference output. Using one library to
generate fixtures for others risks encoding that library's bugs as "expected".

**Decision**: Generate xlsx fixtures by driving the actual Excel application via xlwings, and use
those fixtures as the ground truth for fidelity benchmarking.

**Alternatives considered**: (1) Use openpyxl to generate fixtures (rejected: not ground truth). (2)
Hand-author OOXML (rejected: error-prone and non-representative).

**Consequences**: Fixture generation requires Excel installed and appropriate automation permissions.
CI uses committed fixtures rather than regenerating.

**Commit(s)**: `769cab2`
