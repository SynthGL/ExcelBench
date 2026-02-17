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
