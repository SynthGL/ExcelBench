# Phase 3: Rust Library Integration (PyO3) - Plan + Progress Tracker

> Historical naming note: this document predates the module rename. References to the Python
> module `excelbench_rust` now map to `wolfxl._rust`. The crate path `rust/excelbench_rust/`
> remains current.

Created: 02/07/2026 11:13 PM PST (via pst-timestamp)

## Working State (keep updated)

Current focus:
- [x] DOC-001 (Done): Document rust build workflow

Current blockers:
- None

Quick links:
- Tracker: this document
- Design context: `docs/plans/2026-02-04-excelbench-design.md`
- Adapter registry: `src/excelbench/harness/adapters/__init__.py`

## Goal
Integrate Rust Excel libraries into the existing Python benchmark harness via a PyO3 extension module.

Target libraries:
- `calamine` (read)
- `rust_xlsxwriter` (write)
- `umya-spreadsheet` (read/write)

Primary outcome:
- Running `uv run excelbench benchmark` includes Rust-backed adapters in the results matrix without requiring a second harness.

## Non-goals (for this phase)
- Switching the main `excelbench` package build backend from hatchling to maturin.
- Publishing Rust wheels to PyPI (local dev + CI buildability is sufficient initially).
- Full Tier 2 parity for Rust libraries on day 1 (start with Tier 1, then expand).

## Current State (repo reality check)
- Adapters are optional-import and registered in `src/excelbench/harness/adapters/__init__.py`.
- A Rust-backed reader already exists via `python-calamine` in `src/excelbench/harness/adapters/calamine_adapter.py`.
- `pyproject.toml` defines `maturin` under `[project.optional-dependencies].rust`.

## Architecture

Approach: keep the main project as pure-Python (hatchling), and add a separate Rust extension module that is imported optionally.

High-level flow:
1. Rust extension module (PyO3) exposes minimal workbook operations.
2. Python adapters call into the Rust module and translate results into ExcelBench models (`CellValue`, `CellFormat`, etc.).
3. Existing runner/scoring logic remains unchanged.

Recommended module name:
- Python module: `excelbench_rust`

Recommended layout:
```
rust/
  excelbench_rust/
    Cargo.toml
    pyproject.toml
    src/lib.rs
    src/calamine_backend.rs
    src/rust_xlsxwriter_backend.rs
    src/umya.rs
src/excelbench/harness/adapters/
  rust_calamine_adapter.py
  rust_xlsxwriter_adapter.py
  umya_adapter.py
```

## Compatibility + Conventions

### Data contract (Rust -> Python)
Prefer returning plain Python dict payloads that mirror what the harness already compares.

Cell value payload:
```
{"type": "blank"}
{"type": "string", "value": "Hello"}
{"type": "number", "value": 1.23}
{"type": "boolean", "value": true}
{"type": "error", "value": "#DIV/0!"}
{"type": "formula", "formula": "=A1+B1", "value": "=A1+B1"}
{"type": "date", "value": "2026-02-04"}
{"type": "datetime", "value": "2026-02-04T10:30:00"}
```

Notes:
- Dates/datetimes can be returned as ISO strings to avoid timezone/naive pitfalls.
- Keep the contract intentionally small; formatting/borders can be added later.

### Version reporting
`LibraryInfo.version` is a string. Recommended for Rust adapters:
- `version = excelbench_rust.build_info()["backends"]["calamine"]` (or similar)

If backend versions are hard to extract, use:
- `version = excelbench_rust.__version__` initially, and add backend versions later.

## Milestones (Definition of Done)

M0 - Decisions + preflight
- [ ] Decide whether to keep `python-calamine` adapter alongside new `calamine` binding (recommended: keep both initially).
- [ ] Confirm module name + directory layout.

M1 - Rust extension module skeleton builds locally
- [x] `uv sync --extra rust` works.
- [x] `uv run maturin develop ...` produces importable `excelbench_rust`.
- [x] `python -c "import excelbench_rust; print(excelbench_rust.build_info())"` works.

M2 - Calamine read path integrated end-to-end
- [x] Rust binding: open workbook, sheet names, read cell values.
- [x] Python adapter: shows up in `get_all_adapters()` when module is installed.
- [x] Benchmark run includes the adapter and produces scores.

M3 - rust_xlsxwriter write path integrated end-to-end
- [x] Rust binding: create workbook, add sheets, write Tier 1 cell values, save.
- [x] Python adapter: write-only adapter works with existing verification path.

M4 - umya-spreadsheet integrated (Tier 1 minimum)
- [x] Read basic values and sheet names.
- [x] Write basic values and save.

M5 - CI/dev workflow documented
- [x] `README.md` includes local build instructions.
- [ ] Optional: a script under `scripts/` standardizes the build.

## Tracker (Task IDs)

Status legend:
- Not started / In progress / Blocked / Done

### T0: Project setup + scaffolding

- [x] RUST-001 (Done): Create `rust/excelbench_rust/` crate skeleton
  - Deliverables: `Cargo.toml`, `pyproject.toml`, `src/lib.rs`
  - DoD: `uv run maturin develop --manifest-path rust/excelbench_rust/Cargo.toml` succeeds

- [x] RUST-002 (Done): Implement `build_info()` in Rust
  - Output: dict with wrapper version + enabled backends
  - DoD: `excelbench_rust.build_info()` returns stable keys

- [x] RUST-003 (Done): Add feature flags per backend (`calamine`, `rust_xlsxwriter`, `umya`)
  - DoD: crate compiles with only one backend enabled

- [x] RUST-004 (Done): Expose dependency versions via `build_info()`
  - DoD: `excelbench_rust.build_info()["backend_versions"]` includes resolved crate versions

- [x] RUST-005 (Done): Commit Rust lockfile policy
  - DoD: `rust/excelbench_rust/Cargo.lock` is not ignored and can be committed for reproducible CI builds

### T1: Calamine (read)

- [x] RUST-010 (Done): Implement `CalamineBook.open(path)`
  - DoD: can open `.xlsx` (and `.xls` if supported) without panic

- [x] RUST-011 (Done): Implement `CalamineBook.sheet_names()`
  - DoD: matches Excel sheet order

- [x] RUST-012 (Done): Implement `CalamineBook.read_cell_value(sheet, a1)`
  - DoD: returns dict matching ExcelBench cell payload contract
  - Edge cases: blank/out-of-bounds cells return `{type: "blank"}`

- [x] PY-010 (Done): Add `src/excelbench/harness/adapters/rust_calamine_adapter.py`
  - DoD: adapter passes existing `cell_values` read tests (or adds new tests)

- [x] PY-011 (Done): Register adapter in `src/excelbench/harness/adapters/__init__.py`
  - DoD: adapter appears only when `excelbench_rust` is installed

### T2: rust_xlsxwriter (write)

- [x] RUST-020 (Done): Implement `RustXlsxWriterBook.new()` and `add_sheet(name)`
  - DoD: can create workbook with multiple sheets

- [x] RUST-021 (Done): Implement `write_cell_value(sheet, a1, value_dict)` (Tier 1)
  - DoD: supports string/number/bool/date/datetime/formula/error/blank (best-effort)

- [x] RUST-022 (Done): Implement `save(path)`
  - DoD: file opens in openpyxl, values match expectations

- [x] PY-020 (Done): Add `src/excelbench/harness/adapters/rust_xlsxwriter_adapter.py`
  - DoD: write verification path in `src/excelbench/harness/runner.py` passes for Tier 1

### T3: umya-spreadsheet (read/write)

- [x] RUST-030 (Done): Implement `UmyaBook.open(path)` and `sheet_names()`

- [x] RUST-031 (Done): Implement `UmyaBook.read_cell_value(sheet, a1)`

- [x] RUST-032 (Done): Implement `UmyaBook.new()`, `add_sheet(name)`, `write_cell_value(...)`, `save(path)`

- [x] PY-030 (Done): Add `src/excelbench/harness/adapters/umya_adapter.py`

### T4: Tests + harness integration

- [x] TEST-001 (Done): Add smoke tests for optional Rust module import
  - DoD: tests skip gracefully when module is absent

- [x] TEST-002 (Done): Add fixture-based tests for Rust adapters (Tier 1)
  - DoD: at least `cell_values` read/write sanity checks

- [x] DOC-001 (Done): Document local build workflow
  - Where: `README.md` (and/or `docs/`)
  - Include: Rust toolchain, maturin commands, troubleshooting

## Commands (dev workflow)

Install deps:
```
uv sync --extra rust
```

Build and install the Rust extension into the active venv:
```
uv run maturin develop --manifest-path rust/excelbench_rust/Cargo.toml
```

Run tests:
```
uv run pytest
```

Run benchmark:
```
uv run excelbench benchmark --tests test_files --output results
```

## Known Risks / Gotchas

- Packaging split-brain (hatchling vs maturin): keep Rust as an optional, separate local build artifact until stable.
- Cross-platform builds (macOS arm64/x86_64, linux): avoid platform-specific APIs in the wrapper.
- Excel date semantics differ by library: normalize at the adapter boundary.
- Formulas: some readers return computed values only; track that limitation in adapter notes.

## Session Log (append-only)

Add a short entry each session:

Template:
```
### YYYY-MM-DD HH:MM TZ
- Worked on: RUST-XXX, PY-YYY
- Decisions: ...
- Blockers: ...
- Next: ...
```

### 2026-02-07 23:13 PST
- Created this plan/tracker.
- Next: implement RUST-001..RUST-003 skeleton.

### 02/07/2026 11:24 PM PST (via pst-timestamp)
- Worked on: RUST-001, RUST-002, RUST-003
- Notes: `rust/excelbench_rust/` skeleton builds via maturin; `build_info()` returns stable keys.
- Gotcha: `uv sync` does not preserve the locally-installed extension module; rerun `maturin develop` after syncing.
- Next: RUST-010..RUST-012 (calamine read path)

### 02/07/2026 11:29 PM PST (via pst-timestamp)
- Worked on: RUST-010, RUST-011, RUST-012, PY-010, PY-011
- Notes: Implemented `excelbench_rust.CalamineBook` behind the `calamine` cargo feature and wired `RustCalamineAdapter` into the harness.
- Validation: `uv run pytest` passed; `uv run excelbench benchmark --feature cell_values --output results_dev` produced scores.
- Next: RUST-020..RUST-022 + PY-020 (rust_xlsxwriter write path)

### 02/07/2026 11:35 PM PST (via pst-timestamp)
- Worked on: RUST-020, RUST-021, RUST-022, PY-020
- Notes: Implemented `excelbench_rust.RustXlsxWriterBook` behind the `rust_xlsxwriter` cargo feature and wired `RustXlsxWriterAdapter` into the harness.
- Validation: `uv run excelbench benchmark --feature cell_values --output results_dev2` executed the write verification path.
- Next: umya (RUST-030..RUST-032 + PY-030)

### 02/07/2026 11:36 PM PST (via pst-timestamp)
- Worked on: P0/P1 review fixes (calamine datetime mapping, sheet order, error mapping, a1 parsing util, import guards, lockfile policy)
- Notes: Addressed Codex review items and improved build_info version reporting.
- Next: umya (RUST-030..RUST-032 + PY-030)

### 02/08/2026 12:40 AM PST (via pst-timestamp)
- Worked on: calamine semantic datetime fix, umya integration, TEST-001/TEST-002
- Notes: `Data::DateTime` now maps to `date` vs `datetime` payloads using chrono + midnight check; umya backend + adapter added; rust integration tests added.
- Next: expand Rust formatting support (number formats, text formatting) beyond cell values.

## Resume Checklist (fast re-entry)

1. Confirm venv and deps: `uv sync --extra rust`
2. Build extension: `uv run maturin develop --manifest-path rust/excelbench_rust/Cargo.toml`
3. Sanity check import: `python -c "import excelbench_rust; print(excelbench_rust.build_info())"`
4. Run tests: `uv run pytest`
5. Run a small benchmark slice (optional): `uv run excelbench benchmark --feature cell_values`
