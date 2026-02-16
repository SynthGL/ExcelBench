# calamine + rust_xlsxwriter Sprint Tracker

> Historical naming note: this tracker predates the module rename. References to
> `excelbench_rust` refer to what is now imported as `wolfxl._rust`. The crate path
> `rust/excelbench_rust/` remains current.

Created: 02/14/2026 07:40 AM PST (via pst-timestamp)
Plan: `docs/plans/2026-02-14-fast-rust-excel-library.md`

## Status Key

| Symbol | Meaning |
|--------|---------|
| ` ` | Not started |
| `~` | In progress |
| `x` | Complete |
| `!` | Blocked |
| `-` | Skipped / N/A |

## Dashboard

| Sprint | Name | Tasks | Done | Blocked | Status |
|--------|------|-------|------|---------|--------|
| S0 | Foundation | 7 | 0 | 0 | Not started |
| S1 | Tier 1 Read | 6 | 0 | 0 | Not started |
| S2 | Tier 2 Read | 9 | 0 | 0 | Not started |
| S3 | Write Adapter | 10 | 0 | 0 | Not started |
| S4 | Combined + Optimization | 9 | 0 | 0 | Not started |
| S5 | Polish + Upstream | 6 | 0 | 0 | Not started |

**Overall**: 0/47 tasks complete. Target: 16/16 green R+W on Tier 0-2.

---

## Sprint 0: Foundation (1-2 sessions)

**Goal**: Fork calamine, merge styles PR, prove read speed advantage.
**Deliverable**: `CalamineStyledBook` passes cell_values + all Tier 1 formatting reads.

| ID | Task | Status | Session | Notes |
|----|------|--------|---------|-------|
| S0.1 | Fork calamine repo, merge PR #538 (styles branch), resolve conflicts | ` ` | — | PR #538 last updated 2026-01-30. Check for merge conflicts against main. |
| S0.2 | Add fork as git dependency in `Cargo.toml` (`calamine = { git = "..." }`) | ` ` | — | Use feature flag `calamine_styled` to keep optional. |
| S0.3 | Create `src/calamine_styled/mod.rs` with PyO3 bindings for style-aware reads | ` ` | — | Reuse patterns from `src/umya/mod.rs` and `src/calamine_backend.rs`. |
| S0.4 | Implement `CalamineStyledBook` PyO3 class: `open()`, `sheet_names()`, `read_cell_value()`, `read_cell_format()` | ` ` | — | `read_cell_format()` returns dict matching `CellFormat` fields. |
| S0.5 | Verify open speed < 0.5ms on small fixtures (current calamine: 0.08ms) | ` ` | — | Styles parsing will add overhead — measure delta. |
| S0.6 | Run ExcelBench fidelity on Tier 0+1 read features, record scores | ` ` | — | Expect: cell_values 3, formulas TBD, sheets 3, formatting TBD. |
| S0.7 | Microbenchmark: open+read must be faster than openpyxl (<1ms total) | ` ` | — | Use same microbenchmark script from investigation session. |

**Blockers**: None anticipated. PR #538 is the critical dependency — if it's stale, we build styles parsing ourselves.

---

## Sprint 1: Tier 1 Read Completeness (1-2 sessions)

**Goal**: 9/9 green on Tier 0+1 read.
**Deliverable**: All Tier 0 + Tier 1 features score 3/3 on read.

| ID | Task | Status | Session | Notes |
|----|------|--------|---------|-------|
| S1.1 | Implement theme color resolution (parse `xl/theme/theme1.xml`, resolve `<color theme="N" tint="0.4"/>`) | ` ` | — | PR #538 likely has partial support. Theme XML → RGB hex lookup table. |
| S1.2 | Fix cell_values edge cases: error values (#DIV/0!, #N/A, etc.), booleans, dates, shared strings | ` ` | — | Current calamine Rust backend scores 1/3 — needs error + date handling. |
| S1.3 | Expose formulas read via PyO3 (calamine already parses `<f>` elements internally) | ` ` | — | Check `calamine::DataType` for formula variant. May need `calamine::Sheets::worksheet_formula()`. |
| S1.4 | Implement dimensions read: parse `<dimension ref="A1:G20"/>` + `<col>` widths + row heights | ` ` | — | Dimensions = sheet range + col widths + row heights + default sizes. |
| S1.5 | Run full Tier 0+1 benchmark, verify all 9 features are 3/3 read | ` ` | — | Features: cell_values, formulas, multiple_sheets, text_formatting, bg_colors, number_formats, alignment, borders, dimensions. |
| S1.6 | Performance gate: full read suite must beat openpyxl (target <1ms avg) | ` ` | — | If >1ms, profile and optimize hot paths. |

**Blockers**: S1.1 (theme colors) is the highest risk — OOXML theme resolution is notoriously complex.

---

## Sprint 2: Tier 2 Read (2-3 sessions)

**Goal**: 16/16 green on Tier 0-2 read.
**Deliverable**: All 16 scored features score 3/3 on read.

| ID | Task | Status | Session | Notes |
|----|------|--------|---------|-------|
| S2.1 | Merged cells read: parse `<mergeCells><mergeCell ref="A1:C3"/>` | ` ` | — | Simple XML parse. Return list of range strings. |
| S2.2 | Conditional formatting read: parse `<conditionalFormatting>` (cell_is, color_scale, data_bar, icon_set) | ` ` | — | Complex — multiple rule types with different schemas. Map to ExcelBench CF dict format. |
| S2.3 | Data validation read: parse `<dataValidations>` (list, whole, decimal, date, textLength, custom) | ` ` | — | Extract: type, formula1/2, operator, allowBlank, showError, promptTitle, prompt. |
| S2.4 | Hyperlinks read: parse `<hyperlinks>` + resolve `_rels/sheet1.xml.rels` for external URLs | ` ` | — | Must handle: external URLs, internal cell refs, tooltip text. |
| S2.5 | Images read: parse drawing rels → `xl/drawings/drawing1.xml` → `xl/media/image1.png` | ` ` | — | **Hardest Tier 2 feature.** Need to trace: sheet→drawing rel→drawing XML→image rel→media file. Return anchor cell + image bytes. |
| S2.6 | Comments read: parse `xl/comments1.xml`, map to anchor cells | ` ` | — | Also check for `xl/threadedComments/threadedComment1.xml` (newer Excel format). |
| S2.7 | Freeze panes read: parse `<sheetViews><sheetView><pane>` for xSplit, ySplit, topLeftCell | ` ` | — | Straightforward XML parse. |
| S2.8 | Run full Tier 0-2 benchmark, all 16 features must be 3/3 read | ` ` | — | Gate for Sprint 3 start. |
| S2.9 | Performance gate: full read suite faster than openpyxl | ` ` | — | Target: <0.8ms average per feature. |

**Blockers**: S2.5 (images) is highest risk — drawing XML relationships are multi-layered.

---

## Sprint 3: Write Adapter Completeness (1-2 sessions)

**Goal**: 16/16 green on Tier 0-2 write via rust_xlsxwriter.
**Deliverable**: All 16 scored features score 3/3 on write.

**Context**: rust_xlsxwriter already has Rust API support for ALL features. The current ExcelBench
adapter (`rust_xlsxwriter_adapter.py` / `rust_xlsxwriter_backend.rs`) has stub `return` statements
for Tier 2 methods. This sprint wires up the existing APIs.

| ID | Task | Status | Session | Notes |
|----|------|--------|---------|-------|
| S3.1 | Fix cell_values write: handle error values (#DIV/0!, #N/A, #VALUE!, etc.) | ` ` | — | Currently scores 1/3. Need `worksheet.write_formula()` with error result? Or write error string. |
| S3.2 | Implement `merge_cells()` → `worksheet.merge_range()` | ` ` | — | rust_xlsxwriter has `merge_range(first_row, first_col, last_row, last_col, string, format)`. |
| S3.3 | Implement `add_conditional_format()` → map rule dict to `ConditionalFormat*` types | ` ` | — | rust_xlsxwriter has: `ConditionalFormatCell`, `ConditionalFormatColorScale`, `ConditionalFormatDataBar`, `ConditionalFormatIconSet`. |
| S3.4 | Implement `add_data_validation()` → map validation dict to `DataValidation` | ` ` | — | rust_xlsxwriter `DataValidation` has: `validate`, `criteria`, `value`, `input_message`, etc. |
| S3.5 | Implement `add_hyperlink()` → `worksheet.write_url()` | ` ` | — | Also handle `write_url_with_text()` for display text. tooltip via `set_url_tooltip()`. |
| S3.6 | Implement `add_image()` → `worksheet.insert_image()` | ` ` | — | rust_xlsxwriter `Image::new()` from bytes or file. Set position with `set_offset()`. |
| S3.7 | Implement `add_comment()` → `worksheet.write_comment()` (Note) | ` ` | — | rust_xlsxwriter `Note::new()` with text and author. Set via `worksheet.insert_note()`. |
| S3.8 | Implement `set_freeze_panes()` → `worksheet.set_freeze_panes()` | ` ` | — | `set_freeze_panes(row, col)` — straightforward. |
| S3.9 | Run full Tier 0-2 benchmark, all 16 features must be 3/3 write | ` ` | — | Gate for Sprint 4 start. |
| S3.10 | Performance gate: write must match or beat openpyxl (<2ms avg) | ` ` | — | rust_xlsxwriter already at 2.35ms vs openpyxl 2.36ms — competitive. |

**Blockers**: S3.3 (conditional formatting) is most complex due to multiple rule type mappings.

---

## Sprint 4: Combined Adapter + Optimization (1 session)

**Goal**: Single adapter presenting both backends as one library.
**Deliverable**: Production-ready adapter. 16/16 green R+W. Faster than openpyxl. On dashboard.

| ID | Task | Status | Session | Notes |
|----|------|--------|---------|-------|
| S4.1 | Create `CalamineXlsxWriterAdapter` Python class in `adapters/calamine_xlsxwriter_adapter.py` | ` ` | — | Subclass `ExcelAdapter`. Delegate read→calamine styled, write→rust_xlsxwriter. |
| S4.2 | Register in `adapters/__init__.py`, add to benchmark profiles | ` ` | — | Optional-import guard pattern. Name: "calamine+rxw" or "excelbench-rust". |
| S4.3 | Add bulk read API: `read_sheet_values()` using calamine's `Range::rows()` | ` ` | — | Batch all cells in one PyO3 call instead of cell-by-cell. |
| S4.4 | Add bulk write API: `write_sheet_values()` for throughput benchmarks | ` ` | — | Accept list of (row, col, value) tuples, write in single Rust call. |
| S4.5 | Run full fidelity benchmark: must be 16/16 green R+W | ` ` | — | Final fidelity gate. |
| S4.6 | Run performance benchmark: must beat openpyxl on all features | ` ` | — | Target: <0.8ms R, <1.5ms W average. |
| S4.7 | Generate throughput fixtures and run large-workload perf test | ` ` | — | 10k/100k rows. Measure scaling characteristics. |
| S4.8 | Regenerate dashboard, heatmap, scatter plots with new adapter | ` ` | — | `uv run excelbench html`, `uv run excelbench heatmap`, `uv run excelbench scatter`. |
| S4.9 | Update CLAUDE.md, tracker docs, library-expansion-tracker.md | ` ` | — | Final documentation pass for new adapter. |

**Blockers**: Depends on S2+S3 completion.

---

## Sprint 5: Polish + Upstream (optional, 1 session)

**Goal**: Contribute changes back, reduce fork maintenance.
**Deliverable**: Upstream PRs submitted, fork dependency minimized.

| ID | Task | Status | Session | Notes |
|----|------|--------|---------|-------|
| S5.1 | Submit PR to calamine upstream with Tier 2 parsing additions | ` ` | — | Split into focused PRs: merged_cells, CF, DV, hyperlinks, images, comments, freeze_panes. |
| S5.2 | Review calamine styles PR #538 — if merged upstream, switch from fork | ` ` | — | Monitor: https://github.com/tafia/calamine/pull/538 |
| S5.3 | Submit PR to rust_xlsxwriter for any adapter gaps discovered | ` ` | — | Likely none — rust_xlsxwriter is very complete. |
| S5.4 | Evaluate: can we drop umya-spreadsheet dependency entirely? | ` ` | — | If calamine+rxw covers all umya features, remove umya from feature flags. |
| S5.5 | Update CI to build new adapter (calamine fork as git dep) | ` ` | — | May need `cargo vendor` or GitHub Actions cache for git dep. |
| S5.6 | Final documentation pass | ` ` | — | README, CLAUDE.md, tracker, plan — all updated with final state. |

**Blockers**: None — all tasks are optional improvements.

---

## Performance Baselines (from investigation session)

Captured 02/14/2026 on small fixtures (~11 cells):

| Metric | umya | openpyxl | calamine (py) | calamine (Rust) | Target |
|--------|------|----------|---------------|-----------------|--------|
| open() | 5.31ms | 1.82ms | 0.08ms | 0.95ms | <0.5ms |
| 11 cells read | 0.05ms | 0.008ms | 0.003ms | 8.73ms* | <0.03ms |
| Per-cell read | 0.005ms | 0.0008ms | — | 0.79ms* | <0.003ms |
| write+save | 2.35ms | 2.36ms | — | — | <2ms |

*Calamine Rust backend has a bug: `worksheet_range()` re-parses per cell call.

## Fidelity Baselines

Current scores (02/14/2026, from `results/xlsx/results.json`):

| Feature | calamine R | calamine(rs) R | rxw W | Target |
|---------|-----------|---------------|-------|--------|
| cell_values | 1 | 1 | 1 | 3 R+W |
| formulas | 0 | 0 | 3 | 3 R+W |
| text_formatting | 0 | 0 | 3 | 3 R+W |
| background_colors | 0 | 0 | 3 | 3 R+W |
| number_formats | 0 | 0 | 3 | 3 R+W |
| alignment | 1 | 1 | 3 | 3 R+W |
| borders | 0 | 0 | 3 | 3 R+W |
| dimensions | 0 | 0 | 3 | 3 R+W |
| multiple_sheets | 3 | 3 | 3 | 3 R+W |
| merged_cells | 0 | 0 | 0 | 3 R+W |
| conditional_format | 0 | 0 | 0 | 3 R+W |
| data_validation | 0 | 0 | 0 | 3 R+W |
| hyperlinks | 0 | 0 | 0 | 3 R+W |
| images | 0 | 0 | 0 | 3 R+W |
| comments | 0 | 0 | 0 | 3 R+W |
| freeze_panes | 0 | 0 | 0 | 3 R+W |

**Read**: 2/16 features partially working (cell_values, alignment at 1; multiple_sheets at 3).
**Write**: 8/16 features at score 3 (Tier 0+1 complete); 8 Tier 2 features at 0 (stubbed).

---

## Session Log (append after each work session)

### Template
```
### MM/DD/YYYY — Session N
- **Sprint**: SX
- **Tasks completed**: S0.1, S0.2, ...
- **Tasks in progress**: S0.3 (~50%)
- **Blockers**: [describe any blockers]
- **Key decisions**: [any decisions made]
- **Commits**: `abc1234 message`
- **Next session**: Start with S0.3, then S0.4
```

### 02/14/2026 — Session 0 (Investigation + Planning)
- **Sprint**: Pre-S0 (investigation)
- **Tasks completed**: None (planning phase)
- **Deliverables**:
  - Root cause identified: umya 5.3ms open() is eager OOXML DOM parse
  - Microbenchmark script validated performance characteristics
  - Plan document created: `docs/plans/2026-02-14-fast-rust-excel-library.md`
  - Sprint tracker created: `docs/trackers/calamine-xlsxwriter-sprint.md`
- **Key decisions**:
  - Option B (calamine + rust_xlsxwriter) over Option A (fork umya) — 66x faster reader baseline
  - Fork calamine + merge PR #538 rather than waiting for upstream
  - Extend `excelbench_rust` crate rather than creating new crate
- **Next session**: Start S0.1 (fork calamine, merge styles PR)
