# Fast Rust Excel Library — Plan & Sprint Tracker

Created: 02/14/2026 07:30 AM PST
Status: **PLANNING**

## Problem Statement

Our pyumya adapter (umya-spreadsheet via PyO3) scores Tier A fidelity (13R/15W out of 18) but
is **2x slower than pure-Python openpyxl** on small fixtures. Microbenchmarking reveals the
bottleneck is `reader::xlsx::read()` taking 5.3ms to eagerly parse the full OOXML DOM (styles,
themes, shared strings, drawings, relationships) — versus 1.8ms for openpyxl and 0.08ms for
python-calamine. Per-cell access is fast once open (0.005ms/cell), so the issue is entirely
in the parse phase.

**Goal**: A Rust/PyO3 adapter that scores **16/16 green on Tier 0-2** (both R+W) while being
**faster than openpyxl** on both small and large workloads. Tier 3 (named_ranges, tables) is
explicitly out of scope.

## Approach Evaluation

### Option A: Fork umya-spreadsheet, strip it down
- **Pros**: Already 13R/15W, minimal adapter work, single crate
- **Cons**: 5.3ms `open()` is in the core XML parser which builds a full DOM for round-tripping.
  Stripping features doesn't help because the parser reads everything anyway. Would need to
  rewrite the reader. umya-spreadsheet is ~73k lines — large fork to maintain. 5 feature gaps
  (3 upstream: alignment indent, hyperlinks tooltip, images read).
- **Verdict**: High maintenance burden, hard to make fast without rewriting the reader.

### Option B: calamine (read) + rust_xlsxwriter (write) — "Best of Both Worlds"
- **Pros**: calamine is the fastest Rust xlsx reader (0.08ms open via python-calamine). rust_xlsxwriter
  is the most complete Rust xlsx writer (already supports merged cells, CF, DV, hyperlinks,
  images, comments, freeze panes at the API level). Both are well-maintained, focused, and fast.
- **Cons**: calamine has zero formatting read support today. Would need to extend it with
  styles parsing. Two separate libraries for one adapter.
- **Key discovery**: calamine has **open draft PR #538 "Implement styles"** (updated 2026-01-30)
  that adds exactly the structs we need: Font, Fill, Borders, Alignment, NumberFormat. This is
  in-progress upstream work we can build on or fork from.
- **Verdict**: Best performance ceiling. Clear separation of concerns. Styles PR gives us a head start.

### Option C: Build from scratch ("excelfast")
- **Pros**: Complete control, minimal dependencies, optimized for our exact feature set.
- **Cons**: Massive effort. OOXML has countless quirks (theme color resolution, shared strings,
  implicit styles, 1904 date system, etc.). Would take months.
- **Verdict**: Not worth the effort when calamine + rust_xlsxwriter exist.

## Recommendation: Option B — calamine-extended + rust_xlsxwriter

```
┌──────────────────────────────────────────────────────────┐
│                    ExcelBench Adapter                     │
│              "CalamineXlsxWriterAdapter"                  │
├────────────────────────┬─────────────────────────────────┤
│     READ (calamine)    │    WRITE (rust_xlsxwriter)      │
│                        │                                 │
│  Cell values ✓ (done)  │  Cell values ✓ (done)           │
│  Formulas ✓ (done)     │  Formulas ✓ (done)              │
│  Sheet names ✓ (done)  │  Sheet names ✓ (done)           │
│  Styles (PR #538) ←──── ─→ Formatting ✓ (done)          │
│  Merged cells (new)    │  Merged cells (new adapter)     │
│  Cond. fmt (new)       │  Cond. fmt (new adapter)        │
│  Data valid. (new)     │  Data valid. (new adapter)      │
│  Hyperlinks (new)      │  Hyperlinks (new adapter)       │
│  Images (new)          │  Images (new adapter)            │
│  Comments (new)        │  Comments (new adapter)          │
│  Freeze panes (new)    │  Freeze panes (new adapter)     │
│  Dimensions (new)      │  Dimensions ✓ (done)            │
├────────────────────────┴─────────────────────────────────┤
│                   PyO3 Boundary Layer                     │
│            (bulk APIs, minimal dict allocation)           │
└──────────────────────────────────────────────────────────┘
```

### Why this is the best path

1. **Read speed**: calamine's lazy parser + mmap gives us 0.08ms open (66x faster than umya)
2. **Write quality**: rust_xlsxwriter already has API support for ALL Tier 2 features
3. **Styles head start**: PR #538 implements Font/Fill/Borders/Alignment/NumberFormat parsing
4. **Maintained upstream**: Both libraries have active maintainers and regular releases
5. **Clean boundary**: Read and write are separate concerns — easier to test and optimize independently

### What we need to build

| Component | Work Required | Estimated Lines |
|-----------|--------------|-----------------|
| Fork calamine with styles PR #538 | Merge PR, stabilize API, add theme color resolution | ~500 |
| Add Tier 2 read features to calamine fork | Parse merged cells, CF, DV, hyperlinks, images, comments, freeze panes from sheet XML | ~1,500 |
| Extend rust_xlsxwriter PyO3 adapter | Wire up merge_cells, CF, DV, hyperlinks, images, comments, freeze panes (APIs exist) | ~800 |
| New combined adapter in Python | Dual-library adapter class, format conversion | ~300 |
| PyO3 bindings for new calamine features | Expose styles + Tier 2 data to Python | ~600 |
| Tests | ExcelBench already provides the test suite — run benchmarks | ~0 (existing) |
| **Total** | | **~3,700 lines** |

## Feature Matrix — Current vs Target

### Tier 0 (3 features) — must be 3/3 both R+W

| Feature | calamine R | rxw W | Target |
|---------|-----------|-------|--------|
| cell_values | ✓ (1/3 today, fixable) | ✓ (1/3, fixable) | 3/3 R+W |
| formulas | ✗ (0/3) | ✓ (3/3) | 3/3 R+W |
| multiple_sheets | ✓ (3/3) | ✓ (3/3) | 3/3 R+W |

### Tier 1 (6 features) — must be 3/3 both R+W

| Feature | calamine R | rxw W | Work Needed |
|---------|-----------|-------|-------------|
| text_formatting | ✗ | ✓ (3/3) | Parse styles.xml fonts → CellFormat |
| background_colors | ✗ | ✓ (3/3) | Parse styles.xml fills → CellFormat |
| number_formats | ✗ | ✓ (3/3) | Parse styles.xml numFmts → CellFormat |
| alignment | ✗ | ✓ (3/3) | Parse styles.xml alignment → CellFormat |
| borders | ✗ | ✓ (3/3) | Parse styles.xml borders → BorderInfo |
| dimensions | ✗ | ✓ (3/3) | Parse `<dimension>` + `<col>` / row spans |

### Tier 2 (7 features) — must be 3/3 both R+W

| Feature | calamine R | rxw W | Work Needed |
|---------|-----------|-------|-------------|
| merged_cells | ✗ | ✗ (API exists) | Parse `<mergeCells>` + wire adapter |
| conditional_formatting | ✗ | ✗ (API exists) | Parse `<conditionalFormatting>` + wire adapter |
| data_validation | ✗ | ✗ (API exists) | Parse `<dataValidations>` + wire adapter |
| hyperlinks | ✗ | ✗ (API exists) | Parse `<hyperlinks>` + rels + wire adapter |
| images | ✗ | ✗ (API exists) | Parse drawings + rels + wire adapter |
| comments | ✗ | ✗ (API exists) | Parse comments XML + wire adapter |
| freeze_panes | ✗ | ✗ (API exists) | Parse `<sheetView><pane>` + wire adapter |

## Sprint Plan

### Sprint 0: Foundation (1-2 sessions)

**Goal**: Set up the forked calamine, get styles PR working, prove read speed advantage.

- [ ] **S0.1** — Fork calamine, merge PR #538 (styles branch), resolve conflicts
- [ ] **S0.2** — Build fork as git dependency in `Cargo.toml` (not crates.io)
- [ ] **S0.3** — Add PyO3 bindings for style-aware cell reads in `excelbench_rust`
- [ ] **S0.4** — Create `CalamineStyledBook` PyO3 class with `open()`, `sheet_names()`, `read_cell_value()`, `read_cell_format()`
- [ ] **S0.5** — Verify open speed stays under 0.5ms (current: 0.08ms for values-only)
- [ ] **S0.6** — Run ExcelBench on Tier 0+1 read features, validate scores
- [ ] **S0.7** — Benchmark: open+read must be faster than openpyxl (target: <1ms total)

**Deliverable**: CalamineStyledBook passes cell_values + all Tier 1 formatting reads (6 features).

### Sprint 1: Tier 1 Read Completeness (1-2 sessions)

**Goal**: 9/9 green on Tier 0+1 read.

- [ ] **S1.1** — Implement theme color resolution (calamine PR #538 likely missing this)
  - Parse `xl/theme/theme1.xml` for color scheme
  - Resolve `<color theme="N" tint="0.4"/>` references to hex RGB
- [ ] **S1.2** — Fix cell_values edge cases (error values, booleans, dates)
- [ ] **S1.3** — Fix formulas read (calamine already parses formulas — expose via PyO3)
- [ ] **S1.4** — Implement dimensions read (parse `<dimension ref="A1:G20"/>` + column widths from `<col>`)
- [ ] **S1.5** — Run full Tier 0+1 benchmark, all 9 features must be 3/3 read
- [ ] **S1.6** — Performance gate: read must beat openpyxl on all 9 features

**Deliverable**: 9/9 green Tier 0+1 read, faster than openpyxl.

### Sprint 2: Tier 2 Read (2-3 sessions)

**Goal**: 16/16 green on Tier 0-2 read.

- [ ] **S2.1** — Merged cells read: parse `<mergeCells><mergeCell ref="A1:C3"/></mergeCells>`
- [ ] **S2.2** — Conditional formatting read: parse `<conditionalFormatting>` rules
  - Handle cell_is, color_scale, data_bar, icon_set rule types
  - Map formula references and operator strings
- [ ] **S2.3** — Data validation read: parse `<dataValidations>` elements
  - Handle list, whole, decimal, date, textLength, custom types
  - Extract formula1/formula2, operator, allowBlank, showError
- [ ] **S2.4** — Hyperlinks read: parse `<hyperlinks>` + resolve `_rels/sheet1.xml.rels`
  - Handle external URLs, internal cell refs, tooltip text
- [ ] **S2.5** — Images read: parse drawing relationships + extract image data
  - Resolve `xl/drawings/drawing1.xml` → `xl/media/image1.png`
  - Return anchor cell + image bytes/path
- [ ] **S2.6** — Comments read: parse `xl/comments1.xml`
  - Map to anchor cells, extract author + text
- [ ] **S2.7** — Freeze panes read: parse `<sheetViews><sheetView><pane>` element
  - Extract xSplit, ySplit, topLeftCell, activePane
- [ ] **S2.8** — Run full Tier 0-2 benchmark, all 16 features must be 3/3 read
- [ ] **S2.9** — Performance gate: full read suite faster than openpyxl

**Deliverable**: 16/16 green Tier 0-2 read.

### Sprint 3: Write Adapter Completeness (1-2 sessions)

**Goal**: 16/16 green on Tier 0-2 write via rust_xlsxwriter.

**Note**: rust_xlsxwriter already supports ALL these features at the Rust API level. The current
adapter just has stub `return` statements. This sprint wires up the existing APIs.

- [ ] **S3.1** — Fix cell_values write (currently 1/3 — error values not handled correctly)
- [ ] **S3.2** — Implement `merge_cells()` — call `worksheet.merge_range()`
- [ ] **S3.3** — Implement `add_conditional_format()` — map rule dict to `ConditionalFormat*` types
- [ ] **S3.4** — Implement `add_data_validation()` — map validation dict to `DataValidation`
- [ ] **S3.5** — Implement `add_hyperlink()` — call `worksheet.write_url()`
- [ ] **S3.6** — Implement `add_image()` — call `worksheet.insert_image()`
- [ ] **S3.7** — Implement `add_comment()` — call `worksheet.write_comment()` (Note)
- [ ] **S3.8** — Implement `set_freeze_panes()` — call `worksheet.set_freeze_panes()`
- [ ] **S3.9** — Run full Tier 0-2 benchmark, all 16 features must be 3/3 write
- [ ] **S3.10** — Performance gate: write must match or beat openpyxl

**Deliverable**: 16/16 green Tier 0-2 write.

### Sprint 4: Combined Adapter + Optimization (1 session)

**Goal**: Single adapter that presents as one library, performance-optimized.

- [ ] **S4.1** — Create `CalamineXlsxWriterAdapter` Python class combining both backends
- [ ] **S4.2** — Register in adapter `__init__.py`, add to benchmark profiles
- [ ] **S4.3** — Add bulk read API (`read_sheet_values`) using calamine's `Range::rows()`
- [ ] **S4.4** — Add bulk write API (`write_sheet_values`) for throughput benchmarks
- [ ] **S4.5** — Run full fidelity benchmark: must be 16/16 green R+W
- [ ] **S4.6** — Run performance benchmark: must beat openpyxl on all features
- [ ] **S4.7** — Generate throughput fixtures and run large-workload perf test
- [ ] **S4.8** — Regenerate dashboard, heatmap, scatter plots with new adapter
- [ ] **S4.9** — Update CLAUDE.md, tracker docs, and library-expansion-tracker.md

**Deliverable**: Production-ready adapter. 16/16 green. Faster than openpyxl. On dashboard.

### Sprint 5: Polish + Upstream (optional, 1 session)

**Goal**: Contribute back, reduce maintenance burden.

- [ ] **S5.1** — Submit PR to calamine upstream with our Tier 2 parsing additions
- [ ] **S5.2** — Review calamine styles PR #538 progress — if merged, switch from fork to upstream
- [ ] **S5.3** — Submit PR to rust_xlsxwriter for any missing features discovered
- [ ] **S5.4** — Evaluate: can we drop umya-spreadsheet dependency entirely?
- [ ] **S5.5** — Update CI to build new adapter (may need calamine fork as git dep)
- [ ] **S5.6** — Final documentation pass

**Deliverable**: Reduced fork maintenance. Upstream contributions.

## Key Technical Decisions

### Decision 1: Fork calamine vs wait for PR #538
**Choice**: Fork calamine, merge PR #538 into our fork, extend from there.
**Rationale**: PR #538 has been open since 2024, updated Jan 2026, but still draft. We can't
wait for upstream — fork now, contribute back later. If upstream merges, we switch to the
release and drop our fork.

### Decision 2: One combined adapter vs two separate adapters
**Choice**: One combined `CalamineXlsxWriterAdapter` that uses calamine for read and
rust_xlsxwriter for write.
**Rationale**: From ExcelBench's perspective, it looks like one library. The dual backend is an
implementation detail. Users comparing libraries see one entry in the benchmark.

### Decision 3: Separate crate vs extend excelbench_rust
**Choice**: Extend `excelbench_rust` with new feature-gated backends (`calamine_styled`, etc.).
**Rationale**: Reuse existing PyO3 infrastructure, util.rs, build system. Feature flags keep
compilation optional. No new crate management overhead.

### Decision 4: Performance target
**Choice**: Open+read < 1ms for small fixtures (current openpyxl: 1.82ms). Write+save < 2ms.
**Rationale**: calamine's reader is already 0.08ms. Even with styles parsing overhead, we
should stay well under 1ms. Write via rust_xlsxwriter is already 2.35ms — competitive with
openpyxl's 2.36ms, with room to optimize.

## Risk Register

| Risk | Impact | Likelihood | Mitigation |
|------|--------|------------|------------|
| calamine styles PR #538 has bugs | Medium | Medium | We're forking, so we fix them ourselves |
| Theme color resolution is complex | Medium | High | Start with direct RGB colors, add theme resolution incrementally |
| rust_xlsxwriter CF/DV API doesn't match ExcelBench expectations | Medium | Medium | Test early in Sprint 3, adapt format conversion |
| Images read from calamine fork is complex (drawings XML) | High | High | Images is hardest Tier 2 feature — plan 1 full session |
| Performance regression from styles parsing | Low | Low | calamine's lazy parsing means styles only parsed when accessed |
| Fork diverges too far from upstream calamine | Medium | Low | Keep changes minimal, submit PRs back |

## Performance Targets

| Metric | Current (umya) | Current (openpyxl) | Target (new) |
|--------|---------------|-------------------|--------------|
| Open (small file) | 5.31ms | 1.82ms | **<0.5ms** |
| 11 cells read | 0.05ms | 0.008ms | **<0.03ms** |
| Per-cell read | 0.005ms | 0.0008ms | **<0.003ms** |
| Write+save (11 cells) | 2.35ms | 2.36ms | **<2ms** |
| Benchmark avg (per feature) | 2.61ms R / 2.66ms W | 1.35ms R / 1.86ms W | **<0.8ms R / <1.5ms W** |

## File Locations

| File | Purpose |
|------|---------|
| `docs/plans/2026-02-14-fast-rust-excel-library.md` | This plan (source of truth) |
| `docs/trackers/calamine-xlsxwriter-sprint.md` | Sprint tracker (granular progress) |
| `rust/excelbench_rust/Cargo.toml` | Add calamine fork as git dependency |
| `rust/excelbench_rust/src/calamine_styled/` | New module for styled calamine backend |
| `rust/excelbench_rust/src/rust_xlsxwriter_backend.rs` | Extend with Tier 2 write support |
| `src/excelbench/harness/adapters/calamine_xlsxwriter_adapter.py` | New combined adapter |

## References

- [calamine](https://github.com/tafia/calamine) — Fast Rust xlsx reader
- [calamine PR #538](https://github.com/tafia/calamine/pull/538) — Draft styles support
- [rust_xlsxwriter](https://github.com/jmcnamara/rust_xlsxwriter) — Rust xlsx writer
- [fastxlsx](https://github.com/shuangluoxss/fastxlsx) — Prior art for calamine+rxw combo (no formatting)
- [umya-spreadsheet](https://lib.rs/crates/umya-spreadsheet) — Current R+W library (slow reader)
- [OOXML spec](https://ecma-international.org/publications-and-standards/standards/ecma-376/) — XML schemas for xlsx
