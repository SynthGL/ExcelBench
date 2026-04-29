# Excelize Adapter Design

## Goal

Add `Excelize` as a cross-language context adapter in ExcelBench.

Excelize is valuable because it represents a modern Go spreadsheet stack with strong relevance for backend and systems developers.

## Existing scaffold

Current helper assets:

- `tools/external-oracles/excelize/main.go`
- `tools/external-oracles/excelize/main_test.go`
- `tools/external-oracles/excelize/go.mod`
- `tools/external-oracles/excelize/README.md`

Run commands:

```bash
cd tools/external-oracles/excelize
go test ./...
go run . < request.json
```

The helper currently supports:

- `write_fixture`
- `read_metadata`

Supported write payload keys already include:

- `cells`
- `columns`
- `tables`
- `conditional_formats`
- `charts`
- `pivots`
- `slicers`
- `pictures`

This is a strong base for a write-heavy adapter and a metadata-aware structural adapter.

## Recommended adapter shape

### Public classification

- Language: `go`
- Caps: `R+W` only if stable value/formula reads are implemented in benchmark form
- Fallback classification: `W` if read support stays metadata-only for too long
- Modify: `No`

### Initial scope

P0 should target:

- Tier 0 write: `cell_values`, `formulas`, `multiple_sheets`
- Tier 1 write: `alignment`, `background_colors`, `dimensions`, `number_formats`, `text_formatting`
- Tier 2 write where helper support already exists: `conditional_formatting`, `images`, `freeze_panes`, `merged_cells`
- Pivot-related generation should be treated carefully: good for context, but not scored unless the verification path is fair and stable

Read scope for P0 should be conservative:

- `cell_values`
- `multiple_sheets`
- metadata-backed structural features only where the helper already reports them consistently

## Why this scope

- The Excelize helper already expresses a broad write payload.
- It also reports metadata counts for pivots, slicers, and conditional formatting.
- That makes it a good candidate for an honest write-capability benchmark plus limited read coverage.

## Implementation approach

### Phase 1 - helper-backed adapter

Create a Python adapter that uses subprocess calls to the existing Go helper.

Suggested new files:

- `src/excelbench/harness/adapters/excelize_adapter.py`
- optionally a shared helper module if subprocess patterns overlap with POI

Suggested existing files to update:

- `src/excelbench/harness/adapters/__init__.py`

### Phase 2 - benchmark contract mapping

Write path:

1. translate benchmark workbook spec to the Excelize helper JSON format
2. generate `.xlsx` through `go run .`
3. reuse current verification/oracle path

Read path:

1. use helper `read_metadata` for structural features
2. add direct value/formula read support only if it can be surfaced cleanly and reproducibly
3. mark unsupported features explicitly where necessary

## Capability boundaries to document

Before public inclusion, decide explicitly how to treat:

- pivots
- slicers
- charts
- pictures/images
- table metadata
- conditional formatting richness

These are good Excelize strengths, but they need fair scoring criteria. If the helper only exposes counts and package presence, that belongs in a context/capability section, not in a full-fidelity read claim.

## Testing plan

1. Adapter smoke test for missing Go toolchain or helper path.
2. Small Tier 0 write-path integration test.
3. One advanced-feature integration test using current helper-supported payloads.
4. Dated snapshot run in a separate cross-language context output path.

## Public reporting plan

- Add Excelize to a `Cross-Language Context` section first.
- Do not mix it into the Python hero table.
- Use it to answer the ecosystem-position question, not the Python migration question.

## Decision log

- Chose helper-backed adapter over direct Go embedding because the helper already has tests and a well-defined JSON contract.
- Chose write-first rollout because the helper is already strong at fixture generation and structural emit.

## Exact next implementation targets

1. `src/excelbench/harness/adapters/excelize_adapter.py`
2. `src/excelbench/harness/adapters/__init__.py`
3. adapter smoke/integration tests under the existing adapter test area
4. a dated cross-language snapshot output path, separate from the Python hero snapshot
