# Apache POI Adapter Design

## Goal

Add `Apache POI` as a cross-language context adapter in ExcelBench.

This adapter is not meant to answer the Python migration question. It is meant to show where WolfXL sits relative to a mature Java spreadsheet library.

## Existing scaffold

Current helper assets:

- `tools/external-oracles/apache-poi/src/PoiOracle.java`
- `tools/external-oracles/apache-poi/poi_oracle.py`
- `tools/external-oracles/apache-poi/build.sh`
- `tools/external-oracles/apache-poi/fetch_deps.py`
- `tools/external-oracles/apache-poi/README.md`

Bootstrap command:

```bash
cd tools/external-oracles/apache-poi
./build.sh
```

The helper currently supports:

- `write_fixture`
- `read_metadata`

That means the shortest path is **not** a full cell-by-cell adapter on day one. The shortest path is a write-focused adapter with selective read support where metadata inspection is already stable.

## Recommended adapter shape

### Public classification

- Language: `java`
- Caps: `R+W` only if stable read semantics are implemented for benchmark-required features
- Fallback classification: `W` if read support is too partial for a fair comparison
- Modify: `No`

### Initial scope

P0 should target:

- Tier 0 write: `cell_values`, `formulas`, `multiple_sheets`
- Tier 1 write: `alignment`, `background_colors`, `borders`, `dimensions`, `number_formats`, `text_formatting`
- Tier 2 write where the current helper already has evidence: `comments`, `hyperlinks`, `data_validation`, `merged_cells`, `freeze_panes`, `images`
- Tier 3 write: defer until workbook metadata APIs are mapped cleanly

Read scope for P0 should be conservative:

- `cell_values`
- `formulas`
- `multiple_sheets`
- selected metadata-backed checks only if they can be surfaced in a benchmark-compatible way

## Why this scope

- The existing helper already generates realistic POI workbooks.
- The helper already inspects package-level metadata.
- That makes POI a good benchmark citizen for write fidelity first.
- Forcing deep read parity immediately would slow rollout and increase ambiguity.

## Implementation approach

### Phase 1 - bridge helper into adapter runtime

Create a thin Python adapter that shells out to the existing helper instead of embedding Java logic directly.

Suggested new files:

- `src/excelbench/harness/adapters/apache_poi_adapter.py`
- optionally `src/excelbench/harness/adapters/external_process_adapter_utils.py` if shared subprocess glue is needed

Suggested existing files to update:

- `src/excelbench/harness/adapters/__init__.py`
- any adapter registry/test discovery file that enumerates public adapters

### Phase 2 - benchmark contract mapping

Map ExcelBench adapter methods onto helper-backed operations.

Write path options:

1. translate a benchmark workbook spec directly into the helper JSON contract
2. let the helper produce `.xlsx`
3. verify with the existing benchmark oracle path

Read path options:

1. open workbook via helper `read_metadata` for structural features
2. use a lightweight Java read extension only for values/formulas if needed
3. mark unsupported features honestly instead of faking read support

## Capability boundaries to document

Before public inclusion, write down feature-by-feature status:

- score normally
- score via metadata inspection only
- unsupported in current adapter
- intentionally deferred

Important examples to decide explicitly:

- charts
- pivot tables
- named ranges
- tables
- comments/VML handling
- image preservation vs image semantic readback

## Testing plan

1. Add one adapter smoke test that confirms bootstrap failure returns structured skip, not an opaque crash.
2. Add one write-path integration test for a small Tier 0 workbook.
3. Add one advanced-feature integration test matching a currently supported POI fixture.
4. Run a dated snapshot separate from the Python-first hero table first.

## Public reporting plan

- Add POI only in a `Cross-Language Context` section initially.
- Do not place POI in the primary hero table.
- Label it clearly as ecosystem context, not a Python substitute.

## Decision log

- Chose helper-backed adapter over direct Java embedding because the helper already exists, has pinned dependencies, and reduces integration risk.
- Chose write-first rollout because the scaffold already proves fixture generation and metadata inspection, while full read parity is less certain.

## Exact next implementation targets

1. `src/excelbench/harness/adapters/apache_poi_adapter.py`
2. `src/excelbench/harness/adapters/__init__.py`
3. adapter smoke/integration tests under the existing adapter test area
4. a dated cross-language snapshot output path, separate from the Python hero snapshot
