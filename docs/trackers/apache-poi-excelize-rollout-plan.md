# Apache POI + Excelize Rollout Plan

## Objective

Promote `Apache POI` and `Excelize` from external-oracle helpers into the first public cross-language context tier for ExcelBench.

This is a secondary comparison track. It should strengthen ecosystem credibility without replacing the Python-first benchmark story.

## Why these two first

- `Apache POI` is the strongest mature Java spreadsheet reference point.
- `Excelize` is a strong Go spreadsheet library with good modern developer relevance.
- Both already exist in the external-oracle scaffold, which reduces integration risk.

## Current starting point

Both helpers already exist under the external-oracle system:

- `tools/external-oracles/apache-poi`
- `tools/external-oracles/excelize`

Both already support fixture generation and metadata inspection, and both have smoke coverage documented in `docs/trackers/external-oracle-expansion.md`.

## Phase 1 - Capability audit

Goal: define what each library can be scored on fairly.

### Apache POI

1. Confirm read vs write scope for the adapter surface.
2. Enumerate unsupported or weak features explicitly.
3. Decide whether modify/preserve semantics should be marked `No` or `Preserve-only`.
4. Capture runtime/bootstrap steps for local and CI use.

### Excelize

1. Confirm read vs write scope for the adapter surface.
2. Confirm which advanced features are realistic to score vs generate-only.
3. Decide whether streaming APIs belong in a separate variant adapter later.
4. Capture Go toolchain and invocation assumptions.

## Phase 2 - Adapter contract design

Goal: map both libraries onto the same comparison model used by Python adapters.

Deliverables:

1. Adapter design note for `Apache POI`
2. Adapter design note for `Excelize`
3. Feature-scope matrix with statuses:
   - score normally
   - score with caveat
   - unsupported
   - preserve-only

## Phase 3 - Minimal benchmark adapter

Goal: get each library into a dated snapshot with honest capability boundaries.

### Apache POI

P0 milestone:

1. Implement write-path adapter for Tier 0 and Tier 1.
2. Implement read-path adapter for Tier 0 and whichever metadata paths are stable.
3. Run against canonical fixtures.
4. Render results into a dedicated dated snapshot.

### Excelize

P0 milestone:

1. Implement write-path adapter for Tier 0 and supported Tier 1/Tier 2 features.
2. Implement read-path adapter for cell values, sheets, and any stable metadata.
3. Run against canonical fixtures.
4. Render results into a dedicated dated snapshot.

## Phase 4 - Public reporting

Goal: publish results without muddying the Python-first README story.

Rules:

1. Keep the current hero table Python-first.
2. Add a separate "Cross-Language Context" section.
3. Label POI/Excelize results as ecosystem context, not Python alternatives.
4. Use a separate dated snapshot or subsection if needed.

## Acceptance criteria

Before either library appears in the main public repo README as an active comparison:

1. Adapter runs reproducibly from documented commands.
2. Capability boundaries are documented feature-by-feature.
3. Results are dated and traceable to raw artifacts.
4. At least one smoke test or integration test covers the adapter path.
5. README wording makes clear that the comparison is cross-language context.

## Recommended command path

Start from the existing helper-backed scaffold, not a greenfield adapter.

1. Reuse fixture-generation/inspection logic from external oracles.
2. Wrap the stable operations in ExcelBench adapter interfaces.
3. Keep bootstrap scripts per runtime (`go`, `java`) under `tools/external-oracles/`.

## Suggested order

1. `Apache POI`
2. `Excelize`
3. `ClosedXML`
4. `ExcelJS`

## Success condition

ExcelBench can show:

- a clear Python-first benchmark story for migration decisions
- a separate cross-language context view showing WolfXL relative to mature spreadsheet libraries in Java, Go, .NET, and Node
