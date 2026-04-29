# Cross-Language Comparison Strategy

## Goal

Extend ExcelBench beyond Python-only comparisons without weakening the core public story.

The repo should answer two different questions clearly:

1. **Python replacement question**: what should a Python team use instead of `openpyxl`?
2. **Ecosystem position question**: how does WolfXL compare to mature spreadsheet libraries in other languages?

## Comparison Tiers

### Tier 1 - Primary public comparison

This is the default README and launch-post comparison set.

- `openpyxl`
- `xlsxwriter`
- `xlsxwriter-constmem`
- `python-calamine`
- `pandas`
- `pyexcel` / `tablib` / `pylightxl` as lower-fidelity reference points

Use this tier when the audience is choosing a Python library.

### Tier 2 - Cross-language credibility comparison

This tier is for ecosystem context, research posts, and deeper benchmark reports.

High-priority additions:

- `Apache POI`
- `ClosedXML`
- `Excelize`
- `ExcelJS`

Lower-priority additions:

- `NPOI`
- `EPPlus` if licensing constraints are addressed explicitly

Use this tier to show where WolfXL sits relative to mature spreadsheet tooling outside Python.

## Public Messaging Rules

- Lead with Python alternatives in the main README and launch copy.
- Put cross-language results in a separate section or separate report.
- Do not mix Python drop-in comparisons and cross-language ecosystem rankings in one hero table.
- When a library is not a realistic substitute for Python users, say so directly.

## Why This Split Exists

- Python users care first about migration cost and workflow fit.
- Cross-language libraries are valuable for credibility, not for drop-in replacement decisions.
- A single giant scoreboard makes the main value proposition harder to understand.

## Candidate Priority

### P0

- `Apache POI`
- `Excelize`

### P1

- `ClosedXML`
- `ExcelJS`

### P2

- `NPOI`

## Acceptance Criteria For New Cross-Language Adapters

Before adding a new library to public comparison tables:

1. Adapter is reproducible from source or documented package install.
2. Fixture generation and verification path are documented.
3. Capability boundaries are explicit: read, write, modify, or preserve-only.
4. Platform/runtime caveats are documented.
5. Results are generated in the same dated snapshot format as Python adapters.

## Suggested Rollout

1. Add `Apache POI` and `Excelize` first.
2. Publish them in a secondary "cross-language context" section.
3. Keep the main benchmark hero table Python-first.
4. Promote any cross-language table to primary only if the user story demands it.
