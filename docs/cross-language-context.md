# Cross-Language Context

ExcelBench is still a Python-first benchmark project. The main public question is:

> What should a Python team use instead of `openpyxl`?

The cross-language context snapshot answers a different question:

> How does WolfXL compare to mature spreadsheet tooling outside Python?

## Current Snapshot

- Artifact: [`results-cross-language/README.md`](../results-cross-language/README.md)
- Generated: 2026-04-29 UTC
- Scope: write-path ecosystem context only

## Current Takeaways

- `apache-poi` is now a strong cross-language reference point in this harness with `18/18` green features across the scored write surfaces in the current snapshot.
- `excelize` is now also at `18/18` green features across the scored write surfaces in the current snapshot.
- `pivot_tables` remain out of scope in the scored macOS cross-language snapshot because the shipped fixture does not currently contain pivot OOXML, but a separate pivot capability artifact exists.

## Three-Lane Framing

Use the ExcelBench outputs as three separate lanes, not one mixed claim set.

1. **Python replacement lane**
The main `results-release-*` artifacts answer the migration question for Python teams.

2. **Cross-language context lane**
The `results-cross-language/` artifacts answer the ecosystem-positioning question for technical readers who know POI or Excelize.

3. **Pivot capability lane**
The `results-cross-language-pivots/` artifact answers whether the helpers can detect or emit pivot-bearing workbooks even though the main scored pivot fixture is not valid on macOS.

## Why this matters

- It shows WolfXL is not only competitive inside Python. It also sits well against mature spreadsheet tooling in Java and Go.
- It gives a credibility layer for technical readers who know Apache POI or Excelize and want broader ecosystem context.
- It keeps the Python replacement story clean because the cross-language results live in a separate lane.

## How to cite it

- Say `cross-language context snapshot`, not `primary benchmark result`.
- Cite the artifact path and generation date.
- Keep the Python hero table and the cross-language table separate in launch copy.

## Links

- [Cross-language README](../results-cross-language/README.md)
- [Cross-language context note](../results-cross-language/CONTEXT.md)
- [Pivot capability artifact](../results-cross-language-pivots/README.md)
- [Cross-language comparison strategy](trackers/cross-language-comparison-strategy.md)
