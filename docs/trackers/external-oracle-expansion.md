# External Oracle Expansion Tracker

Date: 2026-04-28
Status: scaffold started; no public benchmark claims yet.

## Goal

Use open-source spreadsheet tools outside the Python/openpyxl ecosystem to
generate and validate OOXML fixtures before the next WolfXL release. These
oracles are meant to find edge cases that openpyxl does not construct deeply,
not to replace the existing ExcelBench adapter matrix.

## Current scaffold

- `src/excelbench/harness/external_oracles.py` defines the subprocess contract.
- External helpers receive a JSON request on stdin and return JSON diagnostics
  on stdout.
- Missing helpers return a structured skip, so Go/Java/.NET/LibreOffice are not
  required for the normal test suite.
- The helper catalog currently reserves entrypoints for:
  - `excelbench-excelize-oracle`
  - `excelbench-libreoffice-oracle`
  - `excelbench-poi-oracle`
  - `excelbench-closedxml-oracle`

## Candidate tools

| Tool | Runtime | Initial role | Status |
|---|---|---|---|
| Excelize | Go | Generate xlsx fixtures for pivots, slicers, charts, conditional formatting, tables, rich formatting, images, and streaming paths. | P0 next implementation target. |
| LibreOffice Calc | CLI / UNO | Open/save/render validator for corruption, repair, and visual/export smoke checks. | P0 next implementation target. |
| Apache POI | Java | Generate and inspect OOXML fixtures with a mature usermodel plus documented chart/pivot limits. | P1 after contract settles. |
| ClosedXML | .NET | Generate high-level table, pivot, conditional-formatting, and rich-cell fixtures. | P1 after contract settles. |
| NPOI | .NET | POI-like .NET comparison if ClosedXML/POI leave .NET-specific gaps. | P2 research. |
| SheetJS CE | JavaScript | Broad-format value/formula sanity checks; advanced styling/charts/pivots appear better suited to Pro. | P2 limited scope. |

## First fixture pack

The first external oracle pack should stay small and high-signal:

1. Excelize pivot cache + pivot table with saved records.
2. Excelize slicer attached to a table or pivot table.
3. Excelize chart with data labels, point colors, and alt text.
4. Excelize conditional formatting with icon sets, color scales, and data bars.
5. Excelize drawing/image anchors with one-cell/two-cell positions.
6. LibreOffice open/save smoke for the same outputs.
7. LibreOffice PDF/export smoke where visual corruption would be obvious.

## Promotion gates

- Helper command is optional and skips cleanly when missing.
- JSON schema is deterministic enough for tests.
- Generated workbooks open in Excel/LibreOffice without repair.
- WolfXL read/modify/save behavior is explicitly classified as pass, fix-now,
  documented defer, or out of scope.
- Stable cases graduate into checked-in fixtures only after a manual truth pass.

