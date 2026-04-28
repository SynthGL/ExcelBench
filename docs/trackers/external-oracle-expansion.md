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
- The helper catalog can point at repo-local helpers when called with
  `external_oracle_catalog(repo_root=...)`.
- The helper catalog currently reserves entrypoints or source helpers for:
  - `tools/external-oracles/excelize` (`go run .`)
  - `tools/external-oracles/libreoffice/libreoffice_oracle.py`
  - `excelbench-poi-oracle`
  - `excelbench-closedxml-oracle`

## Excelize helper

Implemented: `tools/external-oracles/excelize`

Supported operations:

- `write_fixture`: writes an `.xlsx` workbook from the JSON subprocess request.
- `read_metadata`: opens an `.xlsx` workbook and reports sheet-level counts for
  tables, pivots, slicers, and conditional-formatting ranges.

Supported write payload keys:

- `sheets`
- `cells`
- `columns`
- `tables`
- `conditional_formats`
- `charts`
- `pivots`
- `slicers`
- `pictures`

Current smoke coverage:

- Go unit smoke checks that a single request emits table, pivot, pivot cache,
  slicer, slicer cache, chart, drawing, and image-related workbook parts.
- Python integration smoke runs the same helper through
  `run_external_oracle()` when Go is available.

## LibreOffice helper

Implemented: `tools/external-oracles/libreoffice/libreoffice_oracle.py`

Supported operations:

- `open_save_validate`: opens an input workbook and saves it back through
  LibreOffice's Calc Office Open XML export filter.
- `render_validate` / `render_pdf`: opens an input workbook and exports it to
  PDF through `calc_pdf_Export`.

Current smoke coverage:

- Python integration smoke creates a simple workbook and asks the helper to
  render PDF. Missing LibreOffice returns a structured skip.
- Manual Excelize truth pass, 2026-04-28: the Excelize-generated workbook with
  table, pivot cache, pivot table, slicer, chart, drawing, and picture parts
  opened in openpyxl with the expected unsupported-extension warning, read
  values correctly through WolfXL, and rendered to PDF through LibreOffice
  without stderr.

## Candidate tools

| Tool | Runtime | Initial role | Status |
|---|---|---|---|
| Excelize | Go | Generate xlsx fixtures for pivots, slicers, charts, conditional formatting, tables, rich formatting, images, and streaming paths. | Initial helper implemented. |
| LibreOffice Calc | CLI / UNO | Open/save/render validator for corruption, repair, and visual/export smoke checks. | Initial helper implemented. |
| Apache POI | Java | Generate and inspect OOXML fixtures with a mature usermodel plus documented chart/pivot limits. | P1 after contract settles. |
| ClosedXML | .NET | Generate high-level table, pivot, conditional-formatting, and rich-cell fixtures. | P1 after contract settles. |
| NPOI | .NET | POI-like .NET comparison if ClosedXML/POI leave .NET-specific gaps. | P2 research. |
| SheetJS CE | JavaScript | Broad-format value/formula sanity checks; advanced styling/charts/pivots appear better suited to Pro. | P2 limited scope. |

## First fixture pack

The first external oracle pack should stay small and high-signal:

1. Excelize pivot cache + pivot table with saved records. **Scaffolded.**
2. Excelize slicer attached to a table or pivot table. **Table slicer scaffolded.**
3. Excelize chart with data labels, point colors, and alt text. **Basic chart scaffolded.**
4. Excelize conditional formatting with icon sets, color scales, and data bars. **Scaffolded.**
5. Excelize drawing/image anchors with one-cell/two-cell positions. **Basic picture scaffolded.**
6. LibreOffice open/save smoke for the same outputs. **Helper scaffolded.**
7. LibreOffice PDF/export smoke where visual corruption would be obvious. **Helper scaffolded.**

## Promotion gates

- Helper command is optional and skips cleanly when missing.
- JSON schema is deterministic enough for tests.
- Generated workbooks open in Excel/LibreOffice without repair.
- WolfXL read/modify/save behavior is explicitly classified as pass, fix-now,
  documented defer, or out of scope.
- Stable cases graduate into checked-in fixtures only after a manual truth pass.
