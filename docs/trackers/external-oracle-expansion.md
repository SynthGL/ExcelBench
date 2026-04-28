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
- `src/excelbench/harness/external_fixture_specs/` defines tool-specific
  fixture specifications.
- `src/excelbench/harness/external_fixture_pack.py` orchestrates the local
  fixture pack and writes a `manifest.json` under `results_dev_external/`.
- `src/excelbench/harness/external_wolfxl_validation.py` validates generated
  fixtures through WolfXL read + in-place modify-save part preservation.
- `scripts/generate_external_oracle_fixtures.py` regenerates the local fixture
  pack.
- `scripts/validate_external_oracle_fixtures_with_wolfxl.py` runs the WolfXL
  preservation check and writes `wolfxl-validation.json`.
- External helpers receive a JSON request on stdin and return JSON diagnostics
  on stdout.
- Missing helpers return a structured skip, so Go/Java/.NET/LibreOffice are not
  required for the normal test suite.
- The helper catalog can point at repo-local helpers when called with
  `external_oracle_catalog(repo_root=...)`.
- The helper catalog currently reserves entrypoints or source helpers for:
  - `tools/external-oracles/excelize` (`go run .`)
  - `tools/external-oracles/libreoffice/libreoffice_oracle.py`
  - `tools/external-oracles/closedxml` (`dotnet run --project ...`)
  - `tools/external-oracles/npoi` (`dotnet run --project ...`)
  - `excelbench-poi-oracle`

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
- WolfXL validator, 2026-04-28: both generated fixture-pack workbooks passed
  read + in-place modify-save preservation after WolfXL commit `634be84`.

## ClosedXML helper

Implemented: `tools/external-oracles/closedxml`

Supported operations:

- `write_fixture`: writes sheets, cells, tables, conditional formats, and pivot
  tables, rich text, comments, and sheet protection through ClosedXML.
- `read_metadata`: inspects package parts for tables, pivot tables, pivot
  caches, comments, VML drawings, and worksheets.

Current smoke coverage:

- .NET integration smoke writes a table + pivot + conditional-format + rich
  text + comment + sheet-protection workbook through `run_external_oracle()`
  when `dotnet` is available.
- Manual ClosedXML truth pass, 2026-04-28: ClosedXML generated a workbook with
  a table, pivot cache, pivot table, and conditional-format extension records.
  WolfXL initially preserved the parts but inserted the smoke marker without
  the worksheet namespace prefix; WolfXL commit `7640a3f` fixed the patcher,
  and the same workbook now passes read + in-place modify-save preservation.
- Fixture-pack promotion, 2026-04-28: `closedxml_pivot_cf_table` is included
  when `dotnet` is available.
- Fixture-pack expansion, 2026-04-28: `closedxml_rich_comment_protection`
  covers rich shared strings, legacy comments, VML comment drawing parts, and
  sheet protection. The full four-workbook pack passes LibreOffice
  open/render validation and WolfXL read + in-place modify-save preservation.

## NPOI helper

Implemented: `tools/external-oracles/npoi`

Supported operations:

- `write_fixture`: writes sheets, cells, formulas, rich text, legacy comments,
  merged ranges, and sheet protection through NPOI.
- `read_metadata`: inspects package parts for worksheets, shared strings,
  comments, VML drawings, and calc-chain metadata.

Current smoke coverage:

- .NET integration smoke writes a formula + rich text + comment + merged range
  + sheet-protection workbook through `run_external_oracle()` when `dotnet` is
  available.
- Fixture-pack promotion, 2026-04-28: `npoi_formula_comment_merge_protection`
  is included when `dotnet` is available.
- Manual NPOI truth pass, 2026-04-28: the full five-workbook pack, including
  the NPOI formula/comment/rich-text/merge/protection case, passes LibreOffice
  open/render validation and WolfXL read + in-place modify-save preservation.

## Candidate tools

| Tool | Runtime | Initial role | Status |
|---|---|---|---|
| Excelize | Go | Generate xlsx fixtures for pivots, slicers, charts, conditional formatting, tables, rich formatting, images, and streaming paths. | Initial helper implemented. |
| LibreOffice Calc | CLI / UNO | Open/save/render validator for corruption, repair, and visual/export smoke checks. | Initial helper implemented. |
| ClosedXML | .NET | Generate high-level table, pivot, conditional-formatting, and rich-cell fixtures. | Initial helper implemented; first two fixture-pack cases passing. |
| Apache POI | Java | Generate and inspect OOXML fixtures with a mature usermodel plus documented chart/pivot limits. | P1 after contract settles. |
| NPOI | .NET | POI-like .NET comparison for formulas, comments, rich strings, merged ranges, and protection. | Initial helper implemented; first fixture-pack case passing. |
| SheetJS CE | JavaScript | Broad-format value/formula sanity checks; advanced styling/charts/pivots appear better suited to Pro. | P2 limited scope. |

## Additional oracle research

2026-04-28 local runtime inventory:

- Available locally: `java`, `javac`, `node`, `npm`, `pnpm`, `bun`, `go`,
  `dotnet`, `php`, `ruby`, and `soffice`.
- Missing locally: `mvn`, `gradle`, `composer`, `ssconvert`, and
  `libreoffice` as a direct executable name.

High-signal next candidates:

1. Apache POI. POI's XSSF API exposes pivot-table creation/inspection, tables,
   sheet protection, conditional-formatting access, comments/VML drawing
   surfaces, and shared formula metadata. It remains the best missing
   independent OOXML writer, but the helper should avoid assuming Maven/Gradle
   are installed on developer machines. Source: [POI XSSFSheet API][poi-xssf].
2. ExcelJS. The project documents XLSX read/write, styles, merged cells,
   defined names, data validations, comments, tables, rich text, conditional
   formatting, images, sheet protection, streaming I/O, formula values, shared
   formulas, and array formulas. This is a strong Node oracle for style/value
   and streaming edge cases, with pivot support still newer and limited. Source:
   [ExcelJS README][exceljs-readme].
3. xlsx-populate. This is useful for template-preservation and mutation tests:
   it is explicitly a parser/generator focused on preserving existing workbook
   features and styles, with documented support for styles, rich text, data
   validation, hyperlinks, print options, panes, and encryption. Source:
   [xlsx-populate README][xlsx-populate-readme].
4. PhpSpreadsheet. This can independently generate formulas, charts, styles,
   images, merged cells, freeze panes, and sheet protection, but it needs a
   Composer bootstrap before becoming an ergonomic local oracle. Sources:
   [PhpSpreadsheet site][phpspreadsheet-site] and
   [feature cross-reference][phpspreadsheet-features].
5. libxlsxwriter. This C writer can generate formulas, hyperlinks, formatting,
   merged cells, charts, data validation, conditional formatting, images,
   comments, macros, and large-file memory-optimized outputs. It overlaps with
   Rust writer coverage, but gives a different native package producer for XML
   shape comparison. Source: [libxlsxwriter README][libxlsxwriter-readme].
6. Gnumeric `ssconvert`. This is better as a conversion/open-save oracle than a
   fixture generator. It is not installed locally, and LibreOffice already covers
   the current conversion/render validation path. Source:
   [Gnumeric ssconvert manual][gnumeric-ssconvert].

## First fixture pack

The first external oracle pack should stay small and high-signal:

Regenerate locally:

```bash
uv run python scripts/generate_external_oracle_fixtures.py
uv run python scripts/validate_external_oracle_fixtures_with_wolfxl.py
```

Fixtures:

1. `excelize_sales_pivot_slicer_chart`: pivot cache + pivot table with saved
   records, table slicer, chart, icon set, color scale, data bar, table, and
   picture. **Implemented.**
2. `excelize_chart_points_formula_cf`: per-point chart styling, formula cell,
   data bar, and cell-rule conditional formatting. **Implemented.**
3. `closedxml_pivot_cf_table`: ClosedXML table + pivot cache/table +
   conditional-formatting package layout. **Implemented.**
4. `closedxml_rich_comment_protection`: ClosedXML rich shared strings, legacy
   comments, VML comment drawings, and sheet protection. **Implemented.**
5. `npoi_formula_comment_merge_protection`: NPOI formula cell, rich shared
   string, legacy comment, merged range, and sheet protection. **Implemented.**
6. LibreOffice open/save smoke for the same outputs. **Helper scaffolded.**
7. LibreOffice PDF/export smoke where visual corruption would be obvious. **Helper scaffolded.**

## Promotion gates

- Helper command is optional and skips cleanly when missing.
- JSON schema is deterministic enough for tests.
- Generated workbooks open in Excel/LibreOffice without repair.
- WolfXL read/modify/save behavior is explicitly classified as pass, fix-now,
  documented defer, or out of scope.
- Stable cases graduate into checked-in fixtures only after a manual truth pass.

[exceljs-readme]: https://raw.githubusercontent.com/exceljs/exceljs/master/README.md
[gnumeric-ssconvert]: https://gnome.pages.gitlab.gnome.org/gnumeric/manual/sect-files-ssconvert.html
[libxlsxwriter-readme]: https://github.com/jmcnamara/libxlsxwriter
[phpspreadsheet-features]: https://phpspreadsheet.readthedocs.io/en/stable/references/features-cross-reference/
[phpspreadsheet-site]: https://phpspreadsheet.com/
[poi-xssf]: https://poi.apache.org/apidocs/5.0/org/apache/poi/xssf/usermodel/XSSFSheet.html
[xlsx-populate-readme]: https://raw.githubusercontent.com/dtjohnson/xlsx-populate/master/README.md
