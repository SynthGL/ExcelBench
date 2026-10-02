# SheetJS CE helper (`sheetjs` adapter)

Node helper behind `src/excelbench/harness/adapters/sheetjs_adapter.py`. It drives
SheetJS Community Edition (npm package `xlsx`) through its public API only:
`XLSX.readFile`, `XLSX.utils`, and `XLSX.writeFile`.

## Install

```sh
cd tools/external-oracles/sheetjs
npm ci
```

`package.json` pins SheetJS CE **0.20.3** from the official CDN tarball
(`https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz`); `package-lock.json`
records its integrity hash. The npm registry `xlsx` package is stale (0.18.x), so
do not replace the CDN URL with a registry version. The reported version comes
from `XLSX.version` at runtime. Verified with Node 22.

## Protocol

The Python adapter runs `node sheetjs-adapter.cjs` with this directory as cwd,
writes one JSON request on stdin, and reads one JSON object from stdout. Failures
print `{"error": "sheetjs_failed", "message": ...}` and exit 1. The full contract
lives with the Python `JsonModelAdapter`.

| Operation | What the helper does |
| --- | --- |
| `describe` | Returns `library: "sheetjs"`, `version` from `XLSX.version`, capabilities `read`/`write`/`modify`, and the `unsupported_write` map below. |
| `read_model` | `XLSX.readFile(path, {cellFormula, cellNF, cellStyles, cellDates, UTC, sheetStubs})`, then reports cell values (`t`/`v`/`f`/`w`), number formats (`z`), solid fills (`s.fgColor`), `!rows` heights (`hpt`), raw `!cols` widths, `!merges`, hyperlinks (`l`), comments (`c`), and `Workbook.Names`. |
| `write_model` | Builds a workbook with `XLSX.utils.book_new` / `sheet_new` / `book_append_sheet`, replays each op on the cell objects and sheet keys, and saves with `XLSX.writeFile`. |
| `mutate` | `XLSX.readFile` on the template, `XLSX.utils.sheet_add_aoa` for each mutated cell, then `XLSX.writeFile`. This is SheetJS's real modify path: it parses the package into its object model and writes a brand-new package, so parts SheetJS does not model are not carried over. |

### Value mapping

- Formula cells report `=` plus `f`. A formula whose cached result is an error
  (`t: "e"`) reports type `error` with the error literal, per the contract.
- With `cellDates` and `UTC`, date-formatted serials arrive as UTC `Date`
  objects; midnight values report `date`, others `datetime`.
- Written dates are `t: "d"` cells with an ISO string. SheetJS converts them to
  serials and applies its default date format (`m/d/yy`) when no `z` is set.
- Error ops are written as the formulas the contract lists (`1/0`, `NA()`, ...).

### Partial support (reported, not declared unsupported)

- **Cell format, read:** with `cellStyles`, the CE XLSX reader attaches only the
  cell fill to `cell.s`. The helper reports `bg_color` (solid fills) and
  `number_format`; font, alignment, wrap, rotation, and indent keys are absent.
- **Cell format, write:** the CE writer emits only number formats in cell
  styles (`get_cell_style` builds xfs with `fontId`/`fillId`/`borderId` 0). Only
  `number_format` is written (`XLSX.utils.cell_set_number_format`); other keys are
  dropped and logged to stderr.
- **Row heights, write:** `!rows` is serialized inside `<sheetData>`, which the
  CE writer skips when the sheet has no `!ref` (no cells). Heights set on an
  otherwise empty sheet are therefore lost.
- **Page setup:** the CE XLSX reader and writer handle only `!margins`. The
  helper reports and writes `print_title_rows` through the `_xlnm.Print_Titles`
  defined name in `Workbook.Names`. Orientation, fit-to, scale, and header/footer
  have no CE API; they are omitted on read and dropped (logged to stderr) on write.
- **Named ranges:** Excel built-in `_xlnm.*` names are excluded from
  `named_ranges` because they are not user-defined ranges.
- **Comments:** one entry per cell from `cell.c[0]` (`a` author, `t` text,
  `T` threaded flag).

## Unsupported declarations

Read side (`model.unsupported`):

| Feature key | Reason |
| --- | --- |
| `cell_border` | SheetJS CE does not expose cell borders: with cellStyles the XLSX reader attaches only the fill to cell.s |
| `conditional_formats` | SheetJS CE has no conditional formatting API; the XLSX reader skips `<conditionalFormatting>` |
| `data_validations` | SheetJS CE has no data validation API; the XLSX reader skips `<dataValidations>` |
| `images` | SheetJS CE does not parse worksheet drawings, so embedded images are not exposed |
| `pivot_tables` | SheetJS CE does not parse pivot table parts |
| `freeze_panes` | SheetJS CE parses `<sheetView>` only for zoom and right-to-left; panes are not exposed |
| `tables` | SheetJS CE does not parse table parts (`xl/tables/*.xml`); only the sheet autofilter is exposed |
| `sheet_protection` | SheetJS CE's XLSX reader does not parse `<sheetProtection>`; `ws['!protect']` is populated only for XLS/XLSB/XLML inputs |
| `chart_anchors` | SheetJS CE does not parse charts or drawing anchors in worksheets |

Write side (`describe.unsupported_write`):

| Op | Reason |
| --- | --- |
| `cell_border` | SheetJS CE has no border API; the SheetJS CE writer only emits number formats in cell styles |
| `conditional_format` | SheetJS CE has no conditional formatting API; the XLSX writer emits no `<conditionalFormatting>` |
| `data_validation` | SheetJS CE has no data validation API; the XLSX writer emits no `<dataValidations>` |
| `image` | SheetJS CE cannot embed images (drawing parts are not written) |
| `pivot` | SheetJS CE cannot create pivot tables |
| `freeze` | SheetJS CE's XLSX writer emits `<sheetView>` without a `<pane>` element, so panes cannot be set |
| `table` | SheetJS CE cannot create table parts; only a sheet autofilter (`ws['!autofilter']`) is writable |
| `chart` | SheetJS CE cannot create charts |

Sheet protection is writable: `ws['!protect']` keys mirror the raw
`<sheetProtection>` attributes (true means blocked), and `password` is hashed by
SheetJS into the legacy `password` attribute.

## Smoke

From the ExcelBench repo root:

```sh
excelbench benchmark -t fixtures/excel -o /tmp/smoke-sheetjs -a sheetjs
```
