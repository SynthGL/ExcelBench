# ExcelJS Helper

This directory backs two things:

- the scored ExcelBench `exceljs` adapter
  (`src/excelbench/harness/adapters/exceljs_adapter.py`), which drives ExcelJS
  read, write, and template-mutation runs through this helper; and
- the original ExcelJS external oracle (`write_fixture`, `read_metadata`).

Both use the same JSON stdin/stdout contract as the Excelize, LibreOffice,
ClosedXML, and NPOI helpers: one JSON request on stdin, one JSON object on
stdout, exit code 0 on success. Failures print
`{"error": "exceljs_oracle_failed", "message": ...}` and exit 1.

Install the pinned dependencies (ExcelJS 4.4.0, JSZip 3.10.1) once:

```bash
npm install
npm run oracle   # node exceljs-oracle.cjs
```

## Adapter operations (`exceljs-model.cjs`)

- `describe`: library slug `exceljs`, the installed ExcelJS version, and the
  write operations ExcelJS cannot express.
- `read_model` (`input_path`): loads the workbook with `workbook.xlsx.readFile`
  and reports what the ExcelJS object model exposes: cell values and formulas,
  fonts, fills, number formats, alignment, borders, row heights, column widths,
  merges, conditional formats, data validations, hyperlinks, images, notes,
  sheet views, defined names, tables, sheet protection, page setup, and
  header/footer text.
- `write_model` (`output_path`, `payload.ops`): replays ExcelBench write calls
  through the ExcelJS API (`addWorksheet`, `cell.value`, cell styles,
  `addConditionalFormatting`, `dataValidations.add`, `addImage`, `cell.note`,
  `worksheet.views`, `definedNames.add`, `addTable`, `protect`, `pageSetup`,
  `headerFooter`) and saves with `workbook.xlsx.writeFile`.
- `mutate` (`input_path`, `output_path`, `payload.mutations`): loads an existing
  workbook, sets each cell value, and writes a new package.

Declared unsupported (ExcelJS 4.4.0 has no API for them):

- pivot tables: not loaded on read, no creation API;
- charts: not loaded on read, no creation API.

Known ExcelJS 4.4.0 behavior that surfaces as failed test cases rather than
being worked around:

- Loading some workbooks with drawings or tables throws (for example
  `Cannot read properties of undefined (reading 'anchors')`); `read_model` and
  `mutate` report that error.
- Defined names lose their sheet scope (`localSheetId`), and `definedNames.add`
  cannot write a sheet-scoped name.
- Hyperlink tooltips and location-only (internal) hyperlinks are dropped on
  read; internal hyperlinks are written with both a relationship and a location.
- Conditional formatting rules have no `stopIfTrue` support.
- Notes have no author.
- Page setup always carries ExcelJS defaults (`orientation`, `scale`,
  `fitToWidth`, `fitToHeight`) on read and write.
- The legacy `password` attribute of sheet protection is not read.
- Images are written as `oneCellAnchor` elements carrying an `editAs`
  attribute.

## Oracle operations (`exceljs-oracle.cjs`)

- `write_fixture`: writes sheets, cells, formulas, styles, rich text, comments,
  hyperlinks, tables, data validations, merged ranges, freeze panes, images,
  and sheet protection.
- `read_metadata`: inspects package parts for worksheets, tables, drawings,
  media, comments, VML drawings, shared strings, calc-chain metadata, and data
  validations.
