# ExcelJS External Oracle

This optional helper generates `.xlsx` workbooks with ExcelJS. It uses the same
JSON stdin/stdout contract as the Excelize, LibreOffice, ClosedXML, and NPOI
helpers.

Run from this directory once dependencies are installed:

```bash
npm install
npm run oracle
```

Supported operations:

- `write_fixture`: writes sheets, cells, formulas, styles, rich text, comments,
  hyperlinks, tables, data validations, merged ranges, freeze panes, images,
  and sheet protection.
- `read_metadata`: inspects package parts for worksheets, tables, drawings,
  media, comments, VML drawings, shared strings, calc-chain metadata, and data
  validations.

ExcelJS is intentionally a local pre-release oracle, not a normal ExcelBench
adapter or public benchmark claim.
