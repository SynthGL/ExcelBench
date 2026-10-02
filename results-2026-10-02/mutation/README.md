# ExcelBench Mutation Results

*Generated: 2026-10-02T08:10:41Z*
*Template SHA-256: da08cc9e607f267b438ec7ce2480d6eb3650d1c19928f283a18296a2b79d808e*
*Repeats: 3*

## Comparison

| Engine | Wall ms | Peak RSS KB | Features preserved | Unexpected changes | Edits applied | Integrity | Verdict |
|--------|---------|-------------|--------------------|--------------------|---------------|-----------|---------|
| aspose-cells-foss | 1140.000 | 160512.000 | 15/17 | cells (87), custom_xml (1) | yes | ok | Changed cells, custom_xml |
| exceljs | — | — | — | — | — | — | Failed |
| libreoffice | 2550.000 | 264432.000 | 15/17 | cell_styles (16), page_setup (5) | yes | ok | Changed cell_styles, page_setup |
| openpyxl | 1560.000 | 159904.000 | 16/17 | custom_xml (1) | yes | ok | Changed custom_xml |
| sheetjs | 1160.000 | 159248.000 | 8/17 | cell_styles (9), views (3), data_validations (2), conditional_formats (2), protection (1), page_setup (2), charts (2), tables (1), custom_xml (1) | yes | ok | Changed cell_styles, views, data_validations, conditional_formats, protection, page_setup, charts, tables, custom_xml |
| wolfxl | 1470.000 | 162112.000 | 17/17 | none | yes | ok | Preserved |
| zavora-xlsx | 1910.000 | 159184.000 | 8/17 | cells (48), cell_styles (17), data_validations (2), conditional_formats (2), comments (1), charts (2), tables (1), theme (1), custom_xml (1) | yes | ok | Changed cells, cell_styles, data_validations, conditional_formats, comments, charts, tables, theme, custom_xml |

## Notes

Wall time and peak RSS are measured in isolated subprocesses with macOS `/usr/bin/time -l`.
Preservation compares the output's workbook content model against the template's model with the two declared cell edits applied. Part names, relationship ids, XML serialization and document metadata are not content, so an equivalent rewrite scores the same as a byte copy. Cached results of formula cells are recalculation state and are not compared; formula text is. Features preserved counts template features with no unexpected difference. Integrity (dangling relationships, parts without a content type) is reported separately and is not scored. This does not measure formula recalculation or rendering.

## Scope (snapshot runner note)

- **Scorer.** Rerun on 2026-10-02 with the content-model scorer: the template's xlsx content model (vendored verbatim from `officelibs-corpus/src/officecorpus/model/xlsx.py`, sha256 `17fb55bc132caf91a23f07e1df35423dda7f5e14051837f88dd708ccf7b19e61`) with the two declared edits applied is compared with the output's model, plus a `custom_xml` feature (canonical XML and datastore item id). Part names, relationship ids, XML encoding and metadata are not scored. This replaces the earlier part-name/byte score (40 parts + 30 elements + 20 integrity + 10 byte-identical customXml), which scored a byte copy 100 by construction.
- **Old score -> new verdict.** wolfxl 100 -> Preserved (17/17); openpyxl 60 -> Changed custom_xml (16/17); LibreOffice 70 -> Changed cell_styles, page_setup (15/17); aspose-cells-foss 60 -> Changed cells, custom_xml (15/17); SheetJS 20 -> Changed cell_styles, views, data_validations, conditional_formats, protection, page_setup, charts, tables, custom_xml (8/17); zavora-xlsx 0 (integrity-failed) -> Changed cells, cell_styles, data_validations, conditional_formats, comments, charts, tables, theme, custom_xml (8/17); ExcelJS Failed -> Failed.
- **What each engine changes.** openpyxl drops the custom XML data item. LibreOffice gives the header row (`Inputs!A1:H1`), whose font the template leaves without a name or size, Cambria 11 (16 cell_styles diffs) and changes header/footer margins from 0.5 in to 0.512 in on Inputs and Schedule plus the Summary footer (5 page_setup diffs); its comment written as `xl/comments1.xml` and its re-serialized customXml, which cost it 30 points before, are no longer counted. aspose-cells-foss drops 87 string cells across all three sheets and the custom XML item. SheetJS drops the header-row styles, all three freeze panes, both data validations, both conditional formats, the Schedule sheet protection, the Summary landscape orientation and page header, both charts, the table and custom XML, and changes the default font size from 11 to 12. zavora-xlsx blanks 48 string cells on Inputs, gives the header-row font Calibri 11 and the default font color white (`FFFFFF` instead of theme 1) (17 cell_styles diffs), changes the decimal validation's allow-blank flag and drops the cellIs rule's fill, and drops the comment, both charts, the table, the referenced theme color and custom XML. Per-feature counts and the first 10 diffs per engine are in `results.json`.
- **zavora-xlsx edits.** Both declared edits land (`edits_applied: true`). The previous `integrity-failed` came from the openpyxl verifier, which cannot open this output (`IndexError` on an out-of-range style index); the content model reads it.
- **Timing is not comparable.** `vm.loadavg` was `{ 36.85 75.08 147.92 }` immediately before the run (08:10:09Z) and `{ 26.92 68.35 142.52 }` right after it (08:10:45Z), far above 16, because of unrelated host load. Wall-ms and RSS are recorded as measured but must not be compared with each other or with 2026-09-08. Preservation verdicts do not depend on timing.
- `exceljs` 4.4.0 failed: ExcelJS throws `Cannot read properties of undefined (reading 'anchors')` while loading the template (full traceback in `results.json`).
- `apache-poi` and `excelize` are not in this lane: their ExcelBench helpers are write-only (no open-existing-workbook operation), and the suite times each engine in a local `/usr/bin/time -l` subprocess, which a remote container cannot provide.

Command: `excelbench mutation -o results-2026-10-02/mutation --repeats 3`.
