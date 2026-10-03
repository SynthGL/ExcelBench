# ExcelBench Mutation Results

*Generated: 2026-10-03T00:25:19Z*
*Template SHA-256: da08cc9e607f267b438ec7ce2480d6eb3650d1c19928f283a18296a2b79d808e*
*Repeats: 3*

## Comparison

| Engine | Wall ms | Peak RSS KB | Features preserved | Unexpected changes | Edits applied | Integrity | Verdict |
|--------|---------|-------------|--------------------|--------------------|---------------|-----------|---------|
| aspose-cells-foss | 750.000 | 168128.000 | 15/17 | cells (87), custom_xml (1) | yes | ok | Changed cells, custom_xml |
| exceljs | — | — | — | — | — | — | Failed |
| libreoffice | 1660.000 | 264352.000 | 15/17 | cell_styles (16), page_setup (5) | yes | ok | Changed cell_styles, page_setup |
| openpyxl | 860.000 | 167968.000 | 16/17 | custom_xml (1) | yes | ok | Changed custom_xml |
| sheetjs | 800.000 | 167136.000 | 8/17 | cell_styles (9), views (3), data_validations (2), conditional_formats (2), protection (1), page_setup (2), charts (2), tables (1), custom_xml (1) | yes | ok | Changed cell_styles, views, data_validations, conditional_formats, protection, page_setup, charts, tables, custom_xml |
| wolfxl | 780.000 | 170032.000 | 17/17 | none | yes | ok | Preserved |
| zavora-xlsx | 1240.000 | 167184.000 | 8/17 | cells (48), cell_styles (17), data_validations (2), conditional_formats (2), comments (1), charts (2), tables (1), theme (1), custom_xml (1) | yes | ok | Changed cells, cell_styles, data_validations, conditional_formats, comments, charts, tables, theme, custom_xml |

## Notes

Wall time and peak RSS are measured in isolated subprocesses with macOS `/usr/bin/time -l`.
Preservation compares the output's workbook content model against the template's model with the two declared cell edits applied. Part names, relationship ids, XML serialization and document metadata are not content, so an equivalent rewrite scores the same as a byte copy. Cached results of formula cells are recalculation state and are not compared; formula text is. Features preserved counts template features with no unexpected difference. Integrity (dangling relationships, parts without a content type) is reported separately and is not scored. This does not measure formula recalculation or rendering.

## Scope (snapshot runner note)

- **Scorer.** The same content-model scorer as the 2026-10-02 rerun (template content model with the two declared edits applied, compared with the output's model, plus the `custom_xml` feature). Fixture, scorer and engines are unchanged; only the WolfXL version moved (2.0.5 to 2.0.8).
- **Verdicts.** Identical to 2026-10-02 for every engine: wolfxl Preserved (17/17); openpyxl Changed custom_xml (16/17); LibreOffice Changed cell_styles, page_setup (15/17); aspose-cells-foss Changed cells, custom_xml (15/17); SheetJS 8/17 and zavora-xlsx 8/17 with the same changed features; ExcelJS Failed. Every engine that ran applied both declared edits. The per-feature explanations in the 2026-10-02 mutation README still apply.
- `exceljs` 4.4.0 failed the same way as 2026-10-02: ExcelJS throws `Cannot read properties of undefined (reading 'anchors')` while loading the template (full traceback in `results.json`).
- **Timing.** This table is a rerun. The first attempt (started 00:02:15Z from `run.sh`, which waits up to 10 minutes for the 1-minute load to reach 8 or less) began at load `{ 7.18 6.93 5.56 }`, but unrelated jobs pushed it to `{ 21.30 11.57 7.66 }` by the end, and its timings were skewed (openpyxl 60770 ms against 1300 to 3000 ms for the rest). Its verdicts were identical to the table above; it was discarded only for timing. The rerun started at 00:24:57Z once the 1-minute load was back at or below 8 on this 18-core host: `{ 7.61 14.14 13.53 }` before and `{ 6.81 13.53 13.32 }` after. The host was not idle (5-minute average about 14), so read the wall times as a rough ordering: aspose-cells-foss 750 ms, wolfxl 780 ms, sheetjs 800 ms, openpyxl 860 ms, zavora-xlsx 1240 ms, LibreOffice 1660 ms. Differences of a few tens of milliseconds are within noise. Do not compare these timings with 2026-10-02, which ran under much heavier load. Peak RSS is 167 to 170 MB for every engine except LibreOffice (264 MB).
- `apache-poi` and `excelize` are not in this lane: their ExcelBench helpers are write-only (no open-existing-workbook operation), and the suite times each engine in a local `/usr/bin/time -l` subprocess, which a remote container cannot provide.

Command: `excelbench mutation -o results-2026-10-03/mutation --repeats 3` (rerun with the same command after a load gate of 8.0).
