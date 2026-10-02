# ExcelBench Formula Recalculation Results

Oracle: libreoffice/LibreOffice 26.8.0.3 bce0998afefdbc355585ca324285661a2170ba77

Formula values are compared with absolute tolerance 1e-6 or relative tolerance 1e-9 for numbers; strings and booleans compare exactly. Saved to file counts the matching values a save writes into the workbook, for engines that return results through an API.

| Engine | Version | Status | Matched/total | Saved to file | First mismatches |
| --- | --- | --- | --- | --- | --- |
| wolfxl | 2.0.7 | passed | 133/133 | 0/133 | — |
| wolfxl-commercial | 2.3.0 | passed | 133/133 | 133/133 | — |
| libreoffice | 26.8.0.3 (build bce0998afefdbc355585ca324285661a2170ba77) | passed | 133/133 | — | — |
| aspose_cells_foss | 26.7.0 | failed | 25/133 | — | Schedule!B11: expected -308872.0, actual None<br>Schedule!B2: expected 100000.0, actual None<br>Schedule!B3: expected 38000.0, actual None<br>Schedule!B4: expected 62000.0, actual None<br>Schedule!B6: expected 488.0, actual None<br>Schedule!B7: expected 719.0, actual None<br>Schedule!B8: expected 400872.0, actual None<br>Schedule!B9: expected -300872.0, actual None<br>Schedule!C11: expected -306442.0, actual None<br>Schedule!C2: expected 104500.0, actual None |
| zavora | unknown | failed | 0/133 | — | Schedule!B10: expected 0.0, actual None<br>Schedule!B11: expected -308872.0, actual None<br>Schedule!B2: expected 100000.0, actual None<br>Schedule!B3: expected 38000.0, actual None<br>Schedule!B4: expected 62000.0, actual None<br>Schedule!B5: expected 12000.0, actual None<br>Schedule!B6: expected 488.0, actual None<br>Schedule!B7: expected 719.0, actual None<br>Schedule!B8: expected 400872.0, actual None<br>Schedule!B9: expected -300872.0, actual None |

aspose_cells_foss: FOSS FormulaEvaluator returned no value for 107 formulas: Schedule!B2, Schedule!C2, Schedule!D2, Schedule!E2, Schedule!F2

zavora: 133/133 formula values wrong or missing (first mismatch batch: 10 shown, 10 missing, 0 wrong)

## Scope (snapshot runner note)

The calc tier scores engines that expose a recalculation API. Libraries without
one are **not applicable** here, not scored 0:

| Adapter | Calc tier | Reason |
| --- | --- | --- |
| openpyxl, openpyxl-readonly, xlsxwriter, xlsxwriter-constmem, pandas, polars, python-calamine, pylightxl, pyexcel, tablib, xlrd, xlwt | N/A | No formula evaluation engine; they read or write cached values / formula text only |
| sheetjs (CE 0.20.3) | N/A | SheetJS Community Edition does not evaluate formulas |
| exceljs 4.4.0 | N/A | ExcelJS stores formulas and cached results but has no calculation engine |
| apache-poi 5.5.1 | Not run | POI ships `FormulaEvaluator`, but the ExcelBench POI helper has no evaluate operation, so there is no harness path to score it yet |
| excelize 2.10.1 | Not run | Excelize ships `CalcCellValue`, but the ExcelBench Excelize helper has no evaluate operation, so there is no harness path to score it yet |

The `libreoffice` row above is LibreOffice 26.8.0.3 headless recalculation (it is also the oracle that produced `fixtures/calc/expected_values.json`).
`wolfxl` is WolfXL Community 2.0.7 from PyPI. Its score counts the values `Workbook.calculate()` returns; its `save()` writes no calculated values, shown as Saved to file 0/133.
`wolfxl-commercial` is WolfXL Commercial 2.3.0 from SynthGL's authenticated package index. Both editions install as `wolfxl`, so it ran in a separate interpreter named by `EXCELBENCH_WOLFXL_COMMERCIAL_PYTHON`.

Command: `EXCELBENCH_WOLFXL_COMMERCIAL_PYTHON=<commercial venv>/bin/python excelbench calc -o results-2026-10-02/calc` (defaults: `fixtures/calc/financial_model.xlsx`, `fixtures/calc/expected_values.json`).
