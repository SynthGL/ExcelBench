# ExcelBench Formula Recalculation Results

Oracle: libreoffice/LibreOffice 26.8.0.3 bce0998afefdbc355585ca324285661a2170ba77

Formula values are compared with absolute tolerance 1e-6 or relative tolerance 1e-9 for numbers; strings and booleans compare exactly.

| Engine | Status | Matched/total | First mismatches |
| --- | --- | --- | --- |
| wolfxl | failed | 55/133 | Schedule!B11: expected -308872.0, actual None<br>Schedule!B2: expected 100000.0, actual None<br>Schedule!B3: expected 38000.0, actual None<br>Schedule!B4: expected 62000.0, actual None<br>Schedule!B8: expected 400872.0, actual None<br>Schedule!B9: expected -300872.0, actual None<br>Schedule!C11: expected -306442.0, actual None<br>Schedule!C2: expected 104500.0, actual None<br>Schedule!C3: expected 39710.0, actual None<br>Schedule!C4: expected 64790.0, actual None |

wolfxl: 78/133 formula values wrong or missing (first mismatch batch: 10 shown, 10 missing, 0 wrong)
| libreoffice | passed | 133/133 | — |
| aspose_cells_foss | failed | 25/133 | Schedule!B11: expected -308872.0, actual None<br>Schedule!B2: expected 100000.0, actual None<br>Schedule!B3: expected 38000.0, actual None<br>Schedule!B4: expected 62000.0, actual None<br>Schedule!B6: expected 488.0, actual None<br>Schedule!B7: expected 719.0, actual None<br>Schedule!B8: expected 400872.0, actual None<br>Schedule!B9: expected -300872.0, actual None<br>Schedule!C11: expected -306442.0, actual None<br>Schedule!C2: expected 104500.0, actual None |

aspose_cells_foss: FOSS FormulaEvaluator returned no value for 107 formulas: Schedule!B2, Schedule!C2, Schedule!D2, Schedule!E2, Schedule!F2
| zavora | failed | 0/133 | Schedule!B10: expected 0.0, actual None<br>Schedule!B11: expected -308872.0, actual None<br>Schedule!B2: expected 100000.0, actual None<br>Schedule!B3: expected 38000.0, actual None<br>Schedule!B4: expected 62000.0, actual None<br>Schedule!B5: expected 12000.0, actual None<br>Schedule!B6: expected 488.0, actual None<br>Schedule!B7: expected 719.0, actual None<br>Schedule!B8: expected 400872.0, actual None<br>Schedule!B9: expected -300872.0, actual None |

zavora: 133/133 formula values wrong or missing (first mismatch batch: 10 shown, 10 missing, 0 wrong)
