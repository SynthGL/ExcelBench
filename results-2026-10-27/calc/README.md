# ExcelBench Formula Recalculation Results

Oracle: libreoffice/LibreOffice 26.8.0.3 bce0998afefdbc355585ca324285661a2170ba77

Formula values are compared with absolute tolerance 1e-6 or relative tolerance 1e-9 for numbers; strings and booleans compare exactly. Saved to file counts the matching values a save writes into the workbook, for engines that return results through an API.

| Engine | Version | Status | Matched/total | Saved to file | First mismatches |
| --- | --- | --- | --- | --- | --- |
| wolfxl | 2.0.21 | passed | 133/133 | 133/133 | -- |
| wolfxl-commercial | 2.3.0 | carried from 2026-10-03, not rerun | 133/133 | 133/133 | -- |
| libreoffice | 26.8.0.3 (build bce0998afefdbc355585ca324285661a2170ba77) | passed | 133/133 | -- | -- |
| aspose_cells_foss | 26.7.0 | failed | 25/133 | -- | Schedule!B11: expected -308872.0, actual None<br>Schedule!B2: expected 100000.0, actual None<br>Schedule!B3: expected 38000.0, actual None<br>Schedule!B4: expected 62000.0, actual None<br>Schedule!B6: expected 488.0, actual None<br>Schedule!B7: expected 719.0, actual None<br>Schedule!B8: expected 400872.0, actual None<br>Schedule!B9: expected -300872.0, actual None<br>Schedule!C11: expected -306442.0, actual None<br>Schedule!C2: expected 104500.0, actual None |
| zavora | unknown | failed | 0/133 | -- | Schedule!B10: expected 0.0, actual None<br>Schedule!B11: expected -308872.0, actual None<br>Schedule!B2: expected 100000.0, actual None<br>Schedule!B3: expected 38000.0, actual None<br>Schedule!B4: expected 62000.0, actual None<br>Schedule!B5: expected 12000.0, actual None<br>Schedule!B6: expected 488.0, actual None<br>Schedule!B7: expected 719.0, actual None<br>Schedule!B8: expected 400872.0, actual None<br>Schedule!B9: expected -300872.0, actual None |

wolfxl-commercial: no Commercial interpreter was configured. Its row above is carried forward from [2026-10-03](../../results-2026-10-03/calc/results.json), not a new measurement. [commercial-carry-forward.json](commercial-carry-forward.json) preserves that row and its source date. The unmodified [new raw results](results.json) record `unavailable`; that is not a zero correctness score.

aspose_cells_foss: FOSS FormulaEvaluator returned no value for 107 formulas: Schedule!B2, Schedule!C2, Schedule!D2, Schedule!E2, Schedule!F2

zavora: 133/133 formula values wrong or missing (first mismatch batch: 10 shown, 10 missing, 0 wrong)

## Run receipt

- Snapshot label: 2026-10-27 local date (America/Los_Angeles); actual run: 2026-10-28T01:53:53Z to 01:54:17Z. See [driver log](../run-log.txt).
- Command: `.venv-benchmark-native/bin/excelbench calc -o results-2026-10-27/calc`, with the default `fixtures/calc/financial_model.xlsx` and `fixtures/calc/expected_values.json`.
- Environment: Apple M5 Pro, 18 cores, 64 GiB RAM, macOS 26.5.1 (25F80), native arm64 CPython 3.13.9. ExcelBench commit `5f71713413c43f3733e303ba8a4eed6a756760cd`; unchanged harness and fixtures.
- WolfXL Community 2.0.21 from public PyPI: `wolfxl-2.0.21-cp313-cp313-macosx_11_0_arm64.whl`, sha256 `d16ed9ce31fa89338efbe694a96ca504a778ded9e8837b9498ad0cbdfcc0d320`. The installed native extension matches the wheel bytes; see [receipt.json](../receipt.json).
- Community `calculate()` returns 133/133 matching formula values, and its following `save()` writes 133/133 matching cached values. The harness counts saved values with WolfXL's reader. The [saved workbook](wolfxl/financial_model.xlsx) is included so others can check it with a reader of their choice.
- LibreOffice, aspose-cells-foss, and zavora were rerun. No other rows were carried forward. The oracle is the checked-in expected-values file, not newly generated for this refresh.
- Libraries without a formula evaluation engine are not applicable to this tier, not scored zero. This fixture tests one 133-formula financial DAG, not all Excel functions.

See the [snapshot receipt and reproduce commands](../README.md) for package pins, fixture hashes, the timeout boundary, and publication caveats. Historical snapshots remain unchanged.

