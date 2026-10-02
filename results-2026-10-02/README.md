# ExcelBench competitor snapshot 2026-10-02

| Lane | Artifacts | Command |
| --- | --- | --- |
| Fidelity (17 adapters) | [xlsx/](xlsx/README.md) | `excelbench benchmark -t fixtures/excel -o results-2026-10-02/xlsx -a openpyxl -a xlsxwriter -a python-calamine -a aspose-cells-foss -a wolfxl -a pylightxl -a xlrd -a pyexcel -a xlwt -a pandas -a xlsxwriter-constmem -a openpyxl-readonly -a polars -a tablib -a sheetjs -a exceljs -a libreoffice` then `excelbench heatmap -i results-2026-10-02/xlsx/results.json -o results-2026-10-02/xlsx` |
| Template mutation | [mutation/](mutation/README.md) | `excelbench mutation -o results-2026-10-02/mutation --repeats 3` |
| Formula recalculation | [calc/](calc/README.md) | `EXCELBENCH_WOLFXL_COMMERCIAL_PYTHON=<commercial venv>/bin/python excelbench calc -o results-2026-10-02/calc` (rerun later the same day; see Versions) |
| Cross-language context | [cross-language/](cross-language/README.md) | `EXCELBENCH_ORACLE_DOCKER_CONTEXT=pc excelbench cross-language-context -t fixtures/excel -o results-2026-10-02/cross-language` |

## Versions

| Library | Version | Runtime |
| --- | --- | --- |
| wolfxl | 2.0.5 (latest PyPI release when run); the calc rerun used 2.0.7 | Python wheel |
| wolfxl-commercial (calc only) | 2.3.0 from SynthGL's authenticated index | Python wheel, separate interpreter |
| openpyxl / openpyxl-readonly | 3.1.5 | Python |
| aspose-cells-foss | 26.7.0 | Python |
| xlsxwriter / xlsxwriter-constmem | 3.2.9 | Python |
| python-calamine | 0.6.1 | Python |
| pandas | 3.0.0 | Python |
| polars | 1.38.1 | Python |
| pyexcel | 0.7.4 | Python |
| pylightxl | 1.61 | Python |
| tablib | 3.9.0 | Python |
| xlrd | 2.0.2 | Python |
| xlwt | 1.3.0 | Python |
| sheetjs (npm `xlsx`, Community Edition) | 0.20.3 | Node v22.18.0 via `tools/external-oracles/sheetjs` |
| exceljs | 4.4.0 | Node v22.18.0 via `tools/external-oracles/exceljs` |
| libreoffice | 26.8.0.3 (build bce0998afefdbc355585ca324285661a2170ba77) | headless `soffice` + UNO macro via `tools/external-oracles/libreoffice` |
| zavora-xlsx | 0.1.2 | Rust, local `cargo run` (cargo 1.92.0) via `tools/external-oracles/zavora` |
| apache-poi | 5.5.1 | Docker image `excelbench-poi-oracle:5.5.1` (`ef46106ed14d`) on context `pc` (linux/amd64) |
| excelize | 2.10.1 | Docker image `excelbench-excelize-oracle:2.10.1` (`425caf1a873f`) on context `pc` (linux/amd64) |

## Host

macOS 26.5.1 (25F80), Apple M5 Pro, 64 GB, arm64; Python 3.13.9; ExcelBench at `b65866c` plus the uncommitted `officelibs-m2` adapter work.
The fidelity and mutation lanes ran while the host carried heavy unrelated load; fidelity scores are unaffected, mutation timings are not comparable (see the mutation README).

## Lane membership

- Fidelity: every 2026-09-08 adapter plus `sheetjs`, `exceljs`, `libreoffice`. `pivot_tables` is not scored by any adapter, so green-feature denominators are /21.
- Mutation: every engine that can open and edit an existing workbook (`openpyxl`, `wolfxl`, `aspose-cells-foss`, `zavora-xlsx`, `sheetjs`, `exceljs`, `libreoffice`).
- Calc: engines with a recalculation API; non-calculating adapters are listed as N/A in the calc README.
- Cross-language: `apache-poi`, `excelize`, `zavora-xlsx`.
