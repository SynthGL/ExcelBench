# ExcelBench competitor snapshot 2026-10-03

| Lane | Artifacts | Command |
| --- | --- | --- |
| Fidelity (17 adapters) | [xlsx/](xlsx/README.md) | `excelbench benchmark -t fixtures/excel -o results-2026-10-03/xlsx -a openpyxl -a xlsxwriter -a python-calamine -a aspose-cells-foss -a wolfxl -a pylightxl -a xlrd -a pyexcel -a xlwt -a pandas -a xlsxwriter-constmem -a openpyxl-readonly -a polars -a tablib -a sheetjs -a exceljs -a libreoffice` then `excelbench heatmap -i results-2026-10-03/xlsx/results.json -o results-2026-10-03/xlsx` |
| Template mutation | [mutation/](mutation/README.md) | `excelbench mutation -o results-2026-10-03/mutation --repeats 3` |
| Formula recalculation | [calc/](calc/README.md) | `EXCELBENCH_WOLFXL_COMMERCIAL_PYTHON=<commercial venv>/bin/python excelbench calc -o results-2026-10-03/calc` |
| Cross-language context | [cross-language/](cross-language/README.md) | `EXCELBENCH_ORACLE_DOCKER_CONTEXT=pc excelbench cross-language-context -t fixtures/excel -o results-2026-10-03/cross-language` |
| Performance (13 Python adapters, separate host) | [perf/](perf/README.md) | `excelbench perf --tests fixtures/excel --profile xlsx --warmup 3 --iters 25 --iteration-policy fixed --memory-mode getrusage` with 19 `--feature` and 13 `--adapter` flags (full command in [perf/README.md](perf/README.md)) |

The fidelity, mutation, calc and cross-language lanes ran from [run.sh](run.sh) (`EB_VENV` and `EB_COMMERCIAL` name the two interpreters); its console output is [run-log.txt](run-log.txt). The per-lane console logs it writes to `logs/` are gitignored and not committed; every score is in each lane's `results.json`. Perf ran separately on a different host (Apple M4 Pro, Python 3.12.3, the WolfXL 2.0.8 cp312 wheel) with the 2026-10-02 perf command; its receipt, pins and log are in [perf/](perf/README.md). Fixtures, scoring and adapters are the ones used for 2026-10-02; the only code change is the Excelize helper port to 2.11.0 (see the cross-language README).

## Versions

Every library was checked against its registry on 2026-10-02 UTC, shortly before the run. "Bumped" marks a version newer than 2026-10-02.

| Library | Version | Runtime |
| --- | --- | --- |
| wolfxl | 2.0.8 from PyPI (bumped from 2.0.5 / calc 2.0.7), `wolfxl-2.0.8-cp313-cp313-macosx_11_0_arm64.whl`, sha256 `c410d299ebbc4bdf7290d8dd872d0cc251f69316bbbe3939eb7561ad8c0cbfb7` | Python wheel |
| wolfxl-commercial (calc only) | 2.3.0 (newest Commercial release) from SynthGL's authenticated index, `wolfxl-2.3.0-cp313-cp313-macosx_11_0_arm64.whl`, sha256 `b49ce6ff3bfea1f829ea4f922b78485e0badb71633406cd492d127a95b9ee1ec` | Python wheel, separate interpreter |
| openpyxl / openpyxl-readonly | 3.1.5 | Python |
| aspose-cells-foss | 26.7.0 | Python |
| xlsxwriter / xlsxwriter-constmem | 3.2.9 | Python |
| python-calamine | 0.8.2 (bumped from 0.6.1) | Python |
| pandas | 3.0.6 (bumped from 3.0.0) | Python |
| polars | 1.44.2 (bumped from 1.38.1) | Python |
| pyexcel | 0.7.6 (bumped from 0.7.4) | Python |
| pylightxl | 1.61 | Python |
| tablib | 3.10.0 (bumped from 3.9.0) | Python |
| xlrd | 2.0.2 | Python |
| xlwt | 1.3.0 | Python |
| sheetjs (npm `xlsx`, Community Edition) | 0.20.3 (latest on the SheetJS CDN) | Node v22.18.0 via `tools/external-oracles/sheetjs` |
| exceljs | 4.4.0 | Node v22.18.0 via `tools/external-oracles/exceljs` |
| libreoffice | 26.8.0.3 (build bce0998afefdbc355585ca324285661a2170ba77) | headless `soffice` + UNO macro via `tools/external-oracles/libreoffice` |
| zavora-xlsx | 0.1.2 | Rust, local `cargo run` (cargo 1.92.0) via `tools/external-oracles/zavora` |
| apache-poi | 5.5.1 | Docker image `excelbench-poi-oracle:5.5.1` (`ef46106ed14d`) on context `pc` (linux/amd64) |
| excelize | 2.11.0 (bumped from 2.10.1) | Docker image `excelbench-excelize-oracle:2.11.0` (`79f691f8e6d0`) on context `pc` (linux/amd64) |

The Python pins are in [pins.txt](pins.txt) and the full environment in [pip-freeze.txt](pip-freeze.txt). `uv.lock` was not changed; the run used its own venv built from `pins.txt`.

## Host

macOS 26.5.1 (25F80), Apple M5 Pro, 64 GB, arm64; Python 3.13.9; ExcelBench `master` at `b453355` plus the Excelize 2.11.0 helper changes (`tools/external-oracles/excelize/{go.mod,go.sum,main.go,Dockerfile}`, the image tag in `src/excelbench/harness/external_oracles.py`, `tools/external-oracles/remote/README.md`).

## Lane membership

- Fidelity: the same 17 adapters as 2026-10-02. `pivot_tables` is not scored by any adapter, so green-feature denominators are /21.
- Mutation: every engine that can open and edit an existing workbook (`openpyxl`, `wolfxl`, `aspose-cells-foss`, `zavora-xlsx`, `sheetjs`, `exceljs`, `libreoffice`).
- Calc: engines with a recalculation API; non-calculating adapters are listed as N/A in the calc README.
- Cross-language: `apache-poi`, `excelize`, `zavora-xlsx`.

## Changes from 2026-10-02

- Fidelity: openpyxl 3.1.5 scores 21/21 read and 21/21 write (133/133 tests in each mode), as on 2026-10-02. WolfXL 2.0.8 also scores 21/21 read and 21/21 write (133/133 in each mode), up from 19/21 read and 18/21 write with 2.0.5; the four features that changed are `named_ranges` (read 2 to 3, write 2 to 3), `tables` (read 2 to 3), `page_setup` (write 2 to 3) and `chart_anchor` (write 0 to 3). Every other adapter has the same per-feature scores as 2026-10-02, including the five bumped Python libraries.
- Cross-language: identical scores (Apache POI 18/21, Excelize 18/21 after the 2.11.0 bump, zavora-xlsx 8/21).
- Calc: identical except the WolfXL Community version (2.0.8, 133/133 calculated, 0/133 saved).
- Mutation: identical verdicts (wolfxl Preserved 17/17, openpyxl 16/17, LibreOffice 15/17, aspose-cells-foss 15/17, SheetJS 8/17, zavora-xlsx 8/17, ExcelJS fails to load the template). The committed table is a rerun at 00:24:57Z with the same command, because unrelated host load spiked during the `run.sh` attempt and skewed its timings; `run-log.txt` covers only the `run.sh` attempt. Details and load readings are in the mutation README.
- Perf ([perf/](perf/README.md)): the only package change is WolfXL 2.0.5 to 2.0.8. WolfXL write fell from 22.28 to 11.26 ms (sum of 19 per-feature p50s; 1.20x to 2.45x openpyxl) and read is flat (13.46 to 13.90 ms; 1.69x to 1.73x). python-calamine (13.92x openpyxl) and polars (3.87x) still read faster than WolfXL. The other adapters moved by -3.8% to +13.3% with no version change, which is run-to-run variance.
