# ExcelBench calc and small-file perf refresh: 2026-10-27

This snapshot measures WolfXL Community 2.0.21 from public PyPI. ExcelBench is maintained by the authors of WolfXL; this is not independent testing. The harness, fixtures, commands, saved workbooks, and raw results are public.

The directory uses the local run date, 2026-10-27 (America/Los_Angeles, UTC-07:00). The run began at **2026-10-28T01:53:53Z**, as preserved in the raw metadata and [driver log](run-log.txt). No timestamps were backdated. Fidelity and mutation were not rerun: continue to cite their 2026-10-03 artifacts and WolfXL 2.0.8 version.

## Results

- [Calc report](calc/README.md) and [raw JSON](calc/results.json): Community 2.0.21 computed **133/133** formula values and saved **133/133** correct cached values in `fixtures/calc/financial_model.xlsx`. This is a single financial fixture, not coverage of every Excel function.
- Commercial was not rerun because `EXCELBENCH_WOLFXL_COMMERCIAL_PYTHON` was not configured. Its **2.3.0, 133/133 computed, 133/133 saved** row is carried forward from [2026-10-03](../results-2026-10-03/calc/results.json), separately preserved in [commercial-carry-forward.json](calc/commercial-carry-forward.json). The new raw suite output honestly records its unavailable status; that is not a zero correctness score.
- [Perf report](perf/README.md) and [raw JSON](perf/results.json): all 13 adapters from the previous perf snapshot ran on the same 19 features, with 3 warmups and 25 recorded iterations per operation. The perf command completed in **14.63 seconds**; the combined calc and perf driver completed in about 38 seconds. Unsupported, failed, and not-applicable operations remain in the raw output and report.
- Speed ratios are within this run only. Host, Python, and WolfXL version changed since 2026-10-03; this is not an isolated before/after version comparison. python-calamine remains the fastest small-file reader in this adapter set.

## Environment and identity

- ExcelBench source commit: `5f71713413c43f3733e303ba8a4eed6a756760cd`, tree `c3033f5741e2b701a5e7f833d65d96a1ac2ebdb5`. The harness, adapters, and fixtures were unchanged when measured. Only the new snapshot driver was added before the run.
- Apple M5 Pro, 18 cores, 64 GiB RAM; macOS 26.5.1 (25F80), native arm64 CPython 3.13.9 (Clang 20.1.4); uv 0.12.3.
- Fresh `.venv-benchmark-native` environment with an editable ExcelBench install and public PyPI dependencies. [requirements.txt](requirements.txt) records every resolved dependency. Non-WolfXL adapter versions match the 2026-10-03 perf pins; transitive dependencies were resolved anew. No local WolfXL source build or private package index was used.
- Wheel: `wolfxl-2.0.21-cp313-cp313-macosx_11_0_arm64.whl`, sha256 `d16ed9ce31fa89338efbe694a96ca504a778ded9e8837b9498ad0cbdfcc0d320`, uploaded 2026-10-27T07:31:16.830604Z. Downloaded from the URL in [receipt.json](receipt.json), hashed, and matched to the public PyPI digest.
- Installed extension: `wolfxl/_rust.cpython-313-darwin.so`, sha256 `8f504835f5a991c5e13405685e4c1f97edc76839b1f60345afebf4bfbe7dbc63`. Its bytes were compared with the downloaded wheel member and were identical. This is the published prebuilt wheel; compiler flags were not independently inspected.
- An initial environment selected an x86_64 interpreter and was rejected before any benchmark ran. All scored output comes from the explicitly selected native arm64 environment.
- Fixture hashes are recorded in [receipt.json](receipt.json). The perf manifest is `984073b67669fb0566b55e35a05447fbaaeb9f0da2d6de1a335baf61de9b382f`, unchanged from the preceding snapshot.

## Measurement boundaries

Perf uses warm in-process timings: three warmups are discarded before 25 measured iterations per library, feature, and operation. Cold start and cold filesystem caches are not measured. Write timing excludes oracle verification. The public WolfXL APIs are `wolfxl.load_workbook` and `wolfxl.Workbook`; other adapters use their existing harness paths without special cases.

The harness has no per-operation timeout. The complete driver was run by a Python `subprocess.run(..., timeout=600)` wrapper, giving calc plus perf a 600-second overall limit; the checked-in shell driver itself does not enforce that wrapper limit. There was one scored attempt and no discarded measurement. The driver records zero exit codes for both commands.

The machine was not exclusively reserved. The 1-minute load was 9.05 before perf and 7.94 after, on an 18-core machine. No other timed workload in this refresh ran locally during perf. This is one local trial, not a controlled lab or a guarantee for other workbooks. [perf-run.txt](perf-run.txt) retains all warnings with checkout-specific prefixes removed. Generated reports use `--` instead of typographic dash placeholders; raw JSON metrics are unmodified.

The README speed column sums per-feature p50 wall times over **19 small feature workbooks**. It mostly measures per-call overhead and does not represent throughput on a large workbook. WolfXL's separate [ecosystem results](https://github.com/SynthGL/wolfxl-oss/tree/main/benchmarks/results), including the [1.6M-cell comparison](https://github.com/SynthGL/wolfxl-oss/blob/main/benchmarks/results/2026-08-18-ecosystem-full-linux-epyc.md), use different workloads, APIs, hosts, and versions. Do not compare those ratios as if the test were the same.

## Reproduce

From the repository root, with uv installed (reference platform: native arm64 macOS, Python 3.13.9):

```bash
uv venv .venv-benchmark-native --python 3.13.9
uv pip install --python .venv-benchmark-native/bin/python --index-url https://pypi.org/simple \
  -e . -r results-2026-10-27/requirements.txt
.venv-benchmark-native/bin/python -c 'import platform, wolfxl; print(platform.machine(), wolfxl.__version__)'
EB_VENV=.venv-benchmark-native bash results-2026-10-27/run.sh
```

On the reference platform the version check must print `arm64 2.0.21`. On other platforms use a supported native Python and record the platform change. The driver writes into this snapshot directory, so run it in a disposable checkout if you want to preserve the committed outputs unchanged. LibreOffice must be installed for its recalculation row; absent optional engines are reported as unavailable. An already configured Commercial interpreter can be supplied through `EXCELBENCH_WOLFXL_COMMERCIAL_PYTHON`; no private index is needed to reproduce the Community row.

For just the calculation lane:

```bash
.venv-benchmark-native/bin/excelbench calc -o results-2026-10-27/calc
```

The exact perf flags are in [run.sh](run.sh). [run-log.txt](run-log.txt), [calc-run.txt](calc-run.txt), [perf-run.txt](perf-run.txt), and [receipt.json](receipt.json) preserve command outcomes, times, load readings, package and source identity.
