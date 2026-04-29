# Launch

## One-liner

WolfXL is a Rust-backed, openpyxl-compatible Excel engine, and ExcelBench is the benchmark suite that measures whether spreadsheet libraries preserve the workbook features people actually care about.

## What is launching

This launch has three lanes:

1. Python replacement lane

WolfXL reaches `18/18` green features in the current scored Python release snapshot.

2. Cross-language context lane

The checked-in cross-language snapshot shows both `Apache POI` and `Excelize` at `18/18` in the scored write lane.

3. Pivot capability lane

The scored macOS fixture is not valid for pivots, so pivot evidence lives in a separate artifact. That artifact shows `Excelize` can emit pivot-bearing workbooks and `openpyxl` can read back the emitted pivot metadata.

## Core message

Most spreadsheet comparisons focus on speed. The real question is whether a library can handle complex workbooks without quietly dropping the parts that matter. ExcelBench measures that directly, and WolfXL now has the evidence to claim both high fidelity and strong performance in Python.

## Launch headline options

- WolfXL reaches 18/18 in our Excel fidelity benchmark
- We built an Excel benchmark that tests what spreadsheet libraries actually preserve
- WolfXL now matches openpyxl on scored fidelity and adds a patch-based modify path

## Main announcement draft

We spent the last stretch turning WolfXL and ExcelBench from a promising project into something we can defend technically.

WolfXL is a Rust-backed, openpyxl-compatible Excel engine for Python. ExcelBench is the benchmark suite we built to answer a simple question that most spreadsheet comparisons skip: can this library handle a real workbook without breaking the parts you care about?

The current release snapshot is the first one that feels clean enough to publish:

- WolfXL: `18/18` green features in the scored Python release lane
- Apache POI: `18/18` in the cross-language scored write lane
- Excelize: `18/18` in the cross-language scored write lane
- Pivot tables: tracked in a separate capability artifact on macOS because the shipped fixture is not scoreable there, while `Excelize` can still emit pivot-bearing workbooks

The important part is not just the score. The benchmark now has distinct lanes:

- a Python replacement lane for migration decisions
- a cross-language lane for ecosystem context
- a separate pivot capability lane when the scored fixture is not valid on this platform

That separation matters because it keeps the claims honest. We are no longer mixing historical snapshots, capability demos, and scored benchmark results into one muddy story.

If you process spreadsheets in Python, the practical takeaway is straightforward: WolfXL now has a strong case as a serious openpyxl alternative, and ExcelBench now has enough rigor to be useful as a benchmark in its own right.

## HN draft

Title:

WolfXL reached 18/18 in our Excel fidelity benchmark

Body:

I built two related projects:

- WolfXL: a Rust-backed, openpyxl-compatible Excel engine for Python
- ExcelBench: a benchmark suite for spreadsheet fidelity and performance

The benchmark question is simple: not just “how fast is this library?”, but “can it actually preserve the workbook features people care about?”

Current state:

- WolfXL hits `18/18` in the scored Python release lane
- Apache POI and Excelize both hit `18/18` in the scored cross-language write lane
- pivot tables are tracked separately on macOS because the shipped fixture is not scoreable there, but Excelize can still emit pivot-bearing workbooks and openpyxl can read the resulting pivot metadata

I think the most interesting part is the benchmark design, not just the project score:

- Python replacement lane for migration decisions
- cross-language lane for ecosystem context
- separate capability lane for pivots when the platform fixture is not valid

Repo links:

- WolfXL: <add repo URL>
- ExcelBench: <add repo URL>

If you work on spreadsheet tooling, I’d especially like feedback on the benchmark methodology and fixture design.

## Investor / technical summary

WolfXL is a Python spreadsheet engine with a Rust core and an openpyxl-style API. ExcelBench is the benchmark harness that measures spreadsheet fidelity rather than just throughput.

The project is now in a much stronger position because the proof is cleaner:

- WolfXL reaches `18/18` in the scored Python release lane
- cross-language reference points are strong: Apache POI `18/18`, Excelize `18/18`
- pivot capability is broken out into a separate artifact instead of being overstated in the main scorecard

That gives us three things:

1. a credible Python replacement story
2. a credible ecosystem-positioning story
3. a benchmark asset that is useful beyond WolfXL itself

## Assets to link

- Python release snapshot: `results-release-2026-04-28/README.md`
- Python release dashboard: `results-release-2026-04-28/DASHBOARD.md`
- Cross-language snapshot: `results-cross-language/README.md`
- Pivot capability artifact: `results-cross-language-pivots/README.md`
- Reporting policy: `docs/public-reporting.md`
- Cross-language context explainer: `docs/cross-language-context.md`

## Claim guardrails

- Say `scored Python release lane` for WolfXL results.
- Say `cross-language context snapshot` for Apache POI and Excelize.
- Say `pivot capability artifact` for the separate pivot evidence.
- Do not merge historical baseline, perf snapshot, and release snapshot into one undated claim.
- Do not imply that the macOS pivot fixture is currently scoreable.

## Recommended order for public rollout

1. Publish the README/docs update and the checked-in artifacts.
2. Post the main launch note using the announcement draft above.
3. Post the HN version with the benchmark-methodology angle.
4. Use the investor / technical summary in direct outreach.
