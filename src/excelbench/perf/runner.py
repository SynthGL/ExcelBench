"""Performance benchmark runner.

Design principles:
- Mirror the fidelity harness' feature surface.
- Measure the library under test only.
- Write timings MUST NOT include oracle verification.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass
from datetime import UTC, datetime
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class PerfConfig:
    warmup: int
    iters: int
    iteration_policy: str
    breakdown: bool


@dataclass(frozen=True)
class PerfStats:
    min: float
    p50: float
    p95: float


@dataclass(frozen=True)
class PerfOpResult:
    wall_ms: PerfStats
    cpu_ms: PerfStats
    rss_peak_mb: float | None = None
    breakdown_ms: dict[str, float] | None = None
    phase_attribution_ms: dict[str, float] | None = None
    op_count: int | None = None
    op_unit: str | None = None


@dataclass(frozen=True)
class PerfFeatureResult:
    feature: str
    library: str
    workload_size: str
    perf: dict[str, PerfOpResult | None]
    notes: str | None = None


@dataclass(frozen=True)
class PerfMetadata:
    benchmark_version: str
    run_date: datetime
    excel_version: str
    platform: str
    profile: str
    python: str
    commit: str | None
    config: PerfConfig


@dataclass(frozen=True)
class PerfResults:
    metadata: PerfMetadata
    libraries: dict[str, dict[str, Any]]
    results: list[PerfFeatureResult]


BENCHMARK_VERSION = "0.1.0"


def run_perf(
    test_dir: Path,
    *,
    adapters: list[Any] | None = None,
    features: list[str] | None = None,
    profile: str = "xlsx",
    warmup: int = 3,
    iters: int = 25,
    iteration_policy: str = "fixed",
    breakdown: bool = False,
) -> PerfResults:
    import platform as _platform

    from excelbench.generator.generate import load_manifest
    from excelbench.harness.adapters import get_all_adapters

    test_dir = Path(test_dir)
    manifest_path = test_dir / "manifest.json"
    if not manifest_path.exists():
        raise FileNotFoundError(f"Manifest not found: {manifest_path}")

    if warmup < 0 or iters <= 0:
        raise ValueError("warmup must be >= 0 and iters must be > 0")
    iteration_policy_normalized = iteration_policy.strip().lower()
    if iteration_policy_normalized != "fixed":
        raise ValueError("iteration_policy must be 'fixed'")

    manifest = load_manifest(manifest_path)

    if adapters is None:
        adapters = get_all_adapters()

    if features:
        normalized = {f.strip().lower() for f in features if f.strip()}
        manifest.files = [f for f in manifest.files if f.feature in normalized]
        if not manifest.files:
            missing_list = ", ".join(sorted(normalized))
            raise ValueError(f"No matching features in manifest: {missing_list}")

    metadata = PerfMetadata(
        benchmark_version=BENCHMARK_VERSION,
        run_date=datetime.now(UTC),
        excel_version=manifest.excel_version,
        platform=f"{_platform.system()}-{_platform.machine()}",
        profile=profile,
        python=_platform.python_version(),
        commit=_get_git_commit(),
        config=PerfConfig(
            warmup=warmup,
            iters=iters,
            iteration_policy=iteration_policy_normalized,
            breakdown=breakdown,
        ),
    )

    libraries = {a.name: _library_info_dict(a.info) for a in adapters}

    results: list[PerfFeatureResult] = []
    for test_file in manifest.files:
        workload = _extract_single_workload(test_file)
        workload_ops = _workload_operations(workload)
        file_path = test_dir / test_file.path
        file_exists = file_path.exists()

        for adapter in adapters:
            notes_parts: list[str] = []
            read_res: PerfOpResult | None = None
            write_res: PerfOpResult | None = None

            if "read" in workload_ops and not adapter.can_read():
                notes_parts.append("Read unsupported")
            if "write" in workload_ops and not adapter.can_write():
                notes_parts.append("Write unsupported")

            if adapter.can_read():
                if "read" not in workload_ops:
                    # Workload explicitly excludes read.
                    pass
                elif not file_exists:
                    notes_parts.append(f"Read skipped: missing input file {test_file.path}")
                elif not adapter.supports_read_path(file_path):
                    notes_parts.append(
                        f"Read not applicable: {adapter.name} does not support "
                        f"{file_path.suffix} input"
                    )
                else:
                    try:
                        read_res = _bench_read(
                            adapter=adapter,
                            test_file=test_file,
                            file_path=file_path,
                            warmup=warmup,
                            iters=iters,
                            breakdown=breakdown,
                        )
                    except Exception as e:
                        notes_parts.append(f"Read failed: {type(e).__name__}: {e}")

            if adapter.can_write():
                if "write" not in workload_ops:
                    # Workload explicitly excludes write.
                    pass
                else:
                    try:
                        write_res = _bench_write(
                            adapter=adapter,
                            test_file=test_file,
                            warmup=warmup,
                            iters=iters,
                            breakdown=breakdown,
                        )
                    except Exception as e:
                        notes_parts.append(f"Write failed: {type(e).__name__}: {e}")

            results.append(
                PerfFeatureResult(
                    feature=test_file.feature,
                    library=adapter.name,
                    workload_size=_standardize_workload_size(
                        test_file=test_file, workload=workload
                    ),
                    perf={"read": read_res, "write": write_res},
                    notes="; ".join(notes_parts) if notes_parts else None,
                )
            )

    return PerfResults(metadata=metadata, libraries=libraries, results=results)


def perf_results_to_json_dict(results: PerfResults) -> dict[str, Any]:
    return {
        "metadata": {
            "benchmark_version": results.metadata.benchmark_version,
            "run_date": results.metadata.run_date.isoformat(),
            "excel_version": results.metadata.excel_version,
            "platform": results.metadata.platform,
            "profile": results.metadata.profile,
            "python": results.metadata.python,
            "commit": results.metadata.commit,
            "config": asdict(results.metadata.config),
        },
        "libraries": results.libraries,
        "results": [_feature_result_to_dict(r) for r in results.results],
    }


def _feature_result_to_dict(r: PerfFeatureResult) -> dict[str, Any]:
    return {
        "feature": r.feature,
        "library": r.library,
        "workload_size": r.workload_size,
        "perf": {
            "read": _op_result_to_dict(r.perf.get("read")),
            "write": _op_result_to_dict(r.perf.get("write")),
        },
        "notes": r.notes,
    }


def _op_result_to_dict(op: PerfOpResult | None) -> dict[str, Any] | None:
    if op is None:
        return None
    return {
        "wall_ms": asdict(op.wall_ms),
        "cpu_ms": asdict(op.cpu_ms),
        "rss_peak_mb": op.rss_peak_mb,
        "breakdown_ms": op.breakdown_ms,
        "phase_attribution_ms": op.phase_attribution_ms,
        "op_count": op.op_count,
        "op_unit": op.op_unit,
    }


def _library_info_dict(info: Any) -> dict[str, Any]:
    return {
        "name": info.name,
        "version": info.version,
        "language": info.language,
        "capabilities": sorted(list(info.capabilities)),
    }


def _get_git_commit() -> str | None:
    import subprocess

    try:
        result = subprocess.run(
            ["git", "rev-parse", "--short", "HEAD"],
            capture_output=True,
            text=True,
            timeout=5,
        )
        if result.returncode == 0:
            return result.stdout.strip()
    except (FileNotFoundError, OSError, subprocess.SubprocessError):
        return None
    return None


def _bench_read(
    *,
    adapter: Any,
    test_file: Any,
    file_path: Path,
    warmup: int,
    iters: int,
    breakdown: bool,
) -> PerfOpResult:
    workload = _extract_single_workload(test_file)
    if workload is not None:
        return _bench_read_workload(
            adapter=adapter,
            file_path=file_path,
            warmup=warmup,
            iters=iters,
            breakdown=breakdown,
            workload=workload,
        )

    wall_samples: list[float] = []
    cpu_samples: list[float] = []
    rss_samples: list[float] = []

    phase_samples: dict[str, list[float]] = {"open": [], "sheets": [], "exercise": [], "close": []}
    attribution_samples: dict[str, list[float]] = {"parse": [], "write": [], "verify": []}

    for i in range(warmup + iters):
        m = _measure_read_iteration(
            adapter=adapter,
            test_file=test_file,
            file_path=file_path,
            breakdown=breakdown,
        )
        if i < warmup:
            continue
        wall_samples.append(m["wall_ms"])
        cpu_samples.append(m["cpu_ms"])
        if m.get("rss_peak_mb") is not None:
            rss_samples.append(float(m["rss_peak_mb"]))
        if breakdown and m.get("breakdown_ms"):
            for k, v in m["breakdown_ms"].items():
                phase_samples.setdefault(k, []).append(float(v))
        coarse = _phase_attribution_from_measurement(op_kind="read", measurement=m)
        for k, v in coarse.items():
            attribution_samples.setdefault(k, []).append(float(v))

    breakdown_out: dict[str, float] | None = None
    if breakdown:
        breakdown_out = {k: _stats(v).p50 for k, v in phase_samples.items() if v}

    return PerfOpResult(
        wall_ms=_stats(wall_samples),
        cpu_ms=_stats(cpu_samples),
        rss_peak_mb=max(rss_samples) if rss_samples else None,
        breakdown_ms=breakdown_out,
        phase_attribution_ms={k: _stats(v).p50 for k, v in attribution_samples.items() if v},
    )


def _bench_read_workload(
    *,
    adapter: Any,
    file_path: Path,
    warmup: int,
    iters: int,
    breakdown: bool,
    workload: dict[str, Any],
) -> PerfOpResult:
    cells = _cells_from_range(workload["range"])
    op_count = len(cells)

    wall_samples: list[float] = []
    cpu_samples: list[float] = []
    rss_samples: list[float] = []
    phase_samples: dict[str, list[float]] = {"open": [], "sheets": [], "exercise": [], "close": []}
    attribution_samples: dict[str, list[float]] = {"parse": [], "write": [], "verify": []}

    for i in range(warmup + iters):
        m = _measure_read_workload_iteration(
            adapter=adapter,
            file_path=file_path,
            workload=workload,
            cells=cells,
            breakdown=breakdown,
        )
        if i < warmup:
            continue
        wall_samples.append(m["wall_ms"])
        cpu_samples.append(m["cpu_ms"])
        if m.get("rss_peak_mb") is not None:
            rss_samples.append(float(m["rss_peak_mb"]))
        if breakdown and m.get("breakdown_ms"):
            for k, v in m["breakdown_ms"].items():
                phase_samples.setdefault(k, []).append(float(v))
        coarse = _phase_attribution_from_measurement(op_kind="read", measurement=m)
        for k, v in coarse.items():
            attribution_samples.setdefault(k, []).append(float(v))

    breakdown_out: dict[str, float] | None = None
    if breakdown:
        breakdown_out = {k: _stats(v).p50 for k, v in phase_samples.items() if v}

    return PerfOpResult(
        wall_ms=_stats(wall_samples),
        cpu_ms=_stats(cpu_samples),
        rss_peak_mb=max(rss_samples) if rss_samples else None,
        breakdown_ms=breakdown_out,
        phase_attribution_ms={k: _stats(v).p50 for k, v in attribution_samples.items() if v},
        op_count=op_count,
        op_unit="cells",
    )


def _measure_read_iteration(
    *,
    adapter: Any,
    test_file: Any,
    file_path: Path,
    breakdown: bool,
) -> dict[str, Any]:
    import resource
    import time

    from excelbench.harness import runner as fidelity

    rss_before = _ru_maxrss_mb(resource.getrusage(resource.RUSAGE_SELF).ru_maxrss)

    wall0 = time.perf_counter_ns()
    cpu0 = time.process_time_ns()

    phases: dict[str, float] = {}

    t0 = time.perf_counter_ns()
    workbook = adapter.open_workbook(file_path)
    t1 = time.perf_counter_ns()
    if breakdown:
        phases["open"] = _ns_to_ms(t1 - t0)

    t0 = time.perf_counter_ns()
    sheet_names = adapter.get_sheet_names(workbook)
    default_sheet = sheet_names[0] if sheet_names else test_file.feature
    t1 = time.perf_counter_ns()
    if breakdown:
        phases["sheets"] = _ns_to_ms(t1 - t0)

    t0 = time.perf_counter_ns()
    for tc in test_file.test_cases:
        _exercise_read_case(
            fidelity=fidelity,
            adapter=adapter,
            workbook=workbook,
            default_sheet=default_sheet,
            test_case=tc,
            feature=test_file.feature,
        )
    t1 = time.perf_counter_ns()
    if breakdown:
        phases["exercise"] = _ns_to_ms(t1 - t0)

    t0 = time.perf_counter_ns()
    adapter.close_workbook(workbook)
    t1 = time.perf_counter_ns()
    if breakdown:
        phases["close"] = _ns_to_ms(t1 - t0)

    wall1 = time.perf_counter_ns()
    cpu1 = time.process_time_ns()

    rss_after = _ru_maxrss_mb(resource.getrusage(resource.RUSAGE_SELF).ru_maxrss)
    rss_peak = max(rss_before, rss_after)

    return {
        "wall_ms": _ns_to_ms(wall1 - wall0),
        "cpu_ms": _ns_to_ms(cpu1 - cpu0),
        "rss_peak_mb": rss_peak,
        "breakdown_ms": phases if breakdown else None,
    }


def _exercise_read_case(
    *,
    fidelity: Any,
    adapter: Any,
    workbook: Any,
    default_sheet: str,
    test_case: Any,
    feature: str,
) -> None:
    expected = test_case.expected
    if feature == "multiple_sheets" and isinstance(expected, dict) and "sheet_names" in expected:
        fidelity.read_sheet_names_actual(adapter, workbook)
        return

    sheet = test_case.sheet or feature or default_sheet
    cell = test_case.cell or f"B{test_case.row}"

    if feature == "cell_values":
        fidelity.read_cell_value_actual(adapter, workbook, sheet, cell, expected)
    elif feature == "formulas":
        fidelity.read_formula_actual(adapter, workbook, sheet, cell)
    elif feature == "text_formatting":
        fidelity.read_text_format_actual(adapter, workbook, sheet, cell)
    elif feature == "background_colors":
        fidelity.read_background_color_actual(adapter, workbook, sheet, cell)
    elif feature == "number_formats":
        fidelity.read_number_format_actual(adapter, workbook, sheet, cell)
    elif feature == "alignment":
        fidelity.read_alignment_actual(adapter, workbook, sheet, cell)
    elif feature == "borders":
        fidelity.read_border_actual(adapter, workbook, sheet, cell)
    elif feature == "dimensions":
        fidelity.read_dimensions_actual(adapter, workbook, sheet, cell, test_case)
    elif feature == "merged_cells":
        fidelity.read_merged_cells_actual(adapter, workbook, sheet, test_case)
    elif feature == "conditional_formatting":
        fidelity.read_conditional_format_actual(adapter, workbook, sheet, expected)
    elif feature == "data_validation":
        fidelity.read_data_validation_actual(adapter, workbook, sheet, expected)
    elif feature == "hyperlinks":
        fidelity.read_hyperlink_actual(adapter, workbook, sheet, expected)
    elif feature == "images":
        fidelity.read_image_actual(adapter, workbook, sheet, expected)
    elif feature == "pivot_tables":
        fidelity.read_pivot_actual(adapter, workbook, sheet, expected)
    elif feature == "comments":
        fidelity.read_comment_actual(adapter, workbook, sheet, expected)
    elif feature == "freeze_panes":
        fidelity.read_freeze_panes_actual(adapter, workbook, sheet, expected)


def _bench_write(
    *,
    adapter: Any,
    test_file: Any,
    warmup: int,
    iters: int,
    breakdown: bool,
) -> PerfOpResult:
    workload = _extract_single_workload(test_file)
    if workload is not None:
        return _bench_write_workload(
            adapter=adapter,
            warmup=warmup,
            iters=iters,
            breakdown=breakdown,
            workload=workload,
        )

    import tempfile

    wall_samples: list[float] = []
    cpu_samples: list[float] = []
    rss_samples: list[float] = []
    phase_samples: dict[str, list[float]] = {
        "create": [],
        "add_sheets": [],
        "exercise": [],
        "save": [],
    }
    attribution_samples: dict[str, list[float]] = {"parse": [], "write": [], "verify": []}

    feature_stem = Path(test_file.feature).name or "feature"
    ext = adapter.output_extension

    with tempfile.TemporaryDirectory() as tmpdir:
        out_dir = Path(tmpdir) / adapter.name
        out_dir.mkdir(parents=True, exist_ok=True)
        out_path = out_dir / f"{feature_stem}{ext}"

        for i in range(warmup + iters):
            m = _measure_write_iteration(
                adapter=adapter,
                test_file=test_file,
                output_path=out_path,
                breakdown=breakdown,
            )
            if i < warmup:
                continue
            wall_samples.append(m["wall_ms"])
            cpu_samples.append(m["cpu_ms"])
            if m.get("rss_peak_mb") is not None:
                rss_samples.append(float(m["rss_peak_mb"]))
            if breakdown and m.get("breakdown_ms"):
                for k, v in m["breakdown_ms"].items():
                    phase_samples.setdefault(k, []).append(float(v))
            coarse = _phase_attribution_from_measurement(op_kind="write", measurement=m)
            for k, v in coarse.items():
                attribution_samples.setdefault(k, []).append(float(v))

    breakdown_out: dict[str, float] | None = None
    if breakdown:
        breakdown_out = {k: _stats(v).p50 for k, v in phase_samples.items() if v}

    return PerfOpResult(
        wall_ms=_stats(wall_samples),
        cpu_ms=_stats(cpu_samples),
        rss_peak_mb=max(rss_samples) if rss_samples else None,
        breakdown_ms=breakdown_out,
        phase_attribution_ms={k: _stats(v).p50 for k, v in attribution_samples.items() if v},
    )


def _bench_write_workload(
    *,
    adapter: Any,
    warmup: int,
    iters: int,
    breakdown: bool,
    workload: dict[str, Any],
) -> PerfOpResult:
    import tempfile

    cells = _cells_from_range(workload["range"])
    op_count = len(cells)
    if str(workload.get("op") or "") == "bulk_write_grid":
        sparse_every = workload.get("sparse_every")
        if isinstance(sparse_every, int) and sparse_every > 1:
            # Count only the filled cells for throughput reporting.
            op_count = (op_count + sparse_every - 1) // sparse_every

    wall_samples: list[float] = []
    cpu_samples: list[float] = []
    rss_samples: list[float] = []
    phase_samples: dict[str, list[float]] = {
        "create": [],
        "add_sheets": [],
        "exercise": [],
        "save": [],
    }
    attribution_samples: dict[str, list[float]] = {"parse": [], "write": [], "verify": []}

    feature_stem = Path(str(workload.get("scenario") or "workload")).name
    ext = adapter.output_extension

    with tempfile.TemporaryDirectory() as tmpdir:
        out_dir = Path(tmpdir) / adapter.name
        out_dir.mkdir(parents=True, exist_ok=True)
        out_path = out_dir / f"{feature_stem}{ext}"

        for i in range(warmup + iters):
            m = _measure_write_workload_iteration(
                adapter=adapter,
                output_path=out_path,
                workload=workload,
                cells=cells,
                breakdown=breakdown,
            )
            if i < warmup:
                continue
            wall_samples.append(m["wall_ms"])
            cpu_samples.append(m["cpu_ms"])
            if m.get("rss_peak_mb") is not None:
                rss_samples.append(float(m["rss_peak_mb"]))
            if breakdown and m.get("breakdown_ms"):
                for k, v in m["breakdown_ms"].items():
                    phase_samples.setdefault(k, []).append(float(v))
            coarse = _phase_attribution_from_measurement(op_kind="write", measurement=m)
            for k, v in coarse.items():
                attribution_samples.setdefault(k, []).append(float(v))

    breakdown_out: dict[str, float] | None = None
    if breakdown:
        breakdown_out = {k: _stats(v).p50 for k, v in phase_samples.items() if v}

    return PerfOpResult(
        wall_ms=_stats(wall_samples),
        cpu_ms=_stats(cpu_samples),
        rss_peak_mb=max(rss_samples) if rss_samples else None,
        breakdown_ms=breakdown_out,
        phase_attribution_ms={k: _stats(v).p50 for k, v in attribution_samples.items() if v},
        op_count=op_count,
        op_unit="cells",
    )


def _measure_read_workload_iteration(
    *,
    adapter: Any,
    file_path: Path,
    workload: dict[str, Any],
    cells: list[str],
    breakdown: bool,
) -> dict[str, Any]:
    import resource
    import time

    rss_before = _ru_maxrss_mb(resource.getrusage(resource.RUSAGE_SELF).ru_maxrss)

    wall0 = time.perf_counter_ns()
    cpu0 = time.process_time_ns()

    phases: dict[str, float] = {}

    t0 = time.perf_counter_ns()
    workbook = adapter.open_workbook(file_path)
    t1 = time.perf_counter_ns()
    if breakdown:
        phases["open"] = _ns_to_ms(t1 - t0)

    t0 = time.perf_counter_ns()
    adapter.get_sheet_names(workbook)
    t1 = time.perf_counter_ns()
    if breakdown:
        phases["sheets"] = _ns_to_ms(t1 - t0)

    t0 = time.perf_counter_ns()
    _run_workload_read(adapter=adapter, workbook=workbook, workload=workload, cells=cells)
    t1 = time.perf_counter_ns()
    if breakdown:
        phases["exercise"] = _ns_to_ms(t1 - t0)

    t0 = time.perf_counter_ns()
    adapter.close_workbook(workbook)
    t1 = time.perf_counter_ns()
    if breakdown:
        phases["close"] = _ns_to_ms(t1 - t0)

    wall1 = time.perf_counter_ns()
    cpu1 = time.process_time_ns()

    rss_after = _ru_maxrss_mb(resource.getrusage(resource.RUSAGE_SELF).ru_maxrss)
    rss_peak = max(rss_before, rss_after)

    return {
        "wall_ms": _ns_to_ms(wall1 - wall0),
        "cpu_ms": _ns_to_ms(cpu1 - cpu0),
        "rss_peak_mb": rss_peak,
        "breakdown_ms": phases if breakdown else None,
    }


def _measure_write_workload_iteration(
    *,
    adapter: Any,
    output_path: Path,
    workload: dict[str, Any],
    cells: list[str],
    breakdown: bool,
) -> dict[str, Any]:
    import resource
    import time

    rss_before = _ru_maxrss_mb(resource.getrusage(resource.RUSAGE_SELF).ru_maxrss)

    wall0 = time.perf_counter_ns()
    cpu0 = time.process_time_ns()

    phases: dict[str, float] = {}

    t0 = time.perf_counter_ns()
    workbook = adapter.create_workbook()
    t1 = time.perf_counter_ns()
    if breakdown:
        phases["create"] = _ns_to_ms(t1 - t0)

    sheet = str(workload.get("sheet") or "S1")
    t0 = time.perf_counter_ns()
    adapter.add_sheet(workbook, sheet)
    t1 = time.perf_counter_ns()
    if breakdown:
        phases["add_sheets"] = _ns_to_ms(t1 - t0)

    t0 = time.perf_counter_ns()
    _run_workload_write(adapter=adapter, workbook=workbook, workload=workload, cells=cells)
    t1 = time.perf_counter_ns()
    if breakdown:
        phases["exercise"] = _ns_to_ms(t1 - t0)

    t0 = time.perf_counter_ns()
    adapter.save_workbook(workbook, output_path)
    t1 = time.perf_counter_ns()
    if breakdown:
        phases["save"] = _ns_to_ms(t1 - t0)

    wall1 = time.perf_counter_ns()
    cpu1 = time.process_time_ns()

    rss_after = _ru_maxrss_mb(resource.getrusage(resource.RUSAGE_SELF).ru_maxrss)
    rss_peak = max(rss_before, rss_after)

    return {
        "wall_ms": _ns_to_ms(wall1 - wall0),
        "cpu_ms": _ns_to_ms(cpu1 - cpu0),
        "rss_peak_mb": rss_peak,
        "breakdown_ms": phases if breakdown else None,
    }


def _run_workload_read(
    *, adapter: Any, workbook: Any, workload: dict[str, Any], cells: list[str]
) -> None:
    sheet = str(workload.get("sheet") or "S1")
    op = str(workload.get("op") or "cell_value")
    if op == "cell_value":
        for cell in cells:
            adapter.read_cell_value(workbook, sheet, cell)
        return

    if op == "formula":
        for cell in cells:
            v = adapter.read_cell_value(workbook, sheet, cell)
            _ = v.formula or v.value
        return

    if op == "bulk_sheet_values":
        fn = getattr(adapter, "read_sheet_values", None)
        if fn is None:
            raise ValueError(f"Adapter does not support bulk sheet reads: {adapter.name}")

        # Prefer passing a range if supported.
        cell_range = str(workload.get("range") or "")
        try:
            data = fn(workbook, sheet, cell_range)
        except TypeError:
            data = fn(workbook, sheet)

        # Best-effort: touch output to avoid lazy containers.
        _touch = None
        if hasattr(data, "to_numpy"):
            _touch = data.to_numpy()
        elif hasattr(data, "values"):
            _touch = data.values  # pandas-style
        else:
            _touch = data
        _ = _touch
        return

    if op == "bg_color":
        for cell in cells:
            fmt = adapter.read_cell_format(workbook, sheet, cell)
            _ = fmt.bg_color
        return

    if op == "number_format":
        for cell in cells:
            fmt = adapter.read_cell_format(workbook, sheet, cell)
            _ = fmt.number_format
        return

    if op == "alignment":
        for cell in cells:
            fmt = adapter.read_cell_format(workbook, sheet, cell)
            _ = (fmt.h_align, fmt.v_align, fmt.wrap)
        return

    if op == "border":
        for cell in cells:
            border = adapter.read_cell_border(workbook, sheet, cell)
            _ = (
                getattr(border.top, "style", None),
                getattr(border.bottom, "style", None),
                getattr(border.left, "style", None),
                getattr(border.right, "style", None),
            )
        return

    raise ValueError(f"Unsupported workload op for read: {op}")


def _run_workload_write(
    *,
    adapter: Any,
    workbook: Any,
    workload: dict[str, Any],
    cells: list[str],
) -> None:
    from excelbench.models import (
        BorderEdge,
        BorderInfo,
        BorderStyle,
        CellFormat,
        CellType,
        CellValue,
    )

    sheet = str(workload.get("sheet") or "S1")
    op = str(workload.get("op") or "cell_value")

    if op == "bulk_write_grid":
        fn = getattr(adapter, "write_sheet_values", None)
        if fn is None:
            raise ValueError(f"Adapter does not support bulk sheet writes: {adapter.name}")

        start_cell, end_cell = _split_range(str(workload.get("range") or "A1"))
        r0, c0 = _cell_to_coord(start_cell)
        r1, c1 = _cell_to_coord(end_cell)
        rows = r1 - r0 + 1
        cols = c1 - c0 + 1

        value_type = str(workload.get("value_type") or "number").strip().lower()

        start = int(workload.get("start") or 1)
        step = int(workload.get("step") or 1)

        string_prefix = str(workload.get("string_prefix") or "V")
        string_mode = str(workload.get("string_mode") or "unique").strip().lower()
        string_value = str(workload.get("string_value") or "X")
        string_length_raw = workload.get("string_length")
        string_length = int(string_length_raw) if isinstance(string_length_raw, int) else None

        sparse_every = workload.get("sparse_every")
        if not isinstance(sparse_every, int) or sparse_every < 1:
            sparse_every = 1

        values: list[list[Any]] = []
        v = start
        linear_idx = 0
        for _r in range(rows):
            row_vals: list[Any] = []
            for _c in range(cols):
                filled = (linear_idx % sparse_every) == 0
                linear_idx += 1

                if not filled:
                    row_vals.append(None)
                    v += step
                    continue

                if value_type == "number":
                    row_vals.append(v)
                elif value_type == "string":
                    if string_mode == "repeated":
                        s = string_value
                    else:
                        s = f"{string_prefix}{v}"
                    if string_length is not None and string_length > 0:
                        if len(s) < string_length:
                            s = s + ("x" * (string_length - len(s)))
                        else:
                            s = s[:string_length]
                    row_vals.append(s)
                else:
                    raise ValueError(f"Unsupported bulk_write_grid value_type: {value_type}")

                v += step
            values.append(row_vals)

        fn(workbook, sheet, start_cell, values)
        return

    if op == "cell_value":
        start = int(workload.get("start") or 1)
        step = int(workload.get("step") or 1)
        value = start
        for cell in cells:
            adapter.write_cell_value(
                workbook, sheet, cell, CellValue(type=CellType.NUMBER, value=value)
            )
            value += step
        return

    if op == "formula":
        formula = str(workload.get("formula") or "=1+1")
        cell_value = CellValue(type=CellType.FORMULA, formula=formula, value=formula)
        for cell in cells:
            adapter.write_cell_value(workbook, sheet, cell, cell_value)
        return

    if op == "bg_color":
        palette = workload.get("palette")
        if not isinstance(palette, list) or not palette:
            palette = ["#FF0000", "#00FF00", "#0000FF", "#FFFF00"]

        for idx, cell in enumerate(cells):
            adapter.write_cell_value(
                workbook, sheet, cell, CellValue(type=CellType.STRING, value="Color")
            )
            color = str(palette[idx % len(palette)])
            adapter.write_cell_format(workbook, sheet, cell, CellFormat(bg_color=color))
        return

    if op == "number_format":
        number_format = str(workload.get("number_format") or "0.00")
        cell_format = CellFormat(number_format=number_format)
        for idx, cell in enumerate(cells):
            adapter.write_cell_value(
                workbook,
                sheet,
                cell,
                CellValue(type=CellType.NUMBER, value=float(idx) + 0.5),
            )
            adapter.write_cell_format(workbook, sheet, cell, cell_format)
        return

    if op == "alignment":
        h_align = str(workload.get("h_align") or "center")
        v_align = str(workload.get("v_align") or "top")
        wrap = bool(workload.get("wrap") if workload.get("wrap") is not None else True)

        cell_format = CellFormat(h_align=h_align, v_align=v_align, wrap=wrap)
        for cell in cells:
            adapter.write_cell_value(
                workbook, sheet, cell, CellValue(type=CellType.STRING, value="Align")
            )
            adapter.write_cell_format(workbook, sheet, cell, cell_format)
        return

    if op == "border":
        style = BorderStyle(str(workload.get("border_style") or "thin"))
        color = str(workload.get("border_color") or "#000000")
        edge = BorderEdge(style=style, color=color)
        border = BorderInfo(top=edge, bottom=edge, left=edge, right=edge)

        for cell in cells:
            adapter.write_cell_value(
                workbook, sheet, cell, CellValue(type=CellType.STRING, value="Border")
            )
            adapter.write_cell_border(workbook, sheet, cell, border)
        return

    raise ValueError(f"Unsupported workload op for write: {op}")


def _extract_single_workload(test_file: Any) -> dict[str, Any] | None:
    """Return workload spec if the test file is a single-workload scenario."""
    tcs = getattr(test_file, "test_cases", None)
    if not isinstance(tcs, list) or len(tcs) != 1:
        return None
    tc = tcs[0]
    expected = getattr(tc, "expected", None)
    if not isinstance(expected, dict):
        return None
    workload = expected.get("workload")
    if not isinstance(workload, dict):
        return None
    if "range" not in workload:
        return None
    return workload


def _workload_operations(workload: dict[str, Any] | None) -> set[str]:
    """Return which operations to run for a workload.

    Default is both read+write. A workload may restrict operations by specifying:

        {"operations": ["read"]}
    """
    if workload is None:
        return {"read", "write"}
    ops = workload.get("operations")
    if not isinstance(ops, list) or not ops:
        return {"read", "write"}
    out: set[str] = set()
    for op in ops:
        if not isinstance(op, str):
            continue
        op_n = op.strip().lower()
        if op_n in {"read", "write"}:
            out.add(op_n)
    return out or {"read", "write"}



def _phase_attribution_from_measurement(
    *, op_kind: str, measurement: dict[str, Any]
) -> dict[str, float]:
    breakdown = measurement.get("breakdown_ms")
    if isinstance(breakdown, dict) and breakdown:
        if op_kind == "read":
            parse = (
                float(breakdown.get("open", 0.0))
                + float(breakdown.get("sheets", 0.0))
                + float(breakdown.get("exercise", 0.0))
            )
            return {
                "parse": parse,
                "write": 0.0,
                "verify": float(breakdown.get("close", 0.0)),
            }
        if op_kind == "write":
            write = (
                float(breakdown.get("create", 0.0))
                + float(breakdown.get("add_sheets", 0.0))
                + float(breakdown.get("exercise", 0.0))
                + float(breakdown.get("save", 0.0))
            )
            return {"parse": 0.0, "write": write, "verify": 0.0}

    wall_ms = float(measurement.get("wall_ms", 0.0))
    if op_kind == "read":
        return {"parse": wall_ms, "write": 0.0, "verify": 0.0}
    if op_kind == "write":
        return {"parse": 0.0, "write": wall_ms, "verify": 0.0}
    return {"parse": 0.0, "write": 0.0, "verify": 0.0}


def _standardize_workload_size(*, test_file: Any, workload: dict[str, Any] | None) -> str:
    if workload is not None:
        try:
            op_count = len(_cells_from_range(str(workload["range"])))
        except (TypeError, ValueError, KeyError):
            op_count = 0
        return _size_from_count(op_count)

    case_count = len(getattr(test_file, "test_cases", []) or [])
    return _size_from_count(case_count)


def _size_from_count(count: int) -> str:
    if count <= 1_000:
        return "small"
    if count <= 10_000:
        return "medium"
    return "large"

def _cells_from_range(range_str: str) -> list[str]:
    start, end = _split_range(range_str)
    return _cells_in_range(start, end)


def _split_range(range_str: str) -> tuple[str, str]:
    clean = range_str.replace("$", "").upper()
    if ":" in clean:
        a, b = clean.split(":", 1)
        return a, b
    return clean, clean


def _cells_in_range(start_cell: str, end_cell: str) -> list[str]:
    start_row, start_col = _cell_to_coord(start_cell)
    end_row, end_col = _cell_to_coord(end_cell)
    cells: list[str] = []
    for r in range(start_row, end_row + 1):
        for c in range(start_col, end_col + 1):
            cells.append(_coord_to_cell(r, c))
    return cells


def _cell_to_coord(cell: str) -> tuple[int, int]:
    import re

    m = re.match(r"([A-Z]+)(\d+)$", cell.upper())
    if not m:
        raise ValueError(f"Invalid cell ref: {cell}")
    col_str, row_str = m.groups()
    col = 0
    for ch in col_str:
        col = col * 26 + (ord(ch) - ord("A") + 1)
    return int(row_str), col


def _coord_to_cell(row: int, col: int) -> str:
    letters = ""
    c = col
    while c > 0:
        c, rem = divmod(c - 1, 26)
        letters = chr(65 + rem) + letters
    return f"{letters}{row}"


def _measure_write_iteration(
    *,
    adapter: Any,
    test_file: Any,
    output_path: Path,
    breakdown: bool,
) -> dict[str, Any]:
    import resource
    import time

    from excelbench.harness import runner as fidelity

    rss_before = _ru_maxrss_mb(resource.getrusage(resource.RUSAGE_SELF).ru_maxrss)

    wall0 = time.perf_counter_ns()
    cpu0 = time.process_time_ns()

    phases: dict[str, float] = {}

    t0 = time.perf_counter_ns()
    workbook = adapter.create_workbook()
    t1 = time.perf_counter_ns()
    if breakdown:
        phases["create"] = _ns_to_ms(t1 - t0)

    t0 = time.perf_counter_ns()
    sheet_names = fidelity._collect_sheet_names(test_file)  # noqa: SLF001
    if not sheet_names:
        sheet_names = [test_file.feature]
    for name in sheet_names:
        adapter.add_sheet(workbook, name)
    t1 = time.perf_counter_ns()
    if breakdown:
        phases["add_sheets"] = _ns_to_ms(t1 - t0)

    t0 = time.perf_counter_ns()
    for tc in test_file.test_cases:
        if isinstance(tc.expected, dict) and "sheet_names" in tc.expected:
            continue
        target_sheet = tc.sheet or test_file.feature
        target_cell = tc.cell or f"B{tc.row}"
        _exercise_write_case(
            fidelity=fidelity,
            adapter=adapter,
            workbook=workbook,
            feature=test_file.feature,
            sheet=target_sheet,
            cell=target_cell,
            test_case=tc,
        )
    t1 = time.perf_counter_ns()
    if breakdown:
        phases["exercise"] = _ns_to_ms(t1 - t0)

    t0 = time.perf_counter_ns()
    adapter.save_workbook(workbook, output_path)
    t1 = time.perf_counter_ns()
    if breakdown:
        phases["save"] = _ns_to_ms(t1 - t0)

    wall1 = time.perf_counter_ns()
    cpu1 = time.process_time_ns()

    rss_after = _ru_maxrss_mb(resource.getrusage(resource.RUSAGE_SELF).ru_maxrss)
    rss_peak = max(rss_before, rss_after)

    return {
        "wall_ms": _ns_to_ms(wall1 - wall0),
        "cpu_ms": _ns_to_ms(cpu1 - cpu0),
        "rss_peak_mb": rss_peak,
        "breakdown_ms": phases if breakdown else None,
    }


def _exercise_write_case(
    *,
    fidelity: Any,
    adapter: Any,
    workbook: Any,
    feature: str,
    sheet: str,
    cell: str,
    test_case: Any,
) -> None:
    expected = test_case.expected
    if feature == "cell_values":
        fidelity._write_cell_value_case(adapter, workbook, sheet, cell, expected)  # noqa: SLF001
    elif feature == "formulas":
        fidelity._write_formula_case(adapter, workbook, sheet, cell, expected)  # noqa: SLF001
    elif feature == "text_formatting":
        fidelity._write_text_format_case(adapter, workbook, sheet, cell, test_case)  # noqa: SLF001
    elif feature == "background_colors":
        fidelity._write_background_color_case(adapter, workbook, sheet, cell, expected)  # noqa: SLF001
    elif feature == "number_formats":
        fidelity._write_number_format_case(adapter, workbook, sheet, cell, expected)  # noqa: SLF001
    elif feature == "alignment":
        fidelity._write_alignment_case(adapter, workbook, sheet, cell, expected)  # noqa: SLF001
    elif feature == "borders":
        fidelity._write_border_case(adapter, workbook, sheet, cell, expected)  # noqa: SLF001
    elif feature == "dimensions":
        fidelity._write_dimensions_case(adapter, workbook, sheet, cell, test_case)  # noqa: SLF001
    elif feature == "multiple_sheets":
        fidelity._write_multi_sheet_case(adapter, workbook, sheet, cell, expected)  # noqa: SLF001
    elif feature == "merged_cells":
        fidelity._write_merged_cells_case(adapter, workbook, sheet, expected)  # noqa: SLF001
    elif feature == "conditional_formatting":
        fidelity._write_conditional_format_case(adapter, workbook, sheet, expected)  # noqa: SLF001
    elif feature == "data_validation":
        fidelity._write_data_validation_case(adapter, workbook, sheet, expected)  # noqa: SLF001
    elif feature == "hyperlinks":
        fidelity._write_hyperlink_case(adapter, workbook, sheet, expected)  # noqa: SLF001
    elif feature == "images":
        fidelity._write_image_case(adapter, workbook, sheet, expected)  # noqa: SLF001
    elif feature == "pivot_tables":
        fidelity._write_pivot_case(adapter, workbook, sheet, expected)  # noqa: SLF001
    elif feature == "comments":
        fidelity._write_comment_case(adapter, workbook, sheet, expected)  # noqa: SLF001
    elif feature == "freeze_panes":
        fidelity._write_freeze_panes_case(adapter, workbook, sheet, expected)  # noqa: SLF001


def _stats(samples: list[float]) -> PerfStats:
    if not samples:
        raise ValueError("No samples")
    s = sorted(samples)
    return PerfStats(min=s[0], p50=_quantile_sorted(s, 0.50), p95=_quantile_sorted(s, 0.95))


def _quantile_sorted(sorted_samples: list[float], q: float) -> float:
    if not sorted_samples:
        raise ValueError("No samples")
    if q <= 0:
        return sorted_samples[0]
    if q >= 1:
        return sorted_samples[-1]
    idx = int((len(sorted_samples) - 1) * q)
    return sorted_samples[idx]


def _ns_to_ms(ns: int) -> float:
    return ns / 1_000_000.0


def _ru_maxrss_mb(ru_maxrss: float) -> float:
    import sys

    # macOS reports bytes; Linux reports kilobytes.
    if sys.platform == "darwin":
        return float(ru_maxrss) / (1024.0 * 1024.0)
    return float(ru_maxrss) / 1024.0
