"""Test runner for executing benchmarks."""

import os
import platform
import tempfile
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from excelbench.generator.generate import load_manifest
from excelbench.harness.adapters import (
    ExcelAdapter,
    ExcelOracleAdapter,
    OpenpyxlAdapter,
    get_all_adapters,
)
from excelbench.models import (
    BenchmarkMetadata,
    BenchmarkResults,
    BorderEdge,
    BorderInfo,
    BorderStyle,
    CellFormat,
    CellType,
    CellValue,
    FeatureScore,
    Importance,
    OperationType,
    TestCase,
    TestFile,
    TestResult,
)

BENCHMARK_VERSION = "0.1.0"


def run_benchmark(
    test_dir: Path,
    adapters: list[ExcelAdapter] | None = None,
    features: list[str] | None = None,
    profile: str = "xlsx",
) -> BenchmarkResults:
    """Run the full benchmark suite.

    Args:
        test_dir: Directory containing test files and manifest.json.
        adapters: List of adapters to test. If None, uses all available.

    Returns:
        BenchmarkResults with all scores.
    """
    test_dir = Path(test_dir)

    # Load manifest
    manifest_path = test_dir / "manifest.json"
    if not manifest_path.exists():
        raise FileNotFoundError(f"Manifest not found: {manifest_path}")

    manifest = load_manifest(manifest_path)

    # Get adapters
    if adapters is None:
        adapters = get_all_adapters()

    if features:
        normalized = {f.strip().lower() for f in features if f.strip()}
        manifest.files = [f for f in manifest.files if f.feature in normalized]
        if not manifest.files:
            missing_list = ", ".join(sorted(normalized))
            raise ValueError(f"No matching features in manifest: {missing_list}")

    # Create metadata
    metadata = BenchmarkMetadata(
        benchmark_version=BENCHMARK_VERSION,
        run_date=datetime.now(UTC),
        excel_version=manifest.excel_version,
        platform=f"{platform.system()}-{platform.machine()}",
        profile=profile,
    )

    # Collect library info
    libraries = {adapter.name: adapter.info for adapter in adapters}

    # Run tests for each file
    all_scores: list[FeatureScore] = []

    for test_file in manifest.files:
        file_path = test_dir / test_file.path

        if not file_path.exists():
            print(f"Warning: Test file not found: {file_path}")
            continue

        print(f"Testing {test_file.feature}...")

        for adapter in adapters:
            score = test_feature(
                adapter=adapter,
                test_file=test_file,
                file_path=file_path,
            )
            score = _annotate_known_limitations(score)
            if (
                test_file.feature == "pivot_tables"
                and platform.system() == "Darwin"
                and not test_file.test_cases
                and not score.notes
            ):
                score.notes = (
                    "Unsupported on macOS without a Windows-generated pivot fixture "
                    "(fixtures/excel/tier2/15_pivot_tables.xlsx)."
                )
            all_scores.append(score)
            print(f"  {adapter.name}: read={score.read_score}, write={score.write_score}")

    return BenchmarkResults(
        metadata=metadata,
        libraries=libraries,
        scores=all_scores,
    )


def test_feature(
    adapter: ExcelAdapter,
    test_file: TestFile,
    file_path: Path,
) -> FeatureScore:
    """Test a single feature with a single adapter.

    Args:
        adapter: The adapter to test.
        test_file: TestFile metadata.
        file_path: Path to the test file.

    Returns:
        FeatureScore with results.
    """
    ext = file_path.suffix.lower() or "<unknown>"
    if not adapter.supports_read_path(file_path):
        return FeatureScore(
            feature=test_file.feature,
            library=adapter.name,
            read_score=None,
            write_score=None,
            notes=f"Not applicable: {adapter.name} does not support {ext} input",
        )

    read_results: list[TestResult] = []
    write_results: list[TestResult] = []

    # Test reading
    if adapter.can_read():
        read_results = test_read(adapter, test_file, file_path)

    # Test writing
    if adapter.can_write():
        write_results = test_write(adapter, test_file, file_path)

    # Calculate scores
    read_score = calculate_score(read_results) if read_results else None
    write_score = calculate_score(write_results) if write_results else None

    return FeatureScore(
        feature=test_file.feature,
        library=adapter.name,
        read_score=read_score,
        write_score=write_score,
        test_results=read_results + write_results,
    )


def _annotate_known_limitations(score: FeatureScore) -> FeatureScore:
    limitation_notes: dict[tuple[str, str], tuple[str, str]] = {
        (
            "python-calamine",
            "alignment",
        ): (
            "read",
            "Known limitation: python-calamine alignment read is limited because "
            "its API does not expose style/alignment metadata.",
        ),
        (
            "python-calamine",
            "cell_values",
        ): (
            "read",
            "Known limitation: python-calamine can surface formula error cells as "
            "blank values in current API responses.",
        ),
        (
            "pylightxl",
            "alignment",
        ): (
            "write",
            "Known limitation: pylightxl alignment write is a no-op because the "
            "library does not support formatting writes.",
        ),
        (
            "pylightxl",
            "cell_values",
        ): (
            "write",
            "Known limitation: pylightxl cell-values write has date/boolean/error "
            "fidelity limits due to writer encoding behavior.",
        ),
    }
    key = (score.library, score.feature)
    limitation = limitation_notes.get(key)
    if limitation is None:
        return score
    side, note = limitation
    if (
        side == "read"
        and score.read_score is not None
        and score.read_score < 3
        and not score.notes
    ):
        score.notes = note
    if (
        side == "write"
        and score.write_score is not None
        and score.write_score < 3
        and not score.notes
    ):
        score.notes = note
    return score


def test_read(
    adapter: ExcelAdapter,
    test_file: TestFile,
    file_path: Path,
) -> list[TestResult]:
    """Test reading a feature.

    Args:
        adapter: The adapter to test.
        test_file: TestFile metadata.
        file_path: Path to the test file.

    Returns:
        List of TestResult for each test case.
    """
    results: list[TestResult] = []

    try:
        workbook = adapter.open_workbook(file_path)
    except Exception as e:
        # Can't even open the file
        for tc in test_file.test_cases:
            results.append(
                TestResult(
                    test_case_id=tc.id,
                    operation=OperationType.READ,
                    passed=False,
                    expected=tc.expected,
                    actual={"error": str(e)},
                    notes="Failed to open workbook",
                    importance=tc.importance,
                    label=tc.label,
                )
            )
        return results

    try:
        try:
            sheet_names = adapter.get_sheet_names(workbook)
            if not sheet_names:
                raise ValueError("No sheets found in workbook")

            default_sheet = sheet_names[0]

            for tc in test_file.test_cases:
                result = test_read_case(
                    adapter,
                    workbook,
                    default_sheet,
                    tc,
                    test_file.feature,
                    OperationType.READ,
                )
                results.append(result)
        except Exception as e:
            for tc in test_file.test_cases:
                results.append(
                    TestResult(
                        test_case_id=tc.id,
                        operation=OperationType.READ,
                        passed=False,
                        expected=tc.expected,
                        actual={"error": str(e)},
                        notes=f"Exception: {type(e).__name__}",
                        importance=tc.importance,
                        label=tc.label,
                    )
                )
    finally:
        adapter.close_workbook(workbook)

    return results


def test_read_case(
    adapter: ExcelAdapter,
    workbook: Any,
    default_sheet: str,
    test_case: TestCase,
    feature: str,
    operation: OperationType,
) -> TestResult:
    """Test reading a single test case.

    Args:
        adapter: The adapter to test.
        workbook: The opened workbook.
        sheet: Sheet name.
        test_case: The test case to verify.
        feature: The feature being tested.

    Returns:
        TestResult for this test case.
    """
    expected = test_case.expected

    if feature == "multiple_sheets" and "sheet_names" in expected:
        actual = read_sheet_names_actual(adapter, workbook)
        passed = compare_results(expected, actual)
        return TestResult(
            test_case_id=test_case.id,
            operation=operation,
            passed=passed,
            expected=expected,
            actual=actual,
            importance=test_case.importance,
            label=test_case.label,
        )

    sheet = test_case.sheet or feature or default_sheet
    cell = test_case.cell or f"B{test_case.row}"

    try:
        if feature == "cell_values":
            actual = read_cell_value_actual(adapter, workbook, sheet, cell, expected)
        elif feature == "formulas":
            actual = read_formula_actual(adapter, workbook, sheet, cell)
        elif feature == "text_formatting":
            actual = read_text_format_actual(adapter, workbook, sheet, cell)
        elif feature == "background_colors":
            actual = read_background_color_actual(adapter, workbook, sheet, cell)
        elif feature == "number_formats":
            actual = read_number_format_actual(adapter, workbook, sheet, cell)
        elif feature == "alignment":
            actual = read_alignment_actual(adapter, workbook, sheet, cell)
        elif feature == "borders":
            actual = read_border_actual(adapter, workbook, sheet, cell)
        elif feature == "dimensions":
            actual = read_dimensions_actual(adapter, workbook, sheet, cell, test_case)
        elif feature == "multiple_sheets":
            # Non-sheet_names test cases are cell value reads from specific sheets
            actual = read_cell_value_actual(adapter, workbook, sheet, cell, expected)
        elif feature == "merged_cells":
            actual = read_merged_cells_actual(adapter, workbook, sheet, test_case)
        elif feature == "conditional_formatting":
            actual = read_conditional_format_actual(adapter, workbook, sheet, expected)
        elif feature == "data_validation":
            actual = read_data_validation_actual(adapter, workbook, sheet, expected)
        elif feature == "hyperlinks":
            actual = read_hyperlink_actual(adapter, workbook, sheet, expected)
        elif feature == "images":
            actual = read_image_actual(adapter, workbook, sheet, expected)
        elif feature == "pivot_tables":
            actual = read_pivot_actual(adapter, workbook, sheet, expected)
        elif feature == "comments":
            actual = read_comment_actual(adapter, workbook, sheet, expected)
        elif feature == "freeze_panes":
            actual = read_freeze_panes_actual(adapter, workbook, sheet, expected)
        else:
            actual = {"error": f"Unknown feature: {feature}"}

        # For comparison, strip CF priority from expected (auto-assigned
        # by write libraries, not controllable) so it doesn't cause false
        # negatives.  The original expected is kept for result reporting.
        cmp_expected = expected
        if feature == "conditional_formatting":
            cmp_expected = _strip_cf_priority(expected)
        passed = compare_results(cmp_expected, actual)

        return TestResult(
            test_case_id=test_case.id,
            operation=operation,
            passed=passed,
            expected=expected,
            actual=actual,
            importance=test_case.importance,
            label=test_case.label,
        )

    except Exception as e:
        return TestResult(
            test_case_id=test_case.id,
            operation=operation,
            passed=False,
            expected=expected,
            actual={"error": str(e)},
            notes=f"Exception: {type(e).__name__}",
            importance=test_case.importance,
            label=test_case.label,
        )


def read_cell_value_actual(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    cell: str,
    expected: dict,
) -> dict:
    """Read cell value and return as comparable dict."""
    cell_value = adapter.read_cell_value(workbook, sheet, cell)

    result = {"type": cell_value.type.value}

    if cell_value.type.value != "blank":
        value = cell_value.value
        # Format dates/datetimes as strings for comparison
        if cell_value.type.value == "date" and hasattr(value, "isoformat"):
            result["value"] = value.strftime("%Y-%m-%d")
        elif cell_value.type.value == "datetime" and hasattr(value, "isoformat"):
            result["value"] = value.strftime("%Y-%m-%dT%H:%M:%S")
        else:
            result["value"] = value

    return result


def read_formula_actual(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    cell: str,
) -> dict:
    cell_value = adapter.read_cell_value(workbook, sheet, cell)
    if cell_value.type != CellType.FORMULA:
        return {"error": f"Expected formula, got {cell_value.type.value}"}
    formula = cell_value.formula or cell_value.value
    # Normalize: add single quotes around unquoted sheet names in cross-sheet refs
    # so that =References!B2 matches ='References'!B2
    formula = _normalize_sheet_quotes(formula)
    return {"type": "formula", "formula": formula}


def read_text_format_actual(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    cell: str,
) -> dict:
    """Read cell formatting and return as comparable dict."""
    fmt = adapter.read_cell_format(workbook, sheet, cell)

    result = {}
    if fmt.bold:
        result["bold"] = True
    if fmt.italic:
        result["italic"] = True
    if fmt.underline:
        result["underline"] = fmt.underline
    if fmt.strikethrough:
        result["strikethrough"] = True
    if fmt.font_name:
        result["font_name"] = fmt.font_name
    if fmt.font_size:
        result["font_size"] = fmt.font_size
    if fmt.font_color:
        result["font_color"] = fmt.font_color.upper()

    return result


def read_background_color_actual(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    cell: str,
) -> dict:
    fmt = adapter.read_cell_format(workbook, sheet, cell)
    result = {}
    if fmt.bg_color:
        result["bg_color"] = fmt.bg_color.upper()
    return result


def read_number_format_actual(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    cell: str,
) -> dict:
    fmt = adapter.read_cell_format(workbook, sheet, cell)
    result = {}
    if fmt.number_format:
        result["number_format"] = _normalize_number_format(fmt.number_format)
    return result


def _normalize_number_format(fmt: str) -> str:
    """Normalize Excel number format strings for cross-library comparison.

    Excel internally stores formats with escape characters and quoting that
    differ from the "simplified" form shown in the UI. This strips:
    - Backslash escapes for literal chars: yyyy\\-mm\\-dd → yyyy-mm-dd
    - Single-char quoted literals: "$"#,##0.00 → $#,##0.00
    - Escaped spaces: "USD"\\ 0.00 → "USD" 0.00
    """
    import re

    # Strip backslash escapes for common literal characters (-, /, ., :, space)
    result = re.sub(r"\\([-/.:\\ ])", r"\1", fmt)
    # Strip single-character quoted literals like "$" → $, but preserve
    # multi-character quoted strings like "USD"
    result = re.sub(r'"(.)"', r"\1", result)
    return result


def read_alignment_actual(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    cell: str,
) -> dict:
    fmt = adapter.read_cell_format(workbook, sheet, cell)
    result = {}
    if fmt.h_align:
        result["h_align"] = fmt.h_align
    if fmt.v_align:
        result["v_align"] = fmt.v_align
    if fmt.wrap is not None:
        result["wrap"] = fmt.wrap
    if fmt.rotation is not None:
        result["rotation"] = fmt.rotation
    if fmt.indent is not None:
        result["indent"] = fmt.indent
    # Excel defaults: h_align="general", v_align="bottom".
    # Some libraries omit defaults. Inject them when the expected
    # value would otherwise compare against an empty dict.
    if "h_align" not in result:
        result["h_align"] = "general"
    if "v_align" not in result:
        result["v_align"] = "bottom"
    return result


def read_border_actual(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    cell: str,
) -> dict:
    """Read cell border and return as comparable dict."""
    border = adapter.read_cell_border(workbook, sheet, cell)

    result = {}

    # Check for uniform border style
    styles = []
    colors = []

    for edge_name, edge in [
        ("top", border.top),
        ("bottom", border.bottom),
        ("left", border.left),
        ("right", border.right),
    ]:
        if edge and edge.style.value != "none":
            result[f"border_{edge_name}"] = edge.style.value
            result[f"border_{edge_name}_color"] = edge.color.upper()
            styles.append(edge.style.value)
            colors.append(edge.color.upper())

    # Only simplify to border_style/border_color when ALL 4 edges are present and identical
    if len(styles) == 4 and len(set(styles)) == 1:
        result["border_style"] = styles[0]
        # Remove individual entries
        for edge in ["top", "bottom", "left", "right"]:
            result.pop(f"border_{edge}", None)

    if len(colors) == 4 and len(set(colors)) == 1:
        result["border_color"] = colors[0]
        for edge in ["top", "bottom", "left", "right"]:
            result.pop(f"border_{edge}_color", None)

    # Diagonal borders
    if border.diagonal_up:
        result["border_diagonal_up"] = border.diagonal_up.style.value
    if border.diagonal_down:
        result["border_diagonal_down"] = border.diagonal_down.style.value

    return result


def read_dimensions_actual(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    cell: str,
    test_case: TestCase,
) -> dict:
    result = {}
    row = test_case.row
    column = _extract_column(cell)
    if "row_height" in test_case.expected:
        result["row_height"] = adapter.read_row_height(workbook, sheet, row)
    if "column_width" in test_case.expected:
        result["column_width"] = adapter.read_column_width(workbook, sheet, column)
    return result


def read_merged_cells_actual(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    test_case: TestCase,
) -> dict:
    expected = test_case.expected
    merged_ranges = adapter.read_merged_ranges(workbook, sheet)
    result: dict[str, Any] = {}
    expected_range = expected.get("merged_range")
    if expected_range:
        match = _find_range(merged_ranges, expected_range)
        result["merged_range"] = match
        if match:
            start_cell, end_cell = _split_range(match)
            if expected.get("top_left_value") is not None:
                result["top_left_value"] = _read_cell_scalar(adapter, workbook, sheet, start_cell)
            if expected.get("non_top_left_nonempty") is not None:
                count = 0
                for cell in _cells_in_range(start_cell, end_cell):
                    if cell == start_cell:
                        continue
                    value = adapter.read_cell_value(workbook, sheet, cell)
                    if value.type != CellType.BLANK and value.value not in (None, ""):
                        count += 1
                result["non_top_left_nonempty"] = count
            if expected.get("top_left_bg_color") is not None:
                fmt = adapter.read_cell_format(workbook, sheet, start_cell)
                result["top_left_bg_color"] = fmt.bg_color.upper() if fmt.bg_color else None
            if expected.get("non_top_left_bg_color") is not None:
                other = _first_non_top_left_cell(start_cell, end_cell)
                if other:
                    fmt = adapter.read_cell_format(workbook, sheet, other)
                    result["non_top_left_bg_color"] = fmt.bg_color.upper() if fmt.bg_color else None
    return result


def _strip_cf_priority(expected: dict) -> dict:
    """Return a copy of *expected* without ``cf_rule.priority``."""
    if "cf_rule" not in expected:
        return expected
    out = dict(expected)
    out["cf_rule"] = {k: v for k, v in out["cf_rule"].items() if k != "priority"}
    return out


def read_conditional_format_actual(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    expected: dict,
) -> dict:
    rules = adapter.read_conditional_formats(workbook, sheet)
    expected_rule = expected.get("cf_rule", expected)
    match = _find_rule(rules, expected_rule)
    if not match:
        return {}
    normalized = dict(match)
    if expected_rule.get("formula") and normalized.get("formula"):
        if _normalize_formula(expected_rule.get("formula")) == _normalize_formula(
            normalized.get("formula")
        ):
            normalized["formula"] = expected_rule.get("formula")
    projected = _project_rule(normalized, expected_rule)
    # Priority is auto-assigned by write libraries and not controllable;
    # skip it to avoid false negatives in write verification.
    projected.pop("priority", None)
    return {"cf_rule": projected}


def read_data_validation_actual(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    expected: dict,
) -> dict:
    validations = adapter.read_data_validations(workbook, sheet)
    expected_rule = expected.get("validation", expected)
    match = _find_validation(validations, expected_rule)
    if not match:
        return {}
    normalized = dict(match)
    for key in ("formula1", "formula2"):
        if expected_rule.get(key) and normalized.get(key):
            if _normalize_formula(expected_rule.get(key)) == _normalize_formula(
                normalized.get(key)
            ):
                normalized[key] = expected_rule.get(key)
    return {"validation": _project_rule(normalized, expected_rule)}


def read_hyperlink_actual(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    expected: dict,
) -> dict:
    links = adapter.read_hyperlinks(workbook, sheet)
    expected_rule = expected.get("hyperlink", expected)
    match = _find_by_key(links, "cell", expected_rule.get("cell"))
    if not match:
        return {}
    normalized = dict(match)
    if expected_rule.get("internal") and normalized.get("target"):
        target = str(normalized["target"]).lstrip("#")
        normalized["target"] = target.replace("'", "")
    return {"hyperlink": _project_rule(normalized, expected_rule)}


def read_image_actual(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    expected: dict,
) -> dict:
    images = adapter.read_images(workbook, sheet)
    expected_rule = expected.get("image", expected)
    match = _find_by_key(images, "cell", expected_rule.get("cell"))
    if not match:
        return {}
    normalized = dict(match)
    expected_path = expected_rule.get("path")
    actual_path = normalized.get("path")
    if expected_path and isinstance(actual_path, str) and actual_path.startswith("/xl/media/"):
        normalized["path"] = expected_path
    return {"image": _project_rule(normalized, expected_rule)}


def read_pivot_actual(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    expected: dict,
) -> dict:
    pivots = adapter.read_pivot_tables(workbook, sheet)
    expected_rule = expected.get("pivot", expected)
    match = _find_by_key(pivots, "name", expected_rule.get("name"))
    if not match and expected_rule.get("target_cell"):
        match = _find_by_key(pivots, "target_cell", expected_rule.get("target_cell"))
    if not match:
        return {}
    normalized = dict(match)
    if normalized.get("source_range"):
        normalized["source_range"] = (
            str(normalized["source_range"]).replace("$", "").replace("'", "")
        )
    if normalized.get("target_cell"):
        value = str(normalized["target_cell"]).replace("$", "").replace("'", "")
        if ":" in value:
            value = value.split(":", 1)[0]
        if (
            "!" not in value
            and expected_rule.get("target_cell")
            and "!" in expected_rule["target_cell"]
        ):
            value = f"{expected_rule['target_cell'].split('!', 1)[0]}!{value}"
        normalized["target_cell"] = value
    return {"pivot": _project_rule(normalized, expected_rule)}


def read_comment_actual(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    expected: dict,
) -> dict:
    comments = adapter.read_comments(workbook, sheet)
    expected_rule = expected.get("comment", expected)
    match = _find_by_key(comments, "cell", expected_rule.get("cell"))
    if not match:
        return {}
    return {"comment": _project_rule(match, expected_rule)}


def read_freeze_panes_actual(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    expected: dict,
) -> dict:
    settings = adapter.read_freeze_panes(workbook, sheet)
    expected_rule = expected.get("freeze", expected)
    return {"freeze": _project_rule(settings, expected_rule)}


def read_sheet_names_actual(adapter: ExcelAdapter, workbook: Any) -> dict:
    return {"sheet_names": adapter.get_sheet_names(workbook)}


def test_write(
    adapter: ExcelAdapter,
    test_file: TestFile,
    file_path: Path,
) -> list[TestResult]:
    """Test writing a feature.

    Creates a new file with the adapter, then reads it back with
    openpyxl to verify correctness.

    Args:
        adapter: The adapter to test.
        test_file: TestFile metadata.
        file_path: Path to the original test file (for reference).

    Returns:
        List of TestResult for each test case.
    """
    results: list[TestResult] = []

    ext = adapter.output_extension
    verifier = get_write_verifier_for_adapter(adapter, test_file.feature)

    with tempfile.TemporaryDirectory() as tmpdir:
        output_dir = Path(tmpdir) / adapter.name
        output_dir.mkdir(parents=True, exist_ok=True)
        feature_stem = Path(test_file.feature).name or "feature"
        output_path = output_dir / f"{feature_stem}{ext}"

        try:
            workbook = adapter.create_workbook()

            sheet_names = _collect_sheet_names(test_file)
            if not sheet_names:
                sheet_names = [test_file.feature]

            for name in sheet_names:
                adapter.add_sheet(workbook, name)

            for tc in test_file.test_cases:
                if "sheet_names" in tc.expected:
                    continue

                target_sheet = tc.sheet or test_file.feature
                target_cell = tc.cell or f"B{tc.row}"

                if test_file.feature == "cell_values":
                    _write_cell_value_case(
                        adapter, workbook, target_sheet, target_cell, tc.expected
                    )
                elif test_file.feature == "formulas":
                    _write_formula_case(adapter, workbook, target_sheet, target_cell, tc.expected)
                elif test_file.feature == "text_formatting":
                    _write_text_format_case(adapter, workbook, target_sheet, target_cell, tc)
                elif test_file.feature == "background_colors":
                    _write_background_color_case(
                        adapter,
                        workbook,
                        target_sheet,
                        target_cell,
                        tc.expected,
                    )
                elif test_file.feature == "number_formats":
                    _write_number_format_case(
                        adapter,
                        workbook,
                        target_sheet,
                        target_cell,
                        tc.expected,
                    )
                elif test_file.feature == "alignment":
                    _write_alignment_case(
                        adapter,
                        workbook,
                        target_sheet,
                        target_cell,
                        tc.expected,
                    )
                elif test_file.feature == "borders":
                    _write_border_case(
                        adapter,
                        workbook,
                        target_sheet,
                        target_cell,
                        tc.expected,
                    )
                elif test_file.feature == "dimensions":
                    _write_dimensions_case(adapter, workbook, target_sheet, target_cell, tc)
                elif test_file.feature == "multiple_sheets":
                    _write_multi_sheet_case(
                        adapter,
                        workbook,
                        target_sheet,
                        target_cell,
                        tc.expected,
                    )
                elif test_file.feature == "merged_cells":
                    _write_merged_cells_case(adapter, workbook, target_sheet, tc.expected)
                elif test_file.feature == "conditional_formatting":
                    _write_conditional_format_case(adapter, workbook, target_sheet, tc.expected)
                elif test_file.feature == "data_validation":
                    _write_data_validation_case(adapter, workbook, target_sheet, tc.expected)
                elif test_file.feature == "hyperlinks":
                    _write_hyperlink_case(adapter, workbook, target_sheet, tc.expected)
                elif test_file.feature == "images":
                    _write_image_case(adapter, workbook, target_sheet, tc.expected)
                elif test_file.feature == "pivot_tables":
                    _write_pivot_case(adapter, workbook, target_sheet, tc.expected)
                elif test_file.feature == "comments":
                    _write_comment_case(adapter, workbook, target_sheet, tc.expected)
                elif test_file.feature == "freeze_panes":
                    _write_freeze_panes_case(adapter, workbook, target_sheet, tc.expected)

            adapter.save_workbook(workbook, output_path)
        except Exception as e:
            for tc in test_file.test_cases:
                results.append(
                    TestResult(
                        test_case_id=tc.id,
                        operation=OperationType.WRITE,
                        passed=False,
                        expected=tc.expected,
                        actual={"error": str(e)},
                        notes=f"Write failed: {type(e).__name__}",
                        importance=tc.importance,
                        label=tc.label,
                    )
                )
            return results

        try:
            verify_wb = verifier.open_workbook(output_path)
        except Exception as e:
            for tc in test_file.test_cases:
                results.append(
                    TestResult(
                        test_case_id=tc.id,
                        operation=OperationType.WRITE,
                        passed=False,
                        expected=tc.expected,
                        actual={"error": str(e)},
                        notes="Failed to open workbook for verification",
                        importance=tc.importance,
                        label=tc.label,
                    )
                )
            return results

        try:
            try:
                sheet_names = verifier.get_sheet_names(verify_wb)
                default_sheet = sheet_names[0] if sheet_names else test_file.feature
                for tc in test_file.test_cases:
                    result = test_read_case(
                        verifier,
                        verify_wb,
                        default_sheet,
                        tc,
                        test_file.feature,
                        OperationType.WRITE,
                    )
                    results.append(result)
            except Exception as e:
                for tc in test_file.test_cases:
                    results.append(
                        TestResult(
                            test_case_id=tc.id,
                            operation=OperationType.WRITE,
                            passed=False,
                            expected=tc.expected,
                            actual={"error": str(e)},
                            notes=f"Verification failed: {type(e).__name__}",
                            importance=tc.importance,
                            label=tc.label,
                        )
                    )
        finally:
            verifier.close_workbook(verify_wb)

    return results


def compare_results(expected: dict, actual: dict) -> bool:
    """Compare expected and actual results.

    Args:
        expected: Expected values.
        actual: Actual values read.

    Returns:
        True if results match.
    """
    if "error" in actual:
        return False

    return _deep_compare(expected, actual)


def _deep_compare(expected: Any, actual: Any) -> bool:
    if isinstance(expected, dict):
        if not isinstance(actual, dict):
            return False
        for key, exp_value in expected.items():
            if key not in actual:
                if exp_value is not None:
                    return False
                continue
            if not _deep_compare(exp_value, actual[key]):
                return False
        return True

    if isinstance(expected, list):
        if not isinstance(actual, list):
            return False
        for exp_item in expected:
            if not any(_deep_compare(exp_item, act_item) for act_item in actual):
                return False
        return True

    if isinstance(expected, tuple):
        if isinstance(actual, (list, tuple)):
            return _deep_compare(list(expected), list(actual))
        return False

    if isinstance(expected, str) and expected.startswith("#"):
        if not isinstance(actual, str):
            return False
        return expected.upper() == actual.upper()

    if isinstance(expected, (int, float)) and isinstance(actual, (int, float)):
        return abs(expected - actual) <= 0.0001

    return expected == actual


def get_write_verifier() -> ExcelAdapter:
    oracle = os.environ.get("EXCELBENCH_WRITE_ORACLE", "auto").lower()
    if oracle == "openpyxl":
        return OpenpyxlAdapter()
    if oracle == "excel":
        return ExcelOracleAdapter()
    if _excel_available():
        return ExcelOracleAdapter()
    return OpenpyxlAdapter()


def get_write_verifier_for_feature(feature: str) -> ExcelAdapter:
    """Choose verifier based on feature complexity."""
    complex_features = {
        "conditional_formatting",
        "data_validation",
        "images",
        "pivot_tables",
        "comments",
        "freeze_panes",
    }
    oracle = os.environ.get("EXCELBENCH_WRITE_ORACLE", "auto").lower()
    if oracle in {"openpyxl", "excel"}:
        return get_write_verifier()
    if platform.system() == "Darwin":
        return OpenpyxlAdapter()
    if feature in complex_features and _excel_available():
        return ExcelOracleAdapter()
    return OpenpyxlAdapter()


def get_write_verifier_for_adapter(adapter: ExcelAdapter, feature: str) -> ExcelAdapter:
    """Choose verifier based on adapter output format and feature.

    .xls output must be verified with xlrd (openpyxl can't read .xls).
    .xlsx output uses the existing feature-based verifier selection.
    """
    if adapter.output_extension == ".xls":
        from excelbench.harness.adapters.xlrd_adapter import XlrdAdapter
        return XlrdAdapter()
    return get_write_verifier_for_feature(feature)


def _excel_available() -> bool:
    try:
        import xlwings as xw

        app = xw.App(visible=False, add_book=False)
        app.quit()
        return True
    except Exception:
        return False


def _collect_sheet_names(test_file: TestFile) -> list[str]:
    sheet_names: list[str] = []
    explicit = False
    for tc in test_file.test_cases:
        if "sheet_names" in tc.expected:
            sheet_names = list(tc.expected["sheet_names"])
            explicit = True
            break
        if test_file.feature == "formulas":
            sheet_names.extend(_extract_formula_sheet_names(tc.expected.get("formula", "")))
        if test_file.feature == "conditional_formatting":
            rule = tc.expected.get("cf_rule", {})
            formula = rule.get("formula")
            if formula:
                sheet_names.extend(_extract_formula_sheet_names(formula))
        if test_file.feature == "data_validation":
            rule = tc.expected.get("validation", {})
            for formula in (rule.get("formula1"), rule.get("formula2")):
                if formula:
                    sheet_names.extend(_extract_formula_sheet_names(formula))
    # Ensure the feature name is included unless sheets were explicitly listed
    if not explicit and test_file.feature not in sheet_names:
        sheet_names.insert(0, test_file.feature)
    for tc in test_file.test_cases:
        name = tc.sheet
        if name and name not in sheet_names:
            sheet_names.append(name)
    return sheet_names


def _extract_column(cell: str) -> str:
    import re

    match = re.match(r"([A-Z]+)", cell.upper())
    if not match:
        return "B"
    return match.group(1)


def _cell_to_coord(cell: str) -> tuple[int, int]:
    import re

    match = re.match(r"([A-Z]+)(\d+)", cell.upper())
    if not match:
        return 1, 1
    col_str, row_str = match.groups()
    col = 0
    for char in col_str:
        col = col * 26 + (ord(char) - ord("A") + 1)
    return int(row_str), col


def _coord_to_cell(row: int, col: int) -> str:
    letters = ""
    while col > 0:
        col, rem = divmod(col - 1, 26)
        letters = chr(65 + rem) + letters
    return f"{letters}{row}"


def _split_range(range_str: str) -> tuple[str, str]:
    clean = range_str.replace("$", "")
    if ":" in clean:
        start, end = clean.split(":", 1)
        return start, end
    return clean, clean


def _cells_in_range(start_cell: str, end_cell: str) -> list[str]:
    start_row, start_col = _cell_to_coord(start_cell)
    end_row, end_col = _cell_to_coord(end_cell)
    cells: list[str] = []
    for r in range(start_row, end_row + 1):
        for c in range(start_col, end_col + 1):
            cells.append(_coord_to_cell(r, c))
    return cells


def _first_non_top_left_cell(start_cell: str, end_cell: str) -> str | None:
    cells = _cells_in_range(start_cell, end_cell)
    return cells[1] if len(cells) > 1 else None


def _normalize_range(range_str: str) -> str:
    return range_str.replace("$", "").upper()


def _find_range(ranges: list[str], expected: str) -> str | None:
    expected_norm = _normalize_range(expected)
    for rng in ranges:
        if _normalize_range(str(rng)) == expected_norm:
            return str(rng)
    return None


def _find_by_key(items: list[dict], key: str, value: Any) -> dict | None:
    for item in items:
        if item.get(key) == value:
            return item
    return None


def _find_rule(rules: list[dict], expected: dict) -> dict | None:
    for rule in rules:
        if expected.get("range") and _normalize_range(rule.get("range", "")) != _normalize_range(
            expected.get("range", "")
        ):
            continue
        if expected.get("rule_type") and rule.get("rule_type") != expected.get("rule_type"):
            continue
        if expected.get("formula") and rule.get("formula"):
            if _normalize_formula(expected["formula"]) != _normalize_formula(rule["formula"]):
                continue
        return rule
    return None


def _find_validation(validations: list[dict], expected: dict) -> dict | None:
    for validation in validations:
        if expected.get("range") and _normalize_range(
            validation.get("range", "")
        ) != _normalize_range(expected.get("range", "")):
            continue
        if expected.get("validation_type") and validation.get("validation_type") != expected.get(
            "validation_type"
        ):
            continue
        if expected.get("formula1"):
            if _normalize_formula(validation.get("formula1")) != _normalize_formula(
                expected.get("formula1")
            ):
                continue
        return validation
    return None


def _project_rule(actual: dict, expected: dict) -> dict:
    projected: dict[str, Any] = {}
    for key in expected.keys():
        value = actual.get(key)
        if key == "path" and value is None:
            value = expected.get(key)
        projected[key] = value
    return projected


def _normalize_formula(value: Any) -> Any:
    if not isinstance(value, str):
        return value
    trimmed = value.strip()
    if trimmed.startswith("="):
        trimmed = trimmed[1:]
    if trimmed.startswith('"') and trimmed.endswith('"'):
        trimmed = trimmed[1:-1]
    return trimmed


def _normalize_sheet_quotes(formula: str) -> str:
    """Add single quotes around unquoted sheet names in cross-sheet references.

    Ensures =References!B2 is normalized to ='References'!B2 so that
    formulas from different libraries compare equal regardless of quoting.
    """
    import re

    def _quote_match(m: re.Match) -> str:
        name = m.group(1)
        cell_ref = m.group(2)
        return f"='{name}'!{cell_ref}"

    # Match =SheetName!CellRef where SheetName is not already quoted
    return re.sub(r"=([A-Za-z0-9_][A-Za-z0-9_ ]*)!(\$?[A-Z]+\$?[0-9]+)", _quote_match, formula)


def _read_cell_scalar(adapter: ExcelAdapter, workbook: Any, sheet: str, cell: str) -> Any:
    cell_value = adapter.read_cell_value(workbook, sheet, cell)
    if cell_value.type == CellType.BLANK:
        return None
    if cell_value.type == CellType.DATE and hasattr(cell_value.value, "strftime"):
        return cell_value.value.strftime("%Y-%m-%d")
    if cell_value.type == CellType.DATETIME and hasattr(cell_value.value, "strftime"):
        return cell_value.value.strftime("%Y-%m-%dT%H:%M:%S")
    return cell_value.value


def _extract_formula_sheet_names(formula: str) -> list[str]:
    import re

    names: list[str] = []
    if not formula:
        return names
    # Sheet names in formulas: 'Sheet Name'!A1 or Sheet1!A1
    for match in re.findall(r"'([^']+)'!", formula):
        names.append(match)
    for match in re.findall(r"\b([A-Za-z0-9_]+)!", formula):
        if match not in names:
            names.append(match)
    return names


def _cell_value_from_expected(expected: dict) -> CellValue:
    type_str = expected.get("type", "string")
    value = expected.get("value")
    if type_str == "blank":
        return CellValue(type=CellType.BLANK)
    if type_str == "boolean":
        return CellValue(type=CellType.BOOLEAN, value=bool(value))
    if type_str == "number":
        return CellValue(type=CellType.NUMBER, value=value)
    if type_str == "date":
        from datetime import date as _date

        if isinstance(value, str):
            value = _date.fromisoformat(value)
        return CellValue(type=CellType.DATE, value=value)
    if type_str == "datetime":
        from datetime import datetime as _datetime

        if isinstance(value, str):
            value = _datetime.fromisoformat(value)
        return CellValue(type=CellType.DATETIME, value=value)
    if type_str == "error":
        return CellValue(type=CellType.ERROR, value=value)
    if type_str == "formula":
        formula = expected.get("formula") or value
        return CellValue(type=CellType.FORMULA, value=value, formula=formula)
    return CellValue(type=CellType.STRING, value=value)


def _cell_value_from_raw(value: Any) -> CellValue:
    if value is None:
        return CellValue(type=CellType.BLANK)
    if isinstance(value, bool):
        return CellValue(type=CellType.BOOLEAN, value=value)
    if isinstance(value, (int, float)):
        return CellValue(type=CellType.NUMBER, value=value)
    return CellValue(type=CellType.STRING, value=value)


def _cell_format_from_expected(expected: dict) -> CellFormat:
    return CellFormat(
        bold=expected.get("bold"),
        italic=expected.get("italic"),
        underline=expected.get("underline"),
        strikethrough=expected.get("strikethrough"),
        font_name=expected.get("font_name"),
        font_size=expected.get("font_size"),
        font_color=expected.get("font_color"),
        bg_color=expected.get("bg_color"),
        number_format=expected.get("number_format"),
        h_align=expected.get("h_align"),
        v_align=expected.get("v_align"),
        wrap=expected.get("wrap"),
        rotation=expected.get("rotation"),
        indent=expected.get("indent"),
    )


def _border_from_expected(expected: dict) -> BorderInfo:
    default_style = expected.get("border_style")
    default_color = expected.get("border_color")
    if default_color and not default_style:
        default_style = "thin"

    def make_edge(style_key: str, color_key: str) -> BorderEdge | None:
        if style_key in expected:
            style_val = expected[style_key]
        else:
            style_val = default_style

        # If a color is specified for this edge but no style, default to "thin"
        if style_val is None and color_key in expected:
            style_val = "thin"

        if style_val is None:
            return None

        if color_key in expected:
            color_val = expected[color_key]
        else:
            color_val = default_color

        if color_val is None:
            color_val = "#000000"

        return BorderEdge(style=BorderStyle(style_val), color=color_val)

    return BorderInfo(
        top=make_edge("border_top", "border_top_color"),
        bottom=make_edge("border_bottom", "border_bottom_color"),
        left=make_edge("border_left", "border_left_color"),
        right=make_edge("border_right", "border_right_color"),
        diagonal_up=make_edge("border_diagonal_up", "border_color")
        if expected.get("border_diagonal_up") is not None
        else None,
        diagonal_down=make_edge("border_diagonal_down", "border_color")
        if expected.get("border_diagonal_down") is not None
        else None,
    )


def _write_cell_value_case(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    cell: str,
    expected: dict,
) -> None:
    cell_value = _cell_value_from_expected(expected)
    adapter.write_cell_value(workbook, sheet, cell, cell_value)

    if cell_value.type in (CellType.DATE, CellType.DATETIME):
        number_format = "yyyy-mm-dd" if cell_value.type == CellType.DATE else "yyyy-mm-dd hh:mm:ss"
        adapter.write_cell_format(workbook, sheet, cell, CellFormat(number_format=number_format))


def _write_formula_case(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    cell: str,
    expected: dict,
) -> None:
    formula = expected.get("formula")
    cell_value = CellValue(type=CellType.FORMULA, formula=formula, value=formula)
    adapter.write_cell_value(workbook, sheet, cell, cell_value)


def _write_text_format_case(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    cell: str,
    test_case: TestCase,
) -> None:
    adapter.write_cell_value(
        workbook,
        sheet,
        cell,
        CellValue(type=CellType.STRING, value=test_case.label),
    )
    adapter.write_cell_format(workbook, sheet, cell, _cell_format_from_expected(test_case.expected))


def _write_background_color_case(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    cell: str,
    expected: dict,
) -> None:
    adapter.write_cell_value(workbook, sheet, cell, CellValue(type=CellType.STRING, value="Color"))
    adapter.write_cell_format(workbook, sheet, cell, _cell_format_from_expected(expected))


def _write_number_format_case(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    cell: str,
    expected: dict,
) -> None:
    number_format = expected.get("number_format")
    value = 1234.5
    value_type = CellType.NUMBER
    if number_format and any(token in number_format for token in ["y", "m", "d"]):
        from datetime import date

        value = date(2026, 2, 4)
        value_type = CellType.DATE
    adapter.write_cell_value(workbook, sheet, cell, CellValue(type=value_type, value=value))
    adapter.write_cell_format(workbook, sheet, cell, _cell_format_from_expected(expected))


def _write_alignment_case(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    cell: str,
    expected: dict,
) -> None:
    adapter.write_cell_value(workbook, sheet, cell, CellValue(type=CellType.STRING, value="Align"))
    adapter.write_cell_format(workbook, sheet, cell, _cell_format_from_expected(expected))


def _write_border_case(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    cell: str,
    expected: dict,
) -> None:
    adapter.write_cell_value(workbook, sheet, cell, CellValue(type=CellType.STRING, value="Border"))
    adapter.write_cell_border(workbook, sheet, cell, _border_from_expected(expected))


def _write_dimensions_case(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    cell: str,
    test_case: TestCase,
) -> None:
    if "row_height" in test_case.expected:
        adapter.set_row_height(workbook, sheet, test_case.row, test_case.expected["row_height"])
    if "column_width" in test_case.expected:
        column = _extract_column(cell)
        adapter.set_column_width(workbook, sheet, column, test_case.expected["column_width"])


def _write_multi_sheet_case(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    cell: str,
    expected: dict,
) -> None:
    if "sheet_names" in expected:
        return
    if "type" in expected:
        _write_cell_value_case(adapter, workbook, sheet, cell, expected)


def _write_merged_cells_case(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    expected: dict,
) -> None:
    cell_range = expected.get("merged_range")
    if cell_range:
        start_cell, _ = _split_range(cell_range)
        if expected.get("top_left_value") is not None:
            adapter.write_cell_value(
                workbook,
                sheet,
                start_cell,
                _cell_value_from_raw(expected.get("top_left_value")),
            )
        if expected.get("top_left_bg_color") is not None:
            adapter.write_cell_format(
                workbook,
                sheet,
                start_cell,
                CellFormat(bg_color=expected.get("top_left_bg_color")),
            )
        adapter.merge_cells(workbook, sheet, cell_range)


def _write_conditional_format_case(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    expected: dict,
) -> None:
    adapter.add_conditional_format(workbook, sheet, expected)


def _write_data_validation_case(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    expected: dict,
) -> None:
    adapter.add_data_validation(workbook, sheet, expected)


def _write_hyperlink_case(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    expected: dict,
) -> None:
    adapter.add_hyperlink(workbook, sheet, expected)


def _write_image_case(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    expected: dict,
) -> None:
    adapter.add_image(workbook, sheet, expected)


def _write_pivot_case(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    expected: dict,
) -> None:
    adapter.add_pivot_table(workbook, sheet, expected)


def _write_comment_case(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    expected: dict,
) -> None:
    adapter.add_comment(workbook, sheet, expected)


def _write_freeze_panes_case(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    expected: dict,
) -> None:
    adapter.set_freeze_panes(workbook, sheet, expected)


def calculate_score(results: list[TestResult]) -> int:
    """Calculate fidelity score from test results.

    Rubric-aligned scoring:
    - 3: all basic and edge cases pass
    - 2: all basic pass, one or more edge fail
    - 1: at least one basic passes but not all basic
    - 0: no basic passes

    Args:
        results: List of test results.

    Returns:
        Score from 0-3.
    """
    if not results:
        return 0

    basic = [r for r in results if r.importance != Importance.EDGE]
    edge = [r for r in results if r.importance == Importance.EDGE]

    if not basic:
        return 0

    basic_passed = [r for r in basic if r.passed]
    if not basic_passed:
        return 0

    if len(basic_passed) == len(basic):
        if all(r.passed for r in edge):
            return 3
        return 2

    return 1
