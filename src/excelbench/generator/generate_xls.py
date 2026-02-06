"""Generate deterministic .xls benchmark fixtures via xlwt."""

import json
from datetime import UTC, date, datetime
from pathlib import Path

import xlwt

from excelbench.generator.generate import write_manifest
from excelbench.models import Importance, Manifest, TestCase, TestFile

GENERATOR_VERSION = "0.1.0-xls"


def _write_header(ws: xlwt.Worksheet) -> None:
    header_style = xlwt.easyxf(
        "font: bold on; pattern: pattern solid, fore_colour gray25;"
    )
    ws.write(0, 0, "Label", header_style)
    ws.write(0, 1, "Test Cell", header_style)
    ws.write(0, 2, "Expected", header_style)
    ws.col(0).width = int(30 * 256)
    ws.col(1).width = int(25 * 256)
    ws.col(2).width = int(50 * 256)


def _write_expected(ws: xlwt.Worksheet, row_1based: int, expected: dict) -> None:
    ws.write(row_1based - 1, 2, json.dumps(expected, ensure_ascii=False))


def _generate_cell_values(output_dir: Path) -> TestFile:
    wb = xlwt.Workbook()
    ws = wb.add_sheet("cell_values")
    _write_header(ws)

    date_style = xlwt.easyxf(num_format_str="yyyy-mm-dd")
    datetime_style = xlwt.easyxf(num_format_str="yyyy-mm-dd hh:mm:ss")

    test_cases: list[TestCase] = []

    def add_case(case_id: str, label: str, row: int, expected: dict) -> None:
        ws.write(row - 1, 0, label)
        _write_expected(ws, row, expected)
        test_cases.append(
            TestCase(
                id=case_id,
                label=label,
                row=row,
                expected=expected,
                importance=Importance.BASIC,
            )
        )

    add_case("string_simple", "String - simple", 2, {"type": "string", "value": "Hello World"})
    ws.write(1, 1, "Hello World")

    add_case("string_unicode", "String - unicode", 3, {"type": "string", "value": "日本語🎉émojis"})
    ws.write(2, 1, "日本語🎉émojis")

    add_case("string_empty", "String - empty", 4, {"type": "blank"})
    ws.write(3, 1, "")

    long_text = "A" * 1000
    add_case("string_long", "String - long (1000 chars)", 5, {"type": "string", "value": long_text})
    ws.write(4, 1, long_text)

    multiline = "Line 1\nLine 2\nLine 3"
    add_case("string_newline", "String - with newlines", 6, {"type": "string", "value": multiline})
    ws.write(5, 1, multiline)

    add_case("number_integer", "Number - integer", 7, {"type": "number", "value": 42})
    ws.write(6, 1, 42)

    add_case("number_float", "Number - float", 8, {"type": "number", "value": 3.14159265358979})
    ws.write(7, 1, 3.14159265358979)

    add_case("number_negative", "Number - negative", 9, {"type": "number", "value": -100.5})
    ws.write(8, 1, -100.5)

    add_case("number_large", "Number - large", 10, {"type": "number", "value": 1234567890123456})
    ws.write(9, 1, 1234567890123456)

    add_case("number_scientific", "Number - scientific notation", 11, {"type": "number", "value": 1.23e-10})
    ws.write(10, 1, 1.23e-10)

    add_case("date_standard", "Date - standard", 12, {"type": "date", "value": "2026-02-04"})
    ws.write(11, 1, date(2026, 2, 4), date_style)

    add_case("datetime", "DateTime - with time", 13, {"type": "datetime", "value": "2026-02-04T10:30:45"})
    ws.write(12, 1, datetime(2026, 2, 4, 10, 30, 45), datetime_style)

    add_case("boolean_true", "Boolean - TRUE", 14, {"type": "boolean", "value": True})
    ws.write(13, 1, True)

    add_case("boolean_false", "Boolean - FALSE", 15, {"type": "boolean", "value": False})
    ws.write(14, 1, False)

    # Literal error tokens avoid formula-cache blank behavior in BIFF.
    add_case("error_div0", "Error - #DIV/0!", 16, {"type": "error", "value": "#DIV/0!"})
    ws.write(15, 1, "#DIV/0!")

    add_case("error_na", "Error - #N/A", 17, {"type": "error", "value": "#N/A"})
    ws.write(16, 1, "#N/A")

    add_case("error_value", "Error - #VALUE!", 18, {"type": "error", "value": "#VALUE!"})
    ws.write(17, 1, "#VALUE!")

    add_case("blank", "Blank cell", 19, {"type": "blank"})
    # Intentionally leave B19 empty.

    path = output_dir / "tier1" / "01_cell_values.xls"
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(path))

    return TestFile(
        path=str(path.relative_to(output_dir)),
        feature="cell_values",
        tier=1,
        file_format="xls",
        test_cases=test_cases,
    )


def _generate_alignment(output_dir: Path) -> TestFile:
    wb = xlwt.Workbook()
    ws = wb.add_sheet("alignment")
    _write_header(ws)

    test_cases: list[TestCase] = []

    def add_case(case_id: str, label: str, row: int, expected: dict, style: xlwt.XFStyle, value: str) -> None:
        ws.write(row - 1, 0, label)
        ws.write(row - 1, 1, value, style)
        _write_expected(ws, row, expected)
        test_cases.append(
            TestCase(
                id=case_id,
                label=label,
                row=row,
                expected=expected,
                importance=Importance.BASIC,
            )
        )

    add_case(
        "h_left",
        "Align - left",
        2,
        {"h_align": "left"},
        xlwt.easyxf("align: horiz left"),
        "Align - left",
    )
    add_case(
        "h_center",
        "Align - center",
        3,
        {"h_align": "center"},
        xlwt.easyxf("align: horiz center"),
        "Align - center",
    )
    add_case(
        "h_right",
        "Align - right",
        4,
        {"h_align": "right"},
        xlwt.easyxf("align: horiz right"),
        "Align - right",
    )
    add_case(
        "v_top",
        "Align - top",
        5,
        {"v_align": "top"},
        xlwt.easyxf("align: vert top"),
        "Align - top",
    )
    add_case(
        "v_center",
        "Align - center",
        6,
        {"v_align": "center"},
        xlwt.easyxf("align: vert center"),
        "Align - center",
    )
    add_case(
        "v_bottom",
        "Align - bottom",
        7,
        {"v_align": "bottom"},
        xlwt.easyxf("align: vert bottom"),
        "Align - bottom",
    )
    add_case(
        "wrap_text",
        "Align - wrap text",
        8,
        {"wrap": True},
        xlwt.easyxf("align: wrap on"),
        "Line 1\nLine 2",
    )
    add_case(
        "rotation_45",
        "Align - rotation 45",
        9,
        {"rotation": 45},
        xlwt.easyxf("align: rotation 45"),
        "Rotated",
    )
    add_case(
        "indent_2",
        "Align - indent 2",
        10,
        {"indent": 2},
        xlwt.easyxf("align: indent 2"),
        "Indented",
    )

    path = output_dir / "tier1" / "06_alignment.xls"
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(path))

    return TestFile(
        path=str(path.relative_to(output_dir)),
        feature="alignment",
        tier=1,
        file_format="xls",
        test_cases=test_cases,
    )


def _generate_dimensions(output_dir: Path) -> TestFile:
    wb = xlwt.Workbook()
    ws = wb.add_sheet("dimensions")
    _write_header(ws)

    test_cases: list[TestCase] = []

    def add_case(case_id: str, label: str, row: int, expected: dict, cell: str) -> None:
        ws.write(row - 1, 0, label)
        _write_expected(ws, row, expected)
        test_cases.append(
            TestCase(
                id=case_id,
                label=label,
                row=row,
                expected=expected,
                cell=cell,
                importance=Importance.BASIC,
            )
        )

    add_case("row_height_30", "Row height - 30", 2, {"row_height": 30}, "B2")
    ws.write(1, 1, "Row height - 30")
    ws.row(1).height_mismatch = True
    ws.row(1).height = int(30 * 20)

    add_case("row_height_45", "Row height - 45", 3, {"row_height": 45}, "B3")
    ws.write(2, 1, "Row height - 45")
    ws.row(2).height_mismatch = True
    ws.row(2).height = int(45 * 20)

    add_case("col_width_20", "Column width - D = 20", 4, {"column_width": 20}, "D4")
    ws.write(3, 3, "Column width D")
    ws.col(3).width = int(20 * 256)

    add_case("col_width_8", "Column width - E = 8", 5, {"column_width": 8}, "E5")
    ws.write(4, 4, "Column width E")
    ws.col(4).width = int(8 * 256)

    path = output_dir / "tier1" / "08_dimensions.xls"
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(path))

    return TestFile(
        path=str(path.relative_to(output_dir)),
        feature="dimensions",
        tier=1,
        file_format="xls",
        test_cases=test_cases,
    )


def _generate_multiple_sheets(output_dir: Path) -> TestFile:
    wb = xlwt.Workbook()
    alpha = wb.add_sheet("Alpha")
    beta = wb.add_sheet("Beta")
    gamma = wb.add_sheet("Gamma")

    for ws in (alpha, beta, gamma):
        _write_header(ws)

    alpha.write(2, 1, "Alpha")
    beta.write(2, 1, "Beta")
    gamma.write(2, 1, "Gamma")

    test_cases = [
        TestCase(
            id="sheet_names",
            label="Sheet names",
            row=2,
            expected={"sheet_names": ["Alpha", "Beta", "Gamma"]},
            sheet="Alpha",
            importance=Importance.BASIC,
        ),
        TestCase(
            id="value_alpha",
            label="Alpha value",
            row=3,
            expected={"type": "string", "value": "Alpha"},
            sheet="Alpha",
            cell="B3",
            importance=Importance.BASIC,
        ),
        TestCase(
            id="value_beta",
            label="Beta value",
            row=3,
            expected={"type": "string", "value": "Beta"},
            sheet="Beta",
            cell="B3",
            importance=Importance.BASIC,
        ),
        TestCase(
            id="value_gamma",
            label="Gamma value",
            row=3,
            expected={"type": "string", "value": "Gamma"},
            sheet="Gamma",
            cell="B3",
            importance=Importance.BASIC,
        ),
    ]

    alpha.write(1, 0, "Sheet names")
    _write_expected(alpha, 2, {"sheet_names": ["Alpha", "Beta", "Gamma"]})
    alpha.write(2, 0, "Alpha value")
    _write_expected(alpha, 3, {"type": "string", "value": "Alpha"})
    beta.write(2, 0, "Beta value")
    _write_expected(beta, 3, {"type": "string", "value": "Beta"})
    gamma.write(2, 0, "Gamma value")
    _write_expected(gamma, 3, {"type": "string", "value": "Gamma"})

    path = output_dir / "tier1" / "09_multiple_sheets.xls"
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(path))

    return TestFile(
        path=str(path.relative_to(output_dir)),
        feature="multiple_sheets",
        tier=1,
        file_format="xls",
        test_cases=test_cases,
    )


def generate_xls(
    output_dir: Path,
    features: list[str] | None = None,
) -> Manifest:
    """Generate an .xls fixture set for format-compatible benchmarks."""
    generators = {
        "cell_values": _generate_cell_values,
        "alignment": _generate_alignment,
        "dimensions": _generate_dimensions,
        "multiple_sheets": _generate_multiple_sheets,
    }
    selected = list(generators.keys())
    if features:
        normalized = {f.strip().lower() for f in features if f.strip()}
        missing = sorted(normalized - set(generators.keys()))
        if missing:
            missing_list = ", ".join(missing)
            raise ValueError(f"Unknown .xls features: {missing_list}")
        selected = [f for f in selected if f in normalized]

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    files: list[TestFile] = []
    for feature in selected:
        test_file = generators[feature](output_dir)
        files.append(test_file)

    manifest = Manifest(
        generated_at=datetime.now(UTC),
        excel_version="xlwt",
        generator_version=GENERATOR_VERSION,
        file_format="xls",
        files=files,
    )
    write_manifest(manifest, output_dir / "manifest.json")
    return manifest

