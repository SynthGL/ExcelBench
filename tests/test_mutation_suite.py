"""Tests for the surgical template-mutation benchmark."""

from __future__ import annotations

import hashlib
import subprocess
import sys
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import openpyxl

from excelbench.harness.mutation import score_preservation
from excelbench.modifiable import modifiable_engines
from excelbench.results.mutation_renderer import render_mutation_report

REPO_ROOT = Path(__file__).resolve().parents[1]
BUILDER = REPO_ROOT / "scripts" / "build_mutation_template.py"
TEMPLATE = REPO_ROOT / "fixtures" / "mutation" / "template_corporate_model.xlsx"
SOFFICE = Path("/opt/homebrew/bin/soffice")


def _sha256(path: Path) -> str:
    """Return the full-file SHA-256 digest."""
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _content_types() -> bytes:
    """Build minimum valid content type declarations for synthetic packages."""
    return (
        b"""<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">
  <Default Extension=\"rels\" """
        b"""ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>
  <Default Extension=\"xml\" ContentType=\"application/xml\"/>
</Types>"""
    )


def _relationship_xml(target: str) -> bytes:
    """Build a workbook relationship XML part targeting one worksheet."""
    return f"""<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"worksheet\" Target=\"{target}\"/>
</Relationships>""".encode()


def _write_synthetic_package(
    path: Path,
    *,
    include_custom_xml: bool = True,
    table_parts: bool = True,
    target: str = "worksheets/sheet1.xml",
) -> None:
    """Write a small package that exercises preservation scorer failure modes."""
    sheet_tail = (
        '<tableParts count="1"><tablePart r:id="rId1"/></tableParts>'
        if table_parts
        else ""
    )
    sheet_xml = (
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f"{sheet_tail}</worksheet>"
    ).encode()
    parts = {
        "[Content_Types].xml": _content_types(),
        "_rels/.rels": b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
        "xl/workbook.xml": b"<workbook/>",
        "xl/_rels/workbook.xml.rels": _relationship_xml(target),
        "xl/worksheets/sheet1.xml": sheet_xml,
        "xl/tables/table1.xml": b"<table/>",
    }
    if include_custom_xml:
        parts["customXml/item1.xml"] = b"<fixture>stable</fixture>"
    with ZipFile(path, "w", compression=ZIP_DEFLATED) as archive:
        for name, payload in parts.items():
            archive.writestr(name, payload)


def test_builder_is_byte_stable() -> None:
    """Repeated generation writes the same package bytes."""
    subprocess.run([sys.executable, str(BUILDER)], cwd=REPO_ROOT, check=True)
    first_digest = _sha256(TEMPLATE)
    subprocess.run([sys.executable, str(BUILDER)], cwd=REPO_ROOT, check=True)
    assert _sha256(TEMPLATE) == first_digest


def test_template_opens_in_openpyxl_and_libreoffice(tmp_path: Path) -> None:
    """The committed fixture parses and LibreOffice renders it to PDF."""
    workbook = openpyxl.load_workbook(TEMPLATE)
    try:
        assert workbook.sheetnames == ["Inputs", "Schedule", "Summary"]
        assert workbook["Inputs"]["B2"].value is not None
    finally:
        workbook.close()

    completed = subprocess.run(
        [
            str(SOFFICE),
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(tmp_path),
            str(TEMPLATE),
        ],
        text=True,
        capture_output=True,
        check=False,
    )
    assert completed.returncode == 0, completed.stderr
    assert (tmp_path / "template_corporate_model.pdf").is_file()


def test_preservation_scorer_reports_lost_custom_part(tmp_path: Path) -> None:
    """A dropped package part reduces the structural score."""
    template = tmp_path / "template.xlsx"
    output = tmp_path / "output.xlsx"
    _write_synthetic_package(template)
    _write_synthetic_package(output, include_custom_xml=False)

    result = score_preservation(template, output)

    assert "customXml/item1.xml" in result["details"]["lost_parts"]
    assert result["details"]["customXml_identical"] is False
    assert result["preservation_score"] < 100


def test_preservation_scorer_reports_dangling_relationship(tmp_path: Path) -> None:
    """A relationship to a missing worksheet is reported as an integrity error."""
    template = tmp_path / "template.xlsx"
    output = tmp_path / "output.xlsx"
    _write_synthetic_package(template)
    _write_synthetic_package(output, target="worksheets/missing.xml")

    result = score_preservation(template, output)

    assert result["details"]["dangling_rels"] == [
        "xl/_rels/workbook.xml.rels: xl/worksheets/missing.xml"
    ]
    assert result["preservation_score"] < 100


def test_preservation_scorer_reports_dropped_table_parts(tmp_path: Path) -> None:
    """Dropping a worksheet table declaration reduces the element score."""
    template = tmp_path / "template.xlsx"
    output = tmp_path / "output.xlsx"
    _write_synthetic_package(template)
    _write_synthetic_package(output, table_parts=False)

    result = score_preservation(template, output)

    assert any(
        entry["element"] == "tableParts"
        and entry["present_in_original"]
        and not entry["present_in_output"]
        for entry in result["details"]["element_diffs"]
    )
    assert result["preservation_score"] < 100


def test_engine_registry_exposes_availability_flags() -> None:
    """Every supported mutation engine reports an explicit availability flag."""
    engines = {engine.name: engine.available() for engine in modifiable_engines()}

    assert set(engines) == {"openpyxl", "wolfxl", "aspose-cells-foss", "zavora-xlsx"}
    assert all(isinstance(available, bool) for available in engines.values())


def test_renderer_writes_comparison_files(tmp_path: Path) -> None:
    """The renderer writes machine-readable and concise Markdown reports."""
    results = {
        "metadata": {
            "template_sha256": "abc",
            "repeats": 3,
            "timestamp": "2026-01-01T00:00:00Z",
        },
        "engines": {
            "openpyxl": {
                "status": "passed",
                "wall_ms_median": 12.5,
                "peak_rss_kb_median": 2048.0,
                "preservation_score": 90.0,
                "details": {"lost_parts": ["customXml/item1.xml"]},
            }
        },
    }

    render_mutation_report(results, tmp_path)

    assert (tmp_path / "results.json").is_file()
    assert "| openpyxl | 12.500 | 2048.000 | 90.000 | 1 | Loss detected |" in (
        tmp_path / "README.md"
    ).read_text(encoding="utf-8")
