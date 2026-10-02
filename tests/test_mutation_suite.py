"""Tests for the surgical template-mutation benchmark."""

from __future__ import annotations

import hashlib
import shutil
import subprocess
import sys
from collections.abc import Callable
from pathlib import Path
from typing import Any
from zipfile import ZIP_DEFLATED, ZipFile

import openpyxl
import pytest

from excelbench.harness.mutation import score_preservation
from excelbench.modifiable import modifiable_engines
from excelbench.results.mutation_renderer import render_mutation_report

REPO_ROOT = Path(__file__).resolve().parents[1]
BUILDER = REPO_ROOT / "scripts" / "build_mutation_template.py"
TEMPLATE = REPO_ROOT / "fixtures" / "mutation" / "template_corporate_model.xlsx"
SOFFICE = Path(shutil.which("soffice") or "/opt/homebrew/bin/soffice")
requires_libreoffice = pytest.mark.skipif(
    not SOFFICE.is_file(), reason="LibreOffice binary not available"
)
MUTATIONS: list[dict[str, Any]] = [
    {"sheet": "Inputs", "cell": "B2", "value": "MUTATED-test"},
    {"sheet": "Inputs", "cell": "C3", "value": 1337},
]
Parts = dict[str, bytes]


def _sha256(path: Path) -> str:
    """Return the full-file SHA-256 digest."""
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _read_parts(path: Path) -> Parts:
    """Read every member of a zip package."""
    with ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def _write_parts(path: Path, parts: Parts) -> Path:
    """Write members into a new zip package."""
    with ZipFile(path, "w", compression=ZIP_DEFLATED) as archive:
        for name, payload in parts.items():
            archive.writestr(name, payload)
    return path


def _replace(parts: Parts, name: str, old: bytes, new: bytes) -> None:
    """Replace one exact byte sequence inside a part, failing if it is absent."""
    assert old in parts[name], (name, old)
    parts[name] = parts[name].replace(old, new)


def _rename(parts: Parts, old: str, new: str) -> None:
    """Move a part to a new name."""
    parts[new] = parts.pop(old)


def _correct_output() -> Parts:
    """Template parts with exactly the two declared edits applied in place."""
    parts = _read_parts(TEMPLATE)
    sheet = "xl/worksheets/sheet1.xml"
    _replace(
        parts,
        sheet,
        b'<c r="B2" t="n"><v>20</v></c>',
        b'<c r="B2" t="inlineStr"><is><t>MUTATED-test</t></is></c>',
    )
    _replace(
        parts,
        sheet,
        b'<c r="C3" t="n"><v>60</v></c>',
        b'<c r="C3" t="n"><v>1337</v></c>',
    )
    return parts


def _equivalent_rewrite(parts: Parts) -> Parts:
    """Re-serialize a correct output the way a non-copying engine might.

    Renames the comment, drawing, chart and custom XML parts, renumbers
    relationship ids, and re-encodes the custom XML with other prefixes, quoting,
    whitespace and declaration.
    """
    parts = dict(parts)
    types = "[Content_Types].xml"

    _rename(parts, "xl/comments/comment1.xml", "xl/comments1.xml")
    _replace(parts, types, b"/xl/comments/comment1.xml", b"/xl/comments1.xml")
    sheet1_rels = "xl/worksheets/_rels/sheet1.xml.rels"
    _replace(parts, sheet1_rels, b"/xl/comments/comment1.xml", b"../comments1.xml")
    _replace(parts, sheet1_rels, b'Id="comments"', b'Id="rId9"')

    _rename(parts, "xl/drawings/drawing1.xml", "xl/drawings/drawing7.xml")
    _rename(
        parts,
        "xl/drawings/_rels/drawing1.xml.rels",
        "xl/drawings/_rels/drawing7.xml.rels",
    )
    _replace(parts, types, b"/xl/drawings/drawing1.xml", b"/xl/drawings/drawing7.xml")
    sheet3_rels = "xl/worksheets/_rels/sheet3.xml.rels"
    _replace(parts, sheet3_rels, b"/xl/drawings/drawing1.xml", b"../drawings/drawing7.xml")
    _replace(parts, sheet3_rels, b'Id="rId1"', b'Id="rId42"')
    _replace(parts, "xl/worksheets/sheet3.xml", b'r:id="rId1"', b'r:id="rId42"')

    _rename(parts, "xl/charts/chart1.xml", "xl/charts/chart5.xml")
    _replace(parts, types, b"/xl/charts/chart1.xml", b"/xl/charts/chart5.xml")
    drawing_rels = "xl/drawings/_rels/drawing7.xml.rels"
    _replace(parts, drawing_rels, b"/xl/charts/chart1.xml", b"../charts/chart5.xml")
    for name in (drawing_rels, "xl/drawings/drawing7.xml"):
        parts[name] = (
            parts[name]
            .replace(b'"rId1"', b'"rIdA"')
            .replace(b'"rId2"', b'"rId1"')
            .replace(b'"rIdA"', b'"rId2"')
        )

    _rename(parts, "customXml/item1.xml", "customXml/item3.xml")
    _rename(parts, "customXml/_rels/item1.xml.rels", "customXml/_rels/item3.xml.rels")
    _rename(parts, "customXml/itemProps1.xml", "customXml/itemProps3.xml")
    _replace(parts, "_rels/.rels", b"customXml/item1.xml", b"/customXml/item3.xml")
    _replace(parts, "_rels/.rels", b'Id="rId4"', b'Id="rId11"')
    _replace(parts, types, b"/customXml/itemProps1.xml", b"/customXml/itemProps3.xml")
    parts["customXml/_rels/item3.xml.rels"] = (
        b"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/"
        b"relationships'><Relationship Target='itemProps3.xml' Id='rId5' "
        b"Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
        b"customXmlProps'/></Relationships>"
    )
    parts["customXml/item3.xml"] = (
        b"<f:p1Fidelity xmlns:f='https://wolfxl.local/fidelity'>"
        b"<f:source>openpyxl-targeted-fixture</f:source></f:p1Fidelity>"
    )
    parts["customXml/itemProps3.xml"] = (
        b"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n"
        b"<p:datastoreItem xmlns:p='http://schemas.openxmlformats.org/officeDocument/"
        b"2006/customXml' p:itemID='{11111111-2222-3333-4444-555555555555}'>"
        b"<p:schemaRefs></p:schemaRefs></p:datastoreItem>"
    )
    return parts


def _drop_chart(parts: Parts) -> None:
    del parts["xl/charts/chart1.xml"]


def _drop_comment(parts: Parts) -> None:
    del parts["xl/comments/comment1.xml"]


def _drop_custom_xml(parts: Parts) -> None:
    del parts["customXml/item1.xml"]


def test_builder_is_byte_stable() -> None:
    """Repeated generation writes the same package bytes."""
    subprocess.run([sys.executable, str(BUILDER)], cwd=REPO_ROOT, check=True)
    first_digest = _sha256(TEMPLATE)
    subprocess.run([sys.executable, str(BUILDER)], cwd=REPO_ROOT, check=True)
    assert _sha256(TEMPLATE) == first_digest


@requires_libreoffice
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


def test_equivalent_rewrite_scores_the_same_as_in_place_edit(tmp_path: Path) -> None:
    """Renamed parts, renumbered rIds and re-encoded customXml are not content changes."""
    in_place = _write_parts(tmp_path / "in_place.xlsx", _correct_output())
    rewritten = _write_parts(tmp_path / "rewritten.xlsx", _equivalent_rewrite(_correct_output()))

    in_place_score = score_preservation(TEMPLATE, in_place, MUTATIONS)
    rewritten_score = score_preservation(TEMPLATE, rewritten, MUTATIONS)

    for score in (in_place_score, rewritten_score):
        assert score["preserved"] is True
        assert score["edits_applied"] is True
        assert score["features_changed"] == {}
        assert score["integrity"] == {"dangling_rels": [], "missing_content_types": []}
    present = rewritten_score["features_present"]
    assert {"charts", "comments", "custom_xml"} <= set(present)
    assert rewritten_score["feature_preservation"] == f"{len(present)}/{len(present)}"
    assert rewritten_score["diagnostics"]["lost_parts"]  # renamed parts are only diagnostics


@pytest.mark.parametrize(
    ("drop", "feature"),
    [
        (_drop_chart, "charts"),
        (_drop_comment, "comments"),
        (_drop_custom_xml, "custom_xml"),
    ],
)
def test_dropped_content_is_reported_under_its_feature(
    tmp_path: Path, drop: Callable[[Parts], None], feature: str
) -> None:
    """Losing a chart, a comment or customXml is a change to exactly that feature."""
    parts = _correct_output()
    drop(parts)
    output = _write_parts(tmp_path / "output.xlsx", parts)

    score = score_preservation(TEMPLATE, output, MUTATIONS)

    assert score["preserved"] is False
    assert score["edits_applied"] is True
    assert set(score["features_changed"]) == {feature}
    present = len(score["features_present"])
    assert score["feature_preservation"] == f"{present - 1}/{present}"
    assert score["sample"][0]["feature"] == feature


def test_unapplied_edits_are_reported_separately(tmp_path: Path) -> None:
    """A byte copy of the template keeps every feature but misses the edits."""
    output = tmp_path / "copy.xlsx"
    shutil.copyfile(TEMPLATE, output)

    score = score_preservation(TEMPLATE, output, MUTATIONS)

    assert score["edits_applied"] is False
    assert score["preserved"] is False
    assert [edit["applied"] for edit in score["edits"]] == [False, False]
    assert score["edits"][1]["actual"] == {"type": "n", "value": 60.0}
    assert set(score["features_changed"]) == {"cells"}


def test_recomputed_formula_results_are_not_changes(tmp_path: Path) -> None:
    """Writing a cached result into a formula cell is recalculation, not a content change."""
    parts = _correct_output()
    _replace(
        parts,
        "xl/worksheets/sheet2.xml",
        b'<c r="B10"><f>SUM(B2:B9)</f><v></v></c>',
        b'<c r="B10"><f>SUM(B2:B9)</f><v>8960</v></c>',
    )
    output = _write_parts(tmp_path / "output.xlsx", parts)

    assert score_preservation(TEMPLATE, output, MUTATIONS)["preserved"] is True


def test_integrity_issues_are_reported_but_not_scored(tmp_path: Path) -> None:
    """A dangling relationship is an integrity finding, separate from preservation."""
    parts = _correct_output()
    _replace(
        parts,
        "xl/worksheets/_rels/sheet1.xml.rels",
        b"</Relationships>",
        b'<Relationship Id="rIdX" Target="/xl/missing.xml" '
        b'Type="http://example.com/unknown"/></Relationships>',
    )
    output = _write_parts(tmp_path / "output.xlsx", parts)

    score = score_preservation(TEMPLATE, output, MUTATIONS)

    assert score["integrity"]["dangling_rels"] == [
        "xl/worksheets/_rels/sheet1.xml.rels: xl/missing.xml"
    ]
    assert score["preserved"] is True


def test_engine_registry_exposes_availability_flags() -> None:
    """Every supported mutation engine reports an explicit availability flag."""
    engines = {engine.name: engine.available() for engine in modifiable_engines()}

    assert set(engines) == {
        "openpyxl",
        "wolfxl",
        "aspose-cells-foss",
        "zavora-xlsx",
        "sheetjs",
        "exceljs",
        "libreoffice",
    }
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
                "preserved": False,
                "edits_applied": True,
                "feature_preservation": "16/17",
                "features_present": [],
                "features_changed": {"custom_xml": 1},
                "details": {"integrity": {"dangling_rels": [], "missing_content_types": []}},
            }
        },
    }

    render_mutation_report(results, tmp_path)

    assert '"verdict": "Changed custom_xml"' in (tmp_path / "results.json").read_text(
        encoding="utf-8"
    )
    assert (
        "| openpyxl | 12.500 | 2048.000 | 16/17 | custom_xml (1) | yes | ok | Changed custom_xml |"
    ) in (tmp_path / "README.md").read_text(encoding="utf-8")
