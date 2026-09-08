"""Isolated mutation timing and OOXML preservation scoring."""

from __future__ import annotations

import hashlib
import json
import os
import posixpath
import re
import shutil
import statistics
import subprocess
import sys
import time
import xml.etree.ElementTree as ET
from pathlib import Path, PurePosixPath
from typing import Any
from zipfile import ZipFile

import openpyxl

from excelbench.modifiable import ModifiableEngine, Mutation, modifiable_engines

_WORKSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_DRIVER = """
import json
import sys
from pathlib import Path
from excelbench.modifiable import modifiable_engines

name, template, output, mutations = sys.argv[1:]
engine = next(item for item in modifiable_engines() if item.name == name)
engine.mutate(Path(template), Path(output), json.loads(mutations))
"""
_ELAPSED_RE = re.compile(r"^\s*(?P<seconds>\d+(?:\.\d+)?)\s+real\b", re.MULTILINE)
_RSS_RE = re.compile(r"^\s*(?P<rss>\d+)\s+maximum resident set size\b", re.MULTILINE)
_ELEMENT_CHECKS = (
    "tableParts",
    "sheetProtection",
    "dataValidations",
    "conditionalFormatting",
    "pageSetup",
    "headerFooter",
    "sheetPr",
    "pane",
    "drawing",
)


def _sha256_file(path: Path) -> str:
    """Return a SHA-256 digest for a complete file."""
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _package_parts(path: Path) -> dict[str, bytes]:
    """Read all file parts from an OOXML zip package."""
    with ZipFile(path) as archive:
        return {
            name: archive.read(name)
            for name in archive.namelist()
            if not name.endswith("/")
        }


def _xml_has_element(xml: bytes, element: str) -> bool:
    """Return whether a worksheet XML part contains the given main-namespace element."""
    try:
        root = ET.fromstring(xml)
    except ET.ParseError:
        return False
    return root.find(f".//{{{_WORKSHEET_NS}}}{element}") is not None


def _element_checklist(
    original_parts: dict[str, bytes], output_parts: dict[str, bytes]
) -> list[dict[str, Any]]:
    """Compare preservation-relevant worksheet elements part by part."""
    checks: list[dict[str, Any]] = []
    for part_name in sorted(
        name
        for name in original_parts
        if name.startswith("xl/worksheets/") and name.endswith(".xml")
    ):
        original_xml = original_parts[part_name]
        output_xml = output_parts.get(part_name, b"")
        for element in _ELEMENT_CHECKS:
            present_in_original = _xml_has_element(original_xml, element)
            present_in_output = _xml_has_element(output_xml, element)
            checks.append(
                {
                    "part": part_name,
                    "element": element,
                    "present_in_original": present_in_original,
                    "present_in_output": present_in_output,
                }
            )
    return checks


def _relationship_target(rel_part: str, target: str) -> str:
    """Resolve an internal relationship target to its package part name."""
    if target.startswith("/"):
        return target.removeprefix("/")
    rel_path = PurePosixPath(rel_part)
    owner_directory = rel_path.parent.parent
    return posixpath.normpath(posixpath.join(str(owner_directory), target))


def _dangling_relationships(parts: dict[str, bytes]) -> list[str]:
    """Return unresolved internal relationship targets declared under ``xl/``."""
    dangling: list[str] = []
    relationship_tag = f"{{{_RELATIONSHIPS_NS}}}Relationship"
    for rel_part, xml in sorted(parts.items()):
        if not rel_part.startswith("xl/") or not rel_part.endswith(".rels"):
            continue
        try:
            root = ET.fromstring(xml)
        except ET.ParseError:
            dangling.append(f"{rel_part}: invalid relationship XML")
            continue
        for relationship in root.findall(relationship_tag):
            if relationship.get("TargetMode") == "External":
                continue
            target = relationship.get("Target")
            if not target:
                dangling.append(f"{rel_part}: missing target")
                continue
            resolved = _relationship_target(rel_part, target)
            if resolved.startswith("xl/") and resolved not in parts:
                dangling.append(f"{rel_part}: {resolved}")
    return dangling


def _missing_content_types(parts: dict[str, bytes]) -> list[str]:
    """Return package parts lacking an OOXML content-type declaration."""
    content_types = parts.get("[Content_Types].xml")
    if content_types is None:
        return ["[Content_Types].xml"]
    try:
        root = ET.fromstring(content_types)
    except ET.ParseError:
        return ["[Content_Types].xml"]
    default_tag = f"{{{_CONTENT_TYPES_NS}}}Default"
    override_tag = f"{{{_CONTENT_TYPES_NS}}}Override"
    defaults = {item.get("Extension") for item in root.findall(default_tag)}
    overrides = {
        item.get("PartName", "").removeprefix("/")
        for item in root.findall(override_tag)
    }
    missing: list[str] = []
    for part_name in sorted(parts):
        if part_name == "[Content_Types].xml":
            continue
        extension = part_name.rsplit(".", 1)[-1] if "." in part_name else ""
        if part_name not in overrides and extension not in defaults:
            missing.append(part_name)
    return missing


def score_preservation(template: Path, output: Path) -> dict[str, Any]:
    """Score a mutation output against its original OOXML package structure."""
    original_parts = _package_parts(template)
    output_parts = _package_parts(output)
    original_names = set(original_parts)
    output_names = set(output_parts)
    lost_parts = sorted(original_names - output_names)
    added_parts = sorted(output_names - original_names)

    checklist = _element_checklist(original_parts, output_parts)
    expected_checks = [entry for entry in checklist if entry["present_in_original"]]
    retained_checks = [entry for entry in expected_checks if entry["present_in_output"]]
    element_score = 30.0
    if expected_checks:
        element_score *= len(retained_checks) / len(expected_checks)

    dangling_rels = _dangling_relationships(output_parts)
    missing_content_types = _missing_content_types(output_parts)
    integrity_issue_count = len(dangling_rels) + len(missing_content_types)
    integrity_score = max(0.0, 20.0 - 5.0 * integrity_issue_count)

    custom_parts = sorted(
        name for name in original_parts if name.startswith("customXml/")
    )
    custom_xml_identical = all(
        name in output_parts
        and hashlib.sha256(original_parts[name]).digest()
        == hashlib.sha256(output_parts[name]).digest()
        for name in custom_parts
    )
    custom_xml_score = 10.0 if custom_xml_identical else 0.0
    parts_score = max(0.0, 40.0 - 10.0 * len(lost_parts))
    preservation_score = round(
        parts_score + element_score + integrity_score + custom_xml_score, 2
    )

    return {
        "preservation_score": preservation_score,
        "details": {
            "lost_parts": lost_parts,
            "added_parts": added_parts,
            "element_diffs": checklist,
            "dangling_rels": dangling_rels,
            "missing_content_types": missing_content_types,
            "customXml_identical": custom_xml_identical,
        },
    }


def _parse_time_output(stderr: str) -> tuple[float | None, float | None]:
    """Extract macOS ``time -l`` elapsed seconds and RSS kilobytes."""
    elapsed_match = _ELAPSED_RE.search(stderr)
    rss_match = _RSS_RE.search(stderr)
    elapsed_seconds = float(elapsed_match.group("seconds")) if elapsed_match else None
    rss_kb = float(rss_match.group("rss")) / 1024 if rss_match else None
    return elapsed_seconds, rss_kb


def _run_subprocess_mutation(
    engine: ModifiableEngine,
    template: Path,
    output: Path,
    mutations: list[Mutation],
) -> tuple[float, float]:
    """Run one isolated mutation and return elapsed milliseconds and RSS kilobytes."""
    command = [
        "/usr/bin/time",
        "-l",
        sys.executable,
        "-c",
        _DRIVER,
        engine.name,
        str(template),
        str(output),
        json.dumps(mutations, sort_keys=True),
    ]
    completed = subprocess.run(command, text=True, capture_output=True, check=False)
    elapsed_seconds, peak_rss_kb = _parse_time_output(completed.stderr)
    if completed.returncode != 0:
        message = (
            completed.stderr.strip()
            or completed.stdout.strip()
            or "mutation driver failed"
        )
        raise RuntimeError(message)
    if elapsed_seconds is None or peak_rss_kb is None:
        raise RuntimeError(
            f"Unable to parse /usr/bin/time -l output: {completed.stderr.strip()}"
        )
    return elapsed_seconds * 1000, peak_rss_kb


def _mutations_for_engine(manifest: dict[str, Any], engine_name: str) -> list[Mutation]:
    """Materialize the fixture's engine-specific mutation values."""
    first = dict(manifest["mutation"])
    first["value"] = str(first["value"]).replace("<engine>", engine_name)
    return [first, dict(manifest["second_mutation"])]


def _mutations_landed(output: Path, mutations: list[Mutation]) -> bool:
    """Verify the requested values after a mutation output is written."""
    try:
        workbook = openpyxl.load_workbook(output, data_only=False)
    except Exception:
        return False
    try:
        return all(
            workbook[str(mutation["sheet"])][str(mutation["cell"])].value
            == mutation["value"]
            for mutation in mutations
        )
    finally:
        workbook.close()


def run_mutation_suite(
    template: Path, output_dir: Path, repeats: int = 3
) -> dict[str, Any]:
    """Benchmark isolated surgical mutations and score package preservation."""
    if repeats < 1:
        raise ValueError("repeats must be at least 1")
    template = Path(template)
    output_dir = Path(output_dir)
    if not template.is_file():
        raise FileNotFoundError(template)
    manifest_path = template.with_name("manifest.json")
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    output_dir.mkdir(parents=True, exist_ok=True)

    engine_results: dict[str, dict[str, Any]] = {}
    for engine in modifiable_engines():
        if not engine.available():
            engine_results[engine.name] = {
                "available": False,
                "status": "unavailable",
                "wall_ms_median": None,
                "peak_rss_kb_median": None,
                "preservation_score": None,
                "details": {"reason": "engine unavailable"},
            }
            continue

        engine_dir = output_dir / engine.name
        engine_dir.mkdir(parents=True, exist_ok=True)
        mutations = _mutations_for_engine(manifest, engine.name)
        wall_times: list[float] = []
        peak_rss_values: list[float] = []
        final_output: Path | None = None
        try:
            for repeat_index in range(repeats):
                input_path = engine_dir / f"repeat-{repeat_index + 1}-input.xlsx"
                output_path = engine_dir / f"repeat-{repeat_index + 1}-output.xlsx"
                shutil.copyfile(template, input_path)
                output_path.unlink(missing_ok=True)
                wall_ms, peak_rss_kb = _run_subprocess_mutation(
                    engine, input_path, output_path, mutations
                )
                wall_times.append(wall_ms)
                peak_rss_values.append(peak_rss_kb)
                final_output = output_path
            if final_output is None:
                raise RuntimeError("mutation suite did not produce an output")
            scored = score_preservation(template, final_output)
            mutations_landed = _mutations_landed(final_output, mutations)
            details = dict(scored["details"])
            details["mutations_landed"] = mutations_landed
            if not mutations_landed:
                details["integrity_failure"] = (
                    "mutated cell values did not match requested values"
                )
            engine_results[engine.name] = {
                "available": True,
                "status": "passed" if mutations_landed else "integrity-failed",
                "wall_ms_median": round(statistics.median(wall_times), 3),
                "peak_rss_kb_median": round(statistics.median(peak_rss_values), 3),
                "preservation_score": scored["preservation_score"]
                if mutations_landed
                else 0.0,
                "details": details,
            }
        except Exception as error:
            engine_results[engine.name] = {
                "available": True,
                "status": "failed",
                "wall_ms_median": None,
                "peak_rss_kb_median": None,
                "preservation_score": None,
                "details": {"error": f"{type(error).__name__}: {error}"},
            }

    return {
        "metadata": {
            "template_sha256": _sha256_file(template),
            "repeats": repeats,
            "timestamp": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
            "platform": os.uname().sysname,
        },
        "engines": engine_results,
    }
