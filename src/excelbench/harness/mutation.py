"""Isolated mutation timing and content-model preservation scoring."""

from __future__ import annotations

import copy
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

from excelbench.harness import content_model_xlsx as content_model
from excelbench.modifiable import (
    ModifiableEngine,
    Mutation,
    _mutation_target,
    modifiable_engines,
)

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
CUSTOM_XML_FEATURE = "custom_xml"
SCORED_FEATURES: tuple[str, ...] = (*content_model.FEATURES, CUSTOM_XML_FEATURE)
_SAMPLE_SIZE = 10
_ROW_SCORE_FIELDS = (
    "preserved",
    "edits_applied",
    "feature_preservation",
    "features_present",
    "features_changed",
)
_ROW_DETAIL_FIELDS = ("edits", "sample", "integrity", "diagnostics")

_REPO_ROOT = Path(__file__).resolve().parents[3]


def _portable_error(error: Exception) -> str:
    """Error text for published results: repository and home paths become placeholders."""
    text = f"{type(error).__name__}: {error}"
    for root, placeholder in ((_REPO_ROOT, "<repo>"), (Path.home(), "~")):
        text = text.replace(str(root), placeholder)
    return text


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
        return {name: archive.read(name) for name in archive.namelist() if not name.endswith("/")}


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
    overrides = {item.get("PartName", "").removeprefix("/") for item in root.findall(override_tag)}
    missing: list[str] = []
    for part_name in sorted(parts):
        if part_name == "[Content_Types].xml":
            continue
        extension = part_name.rsplit(".", 1)[-1] if "." in part_name else ""
        if part_name not in overrides and extension not in defaults:
            missing.append(part_name)
    return missing


def _canonical_xml(data: bytes) -> str:
    """Return C14N 2.0 text that ignores encoding, prefixes, attribute order and padding.

    The raw bytes go to the parser so the BOM and XML declaration select the decoding.
    """
    try:
        return ET.canonicalize(xml_data=data, strip_text=True, rewrite_prefixes=True)
    except ET.ParseError:
        return f"unparsable:{hashlib.sha256(data).hexdigest()}"


def _relationship_sources(package: content_model._Package) -> list[str]:
    """Return every part (``""`` for the package root) that owns a ``.rels`` part."""
    sources: list[str] = []
    for key in package.parts:
        directory, base = posixpath.split(key)
        if posixpath.basename(directory) != "_rels" or not base.endswith(".rels"):
            continue
        sources.append(posixpath.join(posixpath.dirname(directory), base[: -len(".rels")]))
    return sorted(sources)


def _custom_xml_items(path: Path) -> dict[str, dict[str, Any]]:
    """Model custom XML data items independent of part names and relationship ids.

    Items are found by relationship type from any owning part. Each item is its
    canonical XML plus the datastore item id and schema references from its
    properties part; the result is a multiset keyed by a hash of that record.
    """
    package = content_model._Package(path)
    targets = sorted(
        {
            target
            for source in _relationship_sources(package)
            for target in package.targets(source, "customXml")
        }
    )
    items: dict[str, dict[str, Any]] = {}
    for target in targets:
        record: dict[str, Any] = {"xml": _canonical_xml(package.parts[target])}
        for props_part in package.targets(target, "customXmlProps")[:1]:
            try:
                props = ET.fromstring(package.parts[props_part])
            except ET.ParseError:
                record["props"] = "unparsable"
                continue
            item_id = next(
                (value for name, value in props.attrib.items() if name.endswith("itemID")),
                None,
            )
            record["item_id"] = item_id.upper() if item_id else None
            record["schemas"] = sorted(
                element.get("uri", "")
                for element in props.iter()
                if isinstance(element.tag, str) and element.tag.endswith("schemaRef")
            )
        key = hashlib.sha256(json.dumps(record, sort_keys=True).encode()).hexdigest()[:16]
        if key in items:
            items[key]["count"] += 1
        else:
            items[key] = {**record, "count": 1}
    return items


def _compare_custom_xml(
    before: dict[str, dict[str, Any]], after: dict[str, dict[str, Any]]
) -> list[dict[str, Any]]:
    """Diff two custom XML multisets in the content model's diff-record shape."""
    diffs: list[dict[str, Any]] = []
    for key in sorted(set(before) | set(after)):
        old, new = before.get(key), after.get(key)
        if old is not None and new is not None and old["count"] == new["count"]:
            continue
        kind = "missing" if new is None else "added" if old is None else "changed"
        diffs.append(
            {
                "feature": CUSTOM_XML_FEATURE,
                "path": f"/{CUSTOM_XML_FEATURE}/{key}",
                "kind": kind,
                "before": old,
                "after": new,
            }
        )
    return diffs


def _cell_entry(value: Any) -> dict[str, Any] | None:
    """Return the content-model cell entry a correct engine writes for ``value``."""
    if value is None:
        return None
    if isinstance(value, bool):
        return {"type": "b", "value": value}
    if isinstance(value, int | float):
        return {"type": "n", "value": float(value)}
    if isinstance(value, str):
        return {"type": "s", "value": value}
    raise ValueError(f"Unsupported mutation value type: {type(value).__name__}")


def _edit_targets(
    model: dict[str, Any], mutations: list[Mutation]
) -> list[tuple[str, str, dict[str, Any] | None]]:
    """Resolve declared edits to ``(sheet, canonical address, expected cell entry)``."""
    targets: list[tuple[str, str, dict[str, Any] | None]] = []
    for mutation in mutations:
        sheet, cell, value = _mutation_target(mutation)
        if sheet not in model["cells"]:
            raise ValueError(f"Mutation sheet {sheet!r} is not a worksheet in the template")
        targets.append((sheet, cell.replace("$", "").upper(), _cell_entry(value)))
    return targets


def _apply_edits(
    model: dict[str, Any], targets: list[tuple[str, str, dict[str, Any] | None]]
) -> dict[str, Any]:
    """Return the model a correct engine produces: each edited cell holds its new
    constant value (replacing any formula there) and nothing else changes."""
    expected = copy.deepcopy(model)
    for sheet, address, entry in targets:
        cells = expected["cells"][sheet]
        if entry is None:
            cells.pop(address, None)
        else:
            cells[address] = entry
        expected["formulas"][sheet].pop(address, None)
    return expected


def _without_formula_results(
    model: dict[str, Any], formula_cells: set[tuple[str, str]]
) -> dict[str, Any]:
    """Drop cached formula results, which are recalculation state after an edit."""
    projected = dict(model)
    projected["cells"] = {
        sheet: {
            address: entry
            for address, entry in cells.items()
            if (sheet, address) not in formula_cells
        }
        for sheet, cells in model["cells"].items()
    }
    return projected


def _formula_cells(*models: dict[str, Any]) -> set[tuple[str, str]]:
    """Return every ``(sheet, address)`` holding a formula in any of ``models``."""
    return {
        (sheet, address)
        for model in models
        for sheet, formulas in model["formulas"].items()
        for address in formulas
    }


def _has_content(value: Any) -> bool:
    """Return whether a model value records anything beyond empty/false/zero defaults."""
    if isinstance(value, dict):
        return any(_has_content(item) for item in value.values())
    if isinstance(value, list):
        return any(_has_content(item) for item in value)
    return bool(value)


def score_preservation(template: Path, output: Path, mutations: list[Mutation]) -> dict[str, Any]:
    """Score a mutation output by observable workbook content, not package layout.

    The expected result is the template's content model with the declared edits
    applied the way a correct engine applies them. Any other content difference
    in the output is unexpected. Part names, relationship ids, XML serialization
    and document metadata are not content, so an equivalent re-serialization
    scores the same as a byte copy. Cached results of formula cells are
    recalculation state (an engine may keep, drop, or recompute them after an
    edit) and are excluded; formula text is compared under ``formulas``.
    """
    template_model = content_model.extract(template)
    output_model = content_model.extract(output)
    targets = _edit_targets(template_model, mutations)
    expected_model = _apply_edits(template_model, targets)

    formula_cells = _formula_cells(expected_model, output_model)
    template_custom_xml = _custom_xml_items(template)
    unexpected = content_model.compare(
        _without_formula_results(expected_model, formula_cells),
        _without_formula_results(output_model, formula_cells),
    ) + _compare_custom_xml(template_custom_xml, _custom_xml_items(output))

    edits = []
    for sheet, address, entry in targets:
        actual = output_model["cells"].get(sheet, {}).get(address)
        actual_value = (
            None if actual is None else {"type": actual["type"], "value": actual["value"]}
        )
        edits.append(
            {
                "sheet": sheet,
                "cell": address,
                "expected": entry,
                "actual": actual_value,
                "applied": actual_value == entry,
            }
        )
    edits_applied = all(edit["applied"] for edit in edits)

    template_content = {**template_model, CUSTOM_XML_FEATURE: template_custom_xml}
    features_present = [
        feature for feature in SCORED_FEATURES if _has_content(template_content[feature])
    ]
    changed_counts: dict[str, int] = {}
    for diff in unexpected:
        changed_counts[diff["feature"]] = changed_counts.get(diff["feature"], 0) + 1
    features_changed = {
        feature: changed_counts[feature] for feature in SCORED_FEATURES if feature in changed_counts
    }
    kept = [feature for feature in features_present if feature not in features_changed]

    output_parts = _package_parts(output)
    original_parts = _package_parts(template)
    return {
        "preserved": not unexpected and edits_applied,
        "edits_applied": edits_applied,
        "edits": edits,
        "features_present": features_present,
        "features_changed": features_changed,
        "feature_preservation": f"{len(kept)}/{len(features_present)}",
        "sample": unexpected[:_SAMPLE_SIZE],
        "integrity": {
            "dangling_rels": _dangling_relationships(output_parts),
            "missing_content_types": _missing_content_types(output_parts),
        },
        "diagnostics": {
            "lost_parts": sorted(set(original_parts) - set(output_parts)),
            "added_parts": sorted(set(output_parts) - set(original_parts)),
            "element_changes": [
                entry
                for entry in _element_checklist(original_parts, output_parts)
                if entry["present_in_original"] != entry["present_in_output"]
            ],
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
        message = completed.stderr.strip() or completed.stdout.strip() or "mutation driver failed"
        raise RuntimeError(message)
    if elapsed_seconds is None or peak_rss_kb is None:
        raise RuntimeError(f"Unable to parse /usr/bin/time -l output: {completed.stderr.strip()}")
    return elapsed_seconds * 1000, peak_rss_kb


def _mutations_for_engine(manifest: dict[str, Any], engine_name: str) -> list[Mutation]:
    """Materialize the fixture's engine-specific mutation values."""
    first = dict(manifest["mutation"])
    first["value"] = str(first["value"]).replace("<engine>", engine_name)
    return [first, dict(manifest["second_mutation"])]


def _engine_row(
    *,
    available: bool,
    status: str,
    wall_times: list[float] | None = None,
    peak_rss_values: list[float] | None = None,
    scored: dict[str, Any] | None = None,
    details: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Build one engine's result row; unscored rows carry ``None`` score fields."""
    row: dict[str, Any] = {
        "available": available,
        "status": status,
        "wall_ms_median": round(statistics.median(wall_times), 3) if wall_times else None,
        "peak_rss_kb_median": round(statistics.median(peak_rss_values), 3)
        if peak_rss_values
        else None,
    }
    for field in _ROW_SCORE_FIELDS:
        row[field] = scored[field] if scored is not None else None
    row["details"] = dict(details or {})
    if scored is not None:
        row["details"].update({field: scored[field] for field in _ROW_DETAIL_FIELDS})
    return row


def run_mutation_suite(template: Path, output_dir: Path, repeats: int = 3) -> dict[str, Any]:
    """Benchmark isolated surgical mutations and score content preservation."""
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
            engine_results[engine.name] = _engine_row(
                available=False,
                status="unavailable",
                details={"reason": "engine unavailable"},
            )
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
            scored = score_preservation(template, final_output, mutations)
            engine_results[engine.name] = _engine_row(
                available=True,
                status="passed" if scored["edits_applied"] else "edits-missing",
                wall_times=wall_times,
                peak_rss_values=peak_rss_values,
                scored=scored,
            )
        except Exception as error:
            engine_results[engine.name] = _engine_row(
                available=True,
                status="failed",
                details={"error": _portable_error(error)},
            )
    return {
        "metadata": {
            "template_sha256": _sha256_file(template),
            "repeats": repeats,
            "timestamp": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
            "platform": os.uname().sysname,
        },
        "engines": engine_results,
    }
