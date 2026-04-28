"""Shared external fixture specification types."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

JSONDict = dict[str, Any]


@dataclass(frozen=True)
class ExternalFixtureSpec:
    """Definition of a local external-oracle fixture.

    Args:
        fixture_id: Stable fixture identifier.
        tool: Source helper name.
        filename: Workbook filename to generate.
        payload: JSON payload for the source helper.
        expected_parts: OOXML package parts expected in the generated workbook.
        readback_probes: Declarative checks run after WolfXL modify-save. Probe
            kinds currently include ``cell_value``, ``cell_formula``,
            ``cell_style``, ``comment_text``, ``hyperlink_target``,
            ``data_validation``, ``merged_range``, ``table_metadata``, and
            ``zip_contains``.
        notes: Human-readable reason this fixture exists.
    """

    fixture_id: str
    tool: str
    filename: str
    payload: JSONDict
    expected_parts: tuple[str, ...]
    notes: str
    readback_probes: tuple[JSONDict, ...] = ()
