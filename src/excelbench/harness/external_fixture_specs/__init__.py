"""Tool-specific external fixture specifications."""

from __future__ import annotations

from excelbench.harness.external_fixture_specs.apache_poi import apache_poi_fixture_specs
from excelbench.harness.external_fixture_specs.closedxml import closedxml_fixture_specs
from excelbench.harness.external_fixture_specs.excelize import excelize_fixture_specs
from excelbench.harness.external_fixture_specs.exceljs import exceljs_fixture_specs
from excelbench.harness.external_fixture_specs.npoi import npoi_fixture_specs

from .base import ExternalFixtureSpec

__all__ = [
    "ExternalFixtureSpec",
    "apache_poi_fixture_specs",
    "closedxml_fixture_specs",
    "exceljs_fixture_specs",
    "excelize_fixture_specs",
    "external_fixture_specs",
    "npoi_fixture_specs",
]


def external_fixture_specs() -> list[ExternalFixtureSpec]:
    """Return all implemented external fixture specifications."""
    return [
        *excelize_fixture_specs(),
        *closedxml_fixture_specs(),
        *npoi_fixture_specs(),
        *exceljs_fixture_specs(),
        *apache_poi_fixture_specs(),
    ]
