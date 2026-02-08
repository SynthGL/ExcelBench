"""Shared helpers for Rust-backed adapters.

These helpers centralize the dict payload contract used between PyO3 backends
and the Python harness.
"""

from __future__ import annotations

from datetime import date, datetime
from typing import Any

from excelbench.models import CellType, CellValue


def get_rust_backend_version(backend_key: str) -> str:
    """Return the resolved Rust crate version for a backend (if available)."""

    try:
        import excelbench_rust

        info = excelbench_rust.build_info()
        if isinstance(info, dict):
            backend_versions = info.get("backend_versions")
            if isinstance(backend_versions, dict) and backend_versions.get(backend_key):
                return str(backend_versions[backend_key])
        return str(getattr(excelbench_rust, "__version__", "unknown"))
    except Exception:
        return "unknown"


def payload_from_cell_value(value: CellValue) -> dict[str, Any]:
    if value.type == CellType.BLANK:
        return {"type": "blank"}
    if value.type == CellType.STRING:
        return {"type": "string", "value": "" if value.value is None else str(value.value)}
    if value.type == CellType.NUMBER:
        return {"type": "number", "value": value.value}
    if value.type == CellType.BOOLEAN:
        return {"type": "boolean", "value": bool(value.value)}
    if value.type == CellType.FORMULA:
        formula = value.formula or value.value
        return {"type": "formula", "formula": str(formula), "value": str(formula)}
    if value.type == CellType.ERROR:
        return {"type": "error", "value": str(value.value)}
    if value.type == CellType.DATE:
        v = value.value
        if isinstance(v, date) and not isinstance(v, datetime):
            return {"type": "date", "value": v.isoformat()}
        if isinstance(v, str):
            return {"type": "date", "value": v}
        return {"type": "date", "value": str(v)}
    if value.type == CellType.DATETIME:
        v = value.value
        if isinstance(v, datetime):
            return {"type": "datetime", "value": v.replace(microsecond=0).isoformat()}
        if isinstance(v, str):
            return {"type": "datetime", "value": v}
        return {"type": "datetime", "value": str(v)}

    return {"type": "string", "value": "" if value.value is None else str(value.value)}


def cell_value_from_payload(payload: dict[str, Any]) -> CellValue:
    type_str = str(payload.get("type", "blank"))
    value = payload.get("value")

    if type_str == "blank":
        return CellValue(type=CellType.BLANK)
    if type_str == "string":
        return CellValue(type=CellType.STRING, value=value)
    if type_str == "number":
        return CellValue(type=CellType.NUMBER, value=value)
    if type_str == "boolean":
        return CellValue(type=CellType.BOOLEAN, value=bool(value))
    if type_str == "error":
        return CellValue(type=CellType.ERROR, value=value)
    if type_str == "formula":
        formula = payload.get("formula")
        return CellValue(type=CellType.FORMULA, value=value, formula=formula)
    if type_str == "date":
        if isinstance(value, str):
            return CellValue(type=CellType.DATE, value=date.fromisoformat(value))
        return CellValue(type=CellType.DATE, value=value)
    if type_str == "datetime":
        if isinstance(value, str):
            return CellValue(type=CellType.DATETIME, value=datetime.fromisoformat(value))
        return CellValue(type=CellType.DATETIME, value=value)

    return CellValue(type=CellType.STRING, value=str(value) if value is not None else None)
