"""Shared helpers for Rust-backed adapters.

These helpers centralize the dict payload contract used between PyO3 backends
and the Python harness.
"""

from __future__ import annotations

from datetime import date, datetime
from typing import Any

from excelbench.models import BorderEdge, BorderInfo, BorderStyle, CellFormat, CellType, CellValue


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


# ---------------------------------------------------------------------------
# CellFormat / BorderInfo <-> dict converters for the PyO3 boundary
# ---------------------------------------------------------------------------


def format_to_dict(fmt: CellFormat) -> dict[str, Any]:
    """Convert CellFormat to a plain dict (only non-None fields)."""
    d: dict[str, Any] = {}
    if fmt.bold is not None:
        d["bold"] = fmt.bold
    if fmt.italic is not None:
        d["italic"] = fmt.italic
    if fmt.underline is not None:
        d["underline"] = fmt.underline
    if fmt.strikethrough is not None:
        d["strikethrough"] = fmt.strikethrough
    if fmt.font_name is not None:
        d["font_name"] = fmt.font_name
    if fmt.font_size is not None:
        d["font_size"] = fmt.font_size
    if fmt.font_color is not None:
        d["font_color"] = fmt.font_color
    if fmt.bg_color is not None:
        d["bg_color"] = fmt.bg_color
    if fmt.number_format is not None:
        d["number_format"] = fmt.number_format
    if fmt.h_align is not None:
        d["h_align"] = fmt.h_align
    if fmt.v_align is not None:
        d["v_align"] = fmt.v_align
    if fmt.wrap is not None:
        d["wrap"] = fmt.wrap
    if fmt.rotation is not None:
        d["rotation"] = fmt.rotation
    if fmt.indent is not None:
        d["indent"] = fmt.indent
    return d


def dict_to_format(d: dict[str, Any]) -> CellFormat:
    """Convert a dict from the Rust backend to CellFormat."""
    return CellFormat(
        bold=d.get("bold"),
        italic=d.get("italic"),
        underline=d.get("underline"),
        strikethrough=d.get("strikethrough"),
        font_name=d.get("font_name"),
        font_size=d.get("font_size"),
        font_color=d.get("font_color"),
        bg_color=d.get("bg_color"),
        number_format=d.get("number_format"),
        h_align=d.get("h_align"),
        v_align=d.get("v_align"),
        wrap=d.get("wrap"),
        rotation=d.get("rotation"),
        indent=d.get("indent"),
    )


def _edge_to_dict(edge: BorderEdge) -> dict[str, str]:
    return {"style": edge.style.value, "color": edge.color}


def border_to_dict(border: BorderInfo) -> dict[str, Any]:
    """Convert BorderInfo to a plain dict (only non-None edges)."""
    d: dict[str, Any] = {}
    if border.top is not None:
        d["top"] = _edge_to_dict(border.top)
    if border.bottom is not None:
        d["bottom"] = _edge_to_dict(border.bottom)
    if border.left is not None:
        d["left"] = _edge_to_dict(border.left)
    if border.right is not None:
        d["right"] = _edge_to_dict(border.right)
    if border.diagonal_up is not None:
        d["diagonal_up"] = _edge_to_dict(border.diagonal_up)
    if border.diagonal_down is not None:
        d["diagonal_down"] = _edge_to_dict(border.diagonal_down)
    return d


def _dict_to_edge(d: dict[str, str]) -> BorderEdge:
    style_str = d.get("style", "none")
    try:
        style = BorderStyle(style_str)
    except ValueError:
        style = BorderStyle.NONE
    return BorderEdge(style=style, color=d.get("color", "#000000"))


def dict_to_border(d: dict[str, Any]) -> BorderInfo:
    """Convert a dict from the Rust backend to BorderInfo."""
    return BorderInfo(
        top=_dict_to_edge(d["top"]) if "top" in d else None,
        bottom=_dict_to_edge(d["bottom"]) if "bottom" in d else None,
        left=_dict_to_edge(d["left"]) if "left" in d else None,
        right=_dict_to_edge(d["right"]) if "right" in d else None,
        diagonal_up=_dict_to_edge(d["diagonal_up"]) if "diagonal_up" in d else None,
        diagonal_down=_dict_to_edge(d["diagonal_down"]) if "diagonal_down" in d else None,
    )
