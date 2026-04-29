"""wolfxl — hybrid Rust adapter: calamine (read) + rust_xlsxwriter (write).

Combines the fastest Rust Excel reader (calamine with style support) and the
fastest Rust writer (rust_xlsxwriter) into a single full-fidelity R+W adapter.
"""

import posixpath
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from excelbench.harness.adapters.base import ExcelAdapter
from excelbench.harness.adapters.rust_adapter_utils import (
    border_to_dict,
    cell_value_from_payload,
    dict_to_border,
    dict_to_format,
    format_to_dict,
    get_rust_backend_version,
    payload_from_cell_value,
    rust_xlsxwriter_row_index,
)
from excelbench.models import (
    BorderInfo,
    CellFormat,
    CellType,
    CellValue,
    LibraryInfo,
)

JSONDict = dict[str, Any]

try:
    import wolfxl._rust as _excelbench_rust
except ImportError as e:  # pragma: no cover
    raise ImportError("wolfxl._rust unavailable — wolfxl adapter requires it") from e

if getattr(_excelbench_rust, "CalamineStyledBook", None) is None:  # pragma: no cover
    raise ImportError("wolfxl._rust built without calamine backend")
if (
    getattr(_excelbench_rust, "NativeWorkbook", None) is None
    and getattr(_excelbench_rust, "RustXlsxWriterBook", None) is None
):  # pragma: no cover
    raise ImportError("wolfxl._rust built without a native writer backend")


def _writer_class(rust_module: Any) -> Any:
    """Return the current WolfXL writer class, with legacy fallback."""

    native = getattr(rust_module, "NativeWorkbook", None)
    if native is not None:
        return native
    return getattr(rust_module, "RustXlsxWriterBook")


def _read_images_from_xlsx(path: Path, sheet_name: str) -> list[JSONDict]:
    """Read image anchors from OOXML drawing parts.

    WolfXL 2.0's public read model exposes cell-level values/styles through
    CalamineStyledBook; image metadata still lives in OOXML drawing parts, so
    the benchmark adapter reads those parts directly.
    """

    ns = {
        "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    }

    def read_xml(zf: zipfile.ZipFile, name: str) -> ET.Element:
        return ET.fromstring(zf.read(name))

    def rel_targets(zf: zipfile.ZipFile, rels_path: str) -> dict[str, str]:
        root = read_xml(zf, rels_path)
        out: dict[str, str] = {}
        for rel in root.findall("rel:Relationship", ns):
            rid = rel.attrib.get("Id")
            target = rel.attrib.get("Target")
            if rid and target:
                out[rid] = target
        return out

    def resolve(base_part: str, target: str) -> str:
        if target.startswith("/"):
            return target.lstrip("/")
        return posixpath.normpath(posixpath.join(posixpath.dirname(base_part), target))

    def rels_for(part: str) -> str:
        directory, filename = posixpath.split(part)
        return posixpath.join(directory, "_rels", f"{filename}.rels")

    def cell_from_marker(marker: ET.Element) -> str:
        col = int(marker.findtext("xdr:col", "0", ns)) + 1
        row = int(marker.findtext("xdr:row", "0", ns)) + 1
        letters = ""
        while col:
            col, rem = divmod(col - 1, 26)
            letters = chr(65 + rem) + letters
        return f"{letters}{row}"

    with zipfile.ZipFile(path) as zf:
        workbook = read_xml(zf, "xl/workbook.xml")
        workbook_rels = rel_targets(zf, "xl/_rels/workbook.xml.rels")
        sheet_part: str | None = None
        for sheet in workbook.findall("main:sheets/main:sheet", ns):
            if sheet.attrib.get("name") != sheet_name:
                continue
            rid = sheet.attrib.get(f"{{{ns['r']}}}id")
            target = workbook_rels.get(str(rid)) if rid else None
            if target:
                sheet_part = resolve("xl/workbook.xml", target)
                break
        if sheet_part is None:
            return []

        try:
            sheet_rels = rel_targets(zf, rels_for(sheet_part))
        except KeyError:
            return []

        images: list[JSONDict] = []
        for drawing_target in sheet_rels.values():
            if "drawing" not in drawing_target:
                continue
            drawing_part = resolve(sheet_part, drawing_target)
            drawing = read_xml(zf, drawing_part)
            try:
                drawing_rels = rel_targets(zf, rels_for(drawing_part))
            except KeyError:
                drawing_rels = {}

            for anchor_tag, anchor_name in (
                ("xdr:oneCellAnchor", "oneCell"),
                ("xdr:twoCellAnchor", "twoCell"),
            ):
                for anchor in drawing.findall(anchor_tag, ns):
                    marker = anchor.find("xdr:from", ns)
                    blip = anchor.find(".//a:blip", ns)
                    if marker is None or blip is None:
                        continue
                    embed = blip.attrib.get(f"{{{ns['r']}}}embed")
                    media_target = drawing_rels.get(str(embed)) if embed else None
                    if not media_target:
                        continue
                    media_part = resolve(drawing_part, media_target)
                    images.append(
                        {
                            "cell": cell_from_marker(marker),
                            "path": f"/{media_part}",
                            "anchor": anchor_name,
                        }
                    )
        return images


class WolfxlAdapter(ExcelAdapter):
    """Hybrid adapter: calamine-styled reads + rust_xlsxwriter writes."""

    def __init__(self) -> None:
        # Python-side cell cache: avoids FFI on repeated reads of the same cell.
        # Keyed by (workbook_id, sheet, cell) → CellValue.
        self._cell_cache: dict[tuple[int, str, str], CellValue] = {}
        self._workbook_paths: dict[int, Path] = {}

    @property
    def info(self) -> LibraryInfo:
        cal_ver = get_rust_backend_version("calamine")
        rxw_ver = get_rust_backend_version("rust_xlsxwriter")
        if rxw_ver == "unknown":
            rxw_ver = get_rust_backend_version("native")
        return LibraryInfo(
            name="wolfxl",
            version=f"cal={cal_ver}+rxw={rxw_ver}",
            language="rust",
            capabilities={"read", "write", "modify"},
        )

    @property
    def supported_read_extensions(self) -> set[str]:
        return {".xlsx"}

    # =========================================================================
    # Read — delegates to CalamineStyledBook
    # =========================================================================

    def open_workbook(self, path: Path) -> Any:
        import wolfxl._rust as rust

        m: Any = rust
        workbook = getattr(m, "CalamineStyledBook").open(str(path))
        self._workbook_paths[id(workbook)] = path
        return workbook

    def close_workbook(self, workbook: Any) -> None:
        # Evict cached cells for this workbook.
        wb_id = id(workbook)
        self._cell_cache = {k: v for k, v in self._cell_cache.items() if k[0] != wb_id}
        self._workbook_paths.pop(wb_id, None)

    def get_sheet_names(self, workbook: Any) -> list[str]:
        return [str(name) for name in workbook.sheet_names()]

    def read_cell_value(self, workbook: Any, sheet: str, cell: str) -> CellValue:
        key = (id(workbook), sheet, cell)
        cached = self._cell_cache.get(key)
        if cached is not None:
            return cached
        payload = workbook.read_cell_value(sheet, cell)
        if not isinstance(payload, dict):
            result = CellValue(type=CellType.STRING, value=str(payload))
        else:
            result = cell_value_from_payload(payload)
        self._cell_cache[key] = result
        return result

    def read_sheet_values(
        self,
        workbook: Any,
        sheet: str,
        cell_range: str | None = None,
    ) -> list[list[CellValue]]:
        """Bulk read all values from a sheet via CalamineStyledBook.read_sheet_values()."""
        raw = workbook.read_sheet_values(sheet, cell_range)
        return [
            [
                cell_value_from_payload(v)
                if isinstance(v, dict)
                else CellValue(type=CellType.BLANK)
                for v in row
            ]
            for row in raw
        ]

    def read_sheet_values_raw(
        self,
        workbook: Any,
        sheet: str,
        cell_range: str | None = None,
    ) -> list[list[Any]]:
        """Return raw Rust FFI output without cell_value_from_payload() wrapping."""
        result: list[list[Any]] = workbook.read_sheet_values(sheet, cell_range)
        return result

    def read_cell_format(self, workbook: Any, sheet: str, cell: str) -> CellFormat:
        payload = workbook.read_cell_format(sheet, cell)
        if not isinstance(payload, dict) or not payload:
            return CellFormat()
        return dict_to_format(payload)

    def read_cell_border(self, workbook: Any, sheet: str, cell: str) -> BorderInfo:
        payload = workbook.read_cell_border(sheet, cell)
        if not isinstance(payload, dict) or not payload:
            return BorderInfo()
        return dict_to_border(payload)

    def read_row_height(self, workbook: Any, sheet: str, row: int) -> float | None:
        value = workbook.read_row_height(sheet, row)
        if value is None:
            return None
        if isinstance(value, (int, float)):
            return float(value)
        return None

    def read_column_width(self, workbook: Any, sheet: str, column: str) -> float | None:
        value = workbook.read_column_width(sheet, column)
        if value is None:
            return None
        if isinstance(value, (int, float)):
            return float(value)
        return None

    def read_merged_ranges(self, workbook: Any, sheet: str) -> list[str]:
        result = workbook.read_merged_ranges(sheet)
        if isinstance(result, list):
            return [str(x) for x in result]
        return []

    def read_conditional_formats(self, workbook: Any, sheet: str) -> list[JSONDict]:
        result = workbook.read_conditional_formats(sheet)
        if isinstance(result, list):
            return [dict(x) for x in result if isinstance(x, dict)]
        return []

    def read_data_validations(self, workbook: Any, sheet: str) -> list[JSONDict]:
        result = workbook.read_data_validations(sheet)
        if isinstance(result, list):
            return [dict(x) for x in result if isinstance(x, dict)]
        return []

    def read_named_ranges(self, workbook: Any, sheet: str) -> list[JSONDict]:
        result = workbook.read_named_ranges(sheet)
        if isinstance(result, list):
            return [dict(x) for x in result if isinstance(x, dict)]
        return []

    def read_tables(self, workbook: Any, sheet: str) -> list[JSONDict]:
        result = workbook.read_tables(sheet)
        if isinstance(result, list):
            return [dict(x) for x in result if isinstance(x, dict)]
        return []

    def read_hyperlinks(self, workbook: Any, sheet: str) -> list[JSONDict]:
        result = workbook.read_hyperlinks(sheet)
        if isinstance(result, list):
            return [dict(x) for x in result if isinstance(x, dict)]
        return []

    def read_images(self, workbook: Any, sheet: str) -> list[JSONDict]:
        path = self._workbook_paths.get(id(workbook))
        if path is None:
            return []
        try:
            return _read_images_from_xlsx(path, sheet)
        except Exception:
            return []

    def read_pivot_tables(self, workbook: Any, sheet: str) -> list[JSONDict]:
        return []

    def read_comments(self, workbook: Any, sheet: str) -> list[JSONDict]:
        result = workbook.read_comments(sheet)
        if isinstance(result, list):
            return [dict(x) for x in result if isinstance(x, dict)]
        return []

    def read_freeze_panes(self, workbook: Any, sheet: str) -> JSONDict:
        return dict(workbook.read_freeze_panes(sheet))

    # =========================================================================
    # Write — delegates to NativeWorkbook (WolfXL 2.0+) or RustXlsxWriterBook
    # on older WolfXL releases.
    # =========================================================================

    def create_workbook(self) -> Any:
        import wolfxl._rust as rust

        m: Any = rust
        return _writer_class(m)()

    def add_sheet(self, workbook: Any, name: str) -> None:
        workbook.add_sheet(name)

    def write_cell_value(self, workbook: Any, sheet: str, cell: str, value: CellValue) -> None:
        payload = payload_from_cell_value(value)
        workbook.write_cell_value(sheet, cell, payload)

    def write_sheet_values(
        self,
        workbook: Any,
        sheet: str,
        start_cell: str,
        values: list[list[Any]],
    ) -> None:
        """Bulk write a grid of values via the native writer."""
        workbook.write_sheet_values(sheet, start_cell, values)

    def write_cell_format(self, workbook: Any, sheet: str, cell: str, format: CellFormat) -> None:
        d = format_to_dict(format)
        if d:
            workbook.write_cell_format(sheet, cell, d)

    def write_sheet_formats(
        self,
        workbook: Any,
        sheet: str,
        start_cell: str,
        formats: list[list[dict[str, Any] | None]],
    ) -> None:
        """Bulk write a grid of format dicts via the native writer."""
        workbook.write_sheet_formats(sheet, start_cell, formats)

    def write_cell_border(self, workbook: Any, sheet: str, cell: str, border: BorderInfo) -> None:
        d = border_to_dict(border)
        if d:
            workbook.write_cell_border(sheet, cell, d)

    def write_sheet_borders(
        self,
        workbook: Any,
        sheet: str,
        start_cell: str,
        borders: list[list[dict[str, Any] | None]],
    ) -> None:
        """Bulk write a grid of border dicts via the native writer."""
        workbook.write_sheet_borders(sheet, start_cell, borders)

    def set_row_height(self, workbook: Any, sheet: str, row: int, height: float) -> None:
        workbook.set_row_height(sheet, rust_xlsxwriter_row_index(row), height)

    def set_column_width(self, workbook: Any, sheet: str, column: str, width: float) -> None:
        workbook.set_column_width(sheet, column, width)

    def merge_cells(self, workbook: Any, sheet: str, cell_range: str) -> None:
        workbook.merge_cells(sheet, cell_range)

    def add_conditional_format(self, workbook: Any, sheet: str, rule: JSONDict) -> None:
        workbook.add_conditional_format(sheet, rule.get("cf_rule", rule))

    def add_data_validation(self, workbook: Any, sheet: str, validation: JSONDict) -> None:
        workbook.add_data_validation(sheet, validation)

    def add_named_range(self, workbook: Any, sheet: str, named_range: JSONDict) -> None:
        workbook.add_named_range(sheet, named_range)

    def add_table(self, workbook: Any, sheet: str, table: JSONDict) -> None:
        workbook.add_table(sheet, table.get("table", table))

    def add_hyperlink(self, workbook: Any, sheet: str, link: JSONDict) -> None:
        link_data = link.get("hyperlink", link)
        cell = link_data.get("cell")
        display = link_data.get("display") or link_data.get("target")
        if isinstance(cell, str) and display is not None:
            self.write_cell_value(
                workbook,
                sheet,
                cell,
                CellValue(type=CellType.STRING, value=str(display)),
            )
        workbook.add_hyperlink(sheet, link_data)

    def add_image(self, workbook: Any, sheet: str, image: JSONDict) -> None:
        image_data = image.get("image", image)
        path = image_data.get("path")
        if not isinstance(path, str):
            return
        from wolfxl._images import image_to_writer_payload
        from wolfxl.drawing.image import Image

        img = Image(path)
        anchor = image_data.get("cell") or "A1"
        img.anchor = str(anchor)
        workbook.add_image(sheet, image_to_writer_payload(img))

    def add_pivot_table(self, workbook: Any, sheet: str, pivot: JSONDict) -> None:
        raise NotImplementedError("wolfxl pivot tables not implemented")

    def add_comment(self, workbook: Any, sheet: str, comment: JSONDict) -> None:
        workbook.add_comment(sheet, comment)

    def set_freeze_panes(self, workbook: Any, sheet: str, settings: JSONDict) -> None:
        workbook.set_freeze_panes(sheet, settings)

    def save_workbook(self, workbook: Any, path: Path) -> None:
        workbook.save(str(path))
