# ruff: noqa: N802, N815, N816
"""ExcelBench JSON-model helper executed inside LibreOffice's embedded Python.

``libreoffice_uno_adapter.py`` copies this module into
``<profile>/user/Scripts/python`` and launches::

    soffice --headless ... \
        'vnd.sun.star.script:excelbench_uno.py$main?language=Python&location=user'

The request and response file paths arrive through the
``EXCELBENCH_UNO_REQUEST`` / ``EXCELBENCH_UNO_RESPONSE`` environment
variables. Every value reported here comes from the Calc UNO API after
LibreOffice's own OOXML import; nothing parses workbook XML.

Unit conversions mirror what Calc's own xlsx export does with the same data,
so a value reported here is the value LibreOffice would write back:

* row height: Calc stores twips; UNO reports 1/100 mm. Points =
  round(hmm * 1440 / 2540) / 20.
* column width: UNO 1/100 mm divided by the width of the widest digit of the
  workbook default font (the ``Default`` cell style font measured on the
  document reference device, as ``XclRoot::SetCharWidth`` does), truncated to
  two decimals (``XclExpColinfo::SaveXml``).
* borders: (line style, outer width in twips, double-line distance) mapped to
  OOXML border styles with the thresholds of ``lclGetBorderLine`` in
  ``sc/source/filter/excel/xestyle.cxx``.
* text rotation: ``XclTools::GetXclRotation``.
* formulas, defined names, CF/DV formulas, internal hyperlink targets: Calc
  formula tokens printed through ``com.sun.star.sheet.FormulaParser`` with
  the OOXML op-code map and the XL_OOX address convention.
"""

from __future__ import annotations

import json
import os
import traceback
from datetime import datetime, timedelta

import uno
import unohelper
from com.sun.star.awt import FontDescriptor, Point, Rectangle, Size
from com.sun.star.beans import PropertyValue
from com.sun.star.lang import Locale
from com.sun.star.sheet import XDataBarEntry
from com.sun.star.table import BorderLine2, CellAddress, CellRangeAddress

LIBRARY = "libreoffice"
REQUEST_ENV = "EXCELBENCH_UNO_REQUEST"
RESPONSE_ENV = "EXCELBENCH_UNO_RESPONSE"
XLSX_FILTER = "Calc Office Open XML"
CHART_CLSID = "12DCAE26-281F-416F-A234-C3086127382E"

# Minimum scan window for row heights / column widths. Calc reports a height
# and width for every row and column, so custom dimensions are found by
# scanning rows/columns up to max(used area, these bounds).
MIN_SCAN_ROWS = 100
MIN_SCAN_COLS = 26

HMM_PER_TWIP = 2540.0 / 1440.0
EMU_PER_HMM = 360

# Write ops whose content the Calc UNO API cannot express at all.
UNSUPPORTED_WRITE: dict[str, str] = {}


def _const(name):
    return uno.getConstantByName(name)


def _enum(type_name, value):
    return uno.Enum(type_name, value)


def pv(name, value):
    prop = PropertyValue()
    prop.Name = name
    prop.Value = value
    return prop


# =============================================================================
# Entry point
# =============================================================================


def main(*_args):
    response_path = os.environ.get(RESPONSE_ENV)
    try:
        with open(os.environ[REQUEST_ENV], encoding="utf-8") as handle:
            request = json.load(handle)
        response = Helper(XSCRIPTCONTEXT).handle(request)  # noqa: F821
    except Exception as exc:
        message = getattr(exc, "Message", None) or str(exc)
        response = {
            "error": "libreoffice_failed",
            "message": f"{type(exc).__name__}: {message}",
            "traceback": traceback.format_exc(),
        }
    if response_path:
        with open(response_path, "w", encoding="utf-8") as handle:
            json.dump(response, handle, sort_keys=True)


g_exportedScripts = (main,)


# =============================================================================
# Address helpers
# =============================================================================


def col_letter(index):
    """0-based column index -> letters."""
    index += 1
    letters = ""
    while index > 0:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def a1(col, row):
    return f"{col_letter(col)}{row + 1}"


def a1_range(addr):
    start = a1(addr.StartColumn, addr.StartRow)
    end = a1(addr.EndColumn, addr.EndRow)
    return start if start == end else f"{start}:{end}"


def parse_a1(ref):
    """'B12' or '$B$12' -> (col, row) 0-based."""
    ref = ref.replace("$", "").strip().upper()
    col = 0
    i = 0
    while i < len(ref) and ref[i].isalpha():
        col = col * 26 + (ord(ref[i]) - 64)
        i += 1
    if i == 0 or i == len(ref):
        raise ValueError(f"invalid cell reference {ref!r}")
    return col - 1, int(ref[i:]) - 1


def parse_a1_range(ref):
    ref = ref.split("!")[-1]
    if ":" in ref:
        start, end = ref.split(":", 1)
    else:
        start = end = ref
    c1, r1 = parse_a1(start)
    c2, r2 = parse_a1(end)
    return c1, r1, c2, r2


def color_hex(value):
    if value is None or value == -1:
        return None
    return f"#{value & 0xFFFFFF:06X}"


def color_int(hex_color):
    return int(str(hex_color).lstrip("#")[-6:], 16)


def to_json_number(value):
    if isinstance(value, float) and value.is_integer() and abs(value) < 2**53:
        return int(value)
    return value


# =============================================================================
# Helper
# =============================================================================


class Helper:
    def __init__(self, script_context):
        self.script_context = script_context
        self.ctx = script_context.getComponentContext()
        self.smgr = self.ctx.ServiceManager
        self.desktop = script_context.getDesktop()

    def create(self, service):
        return self.smgr.createInstanceWithContext(service, self.ctx)

    # ------------------------------------------------------------------ dispatch

    def handle(self, request):
        operation = request.get("operation")
        if operation == "describe":
            return self.describe()
        if operation == "read_model":
            return {"model": self.read_model(self._input_path(request))}
        if operation == "write_model":
            output = self._output_path(request)
            ops = (request.get("payload") or {}).get("ops") or []
            self.write_model(ops, output)
            return {"written": output}
        if operation == "mutate":
            output = self._output_path(request)
            mutations = (request.get("payload") or {}).get("mutations") or []
            self.mutate(self._input_path(request), mutations, output)
            return {"written": output}
        return {
            "error": "libreoffice_failed",
            "message": f"unsupported operation {operation!r}",
        }

    @staticmethod
    def _input_path(request):
        path = request.get("input_path")
        if not path or not os.path.isfile(path):
            raise ValueError(f"input_path does not exist: {path!r}")
        return os.path.abspath(path)

    @staticmethod
    def _output_path(request):
        path = request.get("output_path")
        if not path:
            raise ValueError("output_path is required")
        path = os.path.abspath(path)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        return path

    # ------------------------------------------------------------------ describe

    def describe(self):
        return {
            "library": LIBRARY,
            "version": self.product_version(),
            "capabilities": ["read", "write", "modify"],
            "unsupported_write": dict(UNSUPPORTED_WRITE),
        }

    def product_version(self):
        provider = self.create("com.sun.star.configuration.ConfigurationProvider")
        access = provider.createInstanceWithArguments(
            "com.sun.star.configuration.ConfigurationAccess",
            (pv("nodepath", "/org.openoffice.Setup/Product"),),
        )
        version = str(access.getByName("ooSetupVersionAboutBox"))
        if access.hasByName("ooSetupVersionAboutBoxSuffix"):
            suffix = access.getByName("ooSetupVersionAboutBoxSuffix")
            if suffix:
                version = f"{version}{suffix}"
        build_id = self.build_id()
        return f"{version} (build {build_id})" if build_id else version

    def build_id(self):
        try:
            expander = self.ctx.getValueByName("/singletons/com.sun.star.util.theMacroExpander")
        except Exception:
            return None
        for rc in ("Resources/versionrc", "program/versionrc", "program/version.ini"):
            try:
                value = expander.expandMacros(f"${{$BRAND_BASE_DIR/{rc}:buildid}}")
            except Exception:
                continue
            if value and "$" not in value and "<" not in value:
                return value
        return None

    # ------------------------------------------------------------------ documents

    def load(self, path):
        url = uno.systemPathToFileUrl(path)
        doc = self.desktop.loadComponentFromURL(
            url, "_blank", 0, (pv("Hidden", True), pv("ReadOnly", False))
        )
        if doc is None:
            raise RuntimeError(f"LibreOffice could not load {path}")
        return doc

    def new_document(self):
        # Not Hidden: soffice already runs --headless, and Calc applies view
        # operations such as freezeAtPosition only to a document whose view
        # has a (headless) window; with Hidden=True they are silently dropped.
        doc = self.desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, ())
        if doc is None:
            raise RuntimeError("LibreOffice could not create a Calc document")
        return doc

    @staticmethod
    def store_xlsx(doc, path):
        doc.storeToURL(
            uno.systemPathToFileUrl(path),
            (pv("FilterName", XLSX_FILTER), pv("Overwrite", True)),
        )
        if not os.path.isfile(path):
            raise RuntimeError(f"LibreOffice did not write {path}")

    @staticmethod
    def close(doc):
        try:
            doc.close(True)
        except Exception:
            doc.dispose()

    # ------------------------------------------------------------------ read

    def read_model(self, path):
        doc = self.load(path)
        try:
            return ModelReader(self, doc).read()
        finally:
            self.close(doc)

    # ------------------------------------------------------------------ write

    def write_model(self, ops, output):
        doc = self.new_document()
        try:
            OpWriter(self, doc).apply(ops)
            self.store_xlsx(doc, output)
        finally:
            self.close(doc)

    def mutate(self, path, mutations, output):
        doc = self.load(path)
        try:
            for mutation in mutations:
                sheet_name = mutation["sheet"]
                if not doc.Sheets.hasByName(sheet_name):
                    raise KeyError(f"sheet not found: {sheet_name}")
                sheet = doc.Sheets.getByName(sheet_name)
                col, row = parse_a1(mutation["cell"])
                set_plain_value(doc, sheet.getCellByPosition(col, row), mutation["value"])
            self.store_xlsx(doc, output)
        finally:
            self.close(doc)


def set_plain_value(doc, cell, value):
    """Set a str/int/float/bool through the cell API (template mutation lane)."""
    if value is None:
        cell.setString("")
        cell.setFormula("")
    elif isinstance(value, bool):
        cell.setValue(1.0 if value else 0.0)
        cell.NumberFormat = standard_format(doc, "LOGICAL")
    elif isinstance(value, (int, float)):
        cell.setValue(float(value))
    else:
        cell.setString(str(value))


def standard_format(doc, kind):
    return doc.NumberFormats.getStandardFormat(
        _const(f"com.sun.star.util.NumberFormat.{kind}"), Locale()
    )


# =============================================================================
# Formula grammar bridge
# =============================================================================


class Grammar:
    """Converts between Calc token arrays and OOXML / Calc formula strings."""

    def __init__(self, doc):
        mapper = doc.createInstance("com.sun.star.sheet.FormulaOpCodeMapper")
        groups = 0
        for group in (
            "SPECIAL",
            "SEPARATORS",
            "ARRAY_SEPARATORS",
            "UNARY_OPERATORS",
            "BINARY_OPERATORS",
            "FUNCTIONS",
        ):
            groups |= _const(f"com.sun.star.sheet.FormulaMapGroup.{group}")
        language = "com.sun.star.sheet.FormulaLanguage."
        convention = "com.sun.star.sheet.AddressConvention."

        def make(lang, conv):
            parser = doc.createInstance("com.sun.star.sheet.FormulaParser")
            parser.FormulaConvention = _const(convention + conv)
            parser.OpCodeMap = mapper.getAvailableMappings(_const(language + lang), groups)
            return parser

        # OOXML: Excel function names, ',' separators, Sheet!A1 references.
        self.ooxml = make("OOXML", "XL_OOX")
        # Calc API grammar: English names, ';' separators, $Sheet.A1 references.
        self.api = make("ENGLISH", "OOO")
        # UI grammar used by the new conditional-format entry API.
        self.native = make("NATIVE", "OOO")

    def tokens_to_ooxml(self, tokens, base):
        return self.ooxml.printFormula(tokens, base)

    def calc_to_ooxml(self, text, base, native=False):
        parser = self.native if native else self.api
        return self.ooxml.printFormula(parser.parseFormula(text, base), base)

    def ooxml_to_tokens(self, text, base):
        return self.ooxml.parseFormula(text, base)

    def ooxml_to_calc(self, text, base, native=False):
        parser = self.native if native else self.api
        return parser.printFormula(self.ooxml.parseFormula(text, base), base)


def strip_eq(text):
    text = str(text).strip()
    return text[1:] if text.startswith("=") else text


# =============================================================================
# Read side
# =============================================================================

UNDERLINE_NAMES = {
    1: "single",
    2: "double",
    3: "dotted",
    5: "dash",
    6: "longDash",
    7: "dashDot",
    8: "dashDotDot",
    9: "smallWave",
    10: "wave",
    11: "doubleWave",
    12: "bold",
    13: "boldDotted",
    14: "boldDash",
    15: "boldLongDash",
    16: "boldDashDot",
    17: "boldDashDotDot",
    18: "boldWave",
}
H_ALIGN_NAMES = {
    "LEFT": "left",
    "CENTER": "center",
    "RIGHT": "right",
    "BLOCK": "justify",
    "REPEAT": "fill",
}
V_ALIGN_NAMES = {1: "top", 2: "center", 3: "bottom", 4: "justify"}

# Twip thresholds of lclGetBorderLine (EXC_BORDER_THICK/MEDIUM/THIN/HAIR).
BORDER_THICK_TWIPS = 50
BORDER_MEDIUM_TWIPS = 35
BORDER_THIN_TWIPS = 15
BORDER_HAIR_TWIPS = 1

# com.sun.star.table.BorderLineStyle
BLS_NONE = 0x7FFF
BLS_SOLID = 0
BLS_DOTTED = 1
BLS_DASHED = 2
BLS_DOUBLE = 3
BLS_FINE_DASHED = 14
BLS_DOUBLE_THIN = 15
BLS_DASH_DOT = 16
BLS_DASH_DOT_DOT = 17

# com.sun.star.sheet.ConditionFormatOperator -> (rule_type, operator)
CF_OPERATORS = {
    0: ("cellIs", "equal"),
    1: ("cellIs", "lessThan"),
    2: ("cellIs", "greaterThan"),
    3: ("cellIs", "lessThanOrEqual"),
    4: ("cellIs", "greaterThanOrEqual"),
    5: ("cellIs", "notEqual"),
    6: ("cellIs", "between"),
    7: ("cellIs", "notBetween"),
    8: ("duplicateValues", None),
    9: ("uniqueValues", None),
    10: ("containsErrors", None),
    11: ("notContainsErrors", None),
    12: ("beginsWith", None),
    13: ("endsWith", None),
    14: ("containsText", None),
    15: ("notContainsText", None),
    16: ("top10", None),
    17: ("top10", None),
    18: ("top10", None),
    19: ("top10", None),
    20: ("aboveAverage", None),
    21: ("aboveAverage", None),
    22: ("aboveAverage", None),
    23: ("aboveAverage", None),
    24: ("expression", None),
}

DV_TYPES = {
    "WHOLE": "whole",
    "DECIMAL": "decimal",
    "DATE": "date",
    "TIME": "time",
    "TEXT_LEN": "textLength",
    "LIST": "list",
    "CUSTOM": "custom",
}
CONDITION_OPERATORS = {
    "EQUAL": "equal",
    "NOT_EQUAL": "notEqual",
    "GREATER": "greaterThan",
    "GREATER_EQUAL": "greaterThanOrEqual",
    "LESS": "lessThan",
    "LESS_EQUAL": "lessThanOrEqual",
    "BETWEEN": "between",
    "NOT_BETWEEN": "notBetween",
}

CHART_TYPES = {
    "com.sun.star.chart2.ColumnChartType": "bar",
    "com.sun.star.chart2.BarChartType": "bar",
    "com.sun.star.chart2.LineChartType": "line",
    "com.sun.star.chart2.PieChartType": "pie",
    "com.sun.star.chart2.AreaChartType": "area",
    "com.sun.star.chart2.ScatterChartType": "scatter",
    "com.sun.star.chart2.BubbleChartType": "bubble",
    "com.sun.star.chart2.NetChartType": "radar",
    "com.sun.star.chart2.FilledNetChartType": "radar",
    "com.sun.star.chart2.CandleStickChartType": "stock",
}

ACTIVE_PANES = {0: "topLeft", 1: "topRight", 2: "bottomLeft", 3: "bottomRight"}

CELL_PROPS = (
    "CharWeight",
    "CharPosture",
    "CharUnderline",
    "CharStrikeout",
    "CharFontName",
    "CharHeight",
    "CharColor",
    "CellBackColor",
    "IsCellBackgroundTransparent",
    "NumberFormat",
    "HoriJustify",
    "HoriJustifyMethod",
    "VertJustify",
    "IsTextWrapped",
    "RotateAngle",
    "ParaIndent",
    "TopBorder2",
    "BottomBorder2",
    "LeftBorder2",
    "RightBorder2",
    "DiagonalBLTR2",
    "DiagonalTLBR2",
)


def hmm_to_twips(value):
    return int(round(value / HMM_PER_TWIP))


def border_style(line):
    """Map a BorderLine2 to an OOXML style (export thresholds), or None."""
    if line is None or line.LineStyle == BLS_NONE:
        return None
    outer = hmm_to_twips(line.OuterLineWidth or line.LineWidth)
    style = line.LineStyle
    if line.LineDistance > 0 or (
        style in (BLS_DOUBLE, BLS_DOUBLE_THIN) and line.InnerLineWidth > 0
    ):
        return "double"
    if outer >= BORDER_THICK_TWIPS:
        return "thick"
    if outer >= BORDER_MEDIUM_TWIPS:
        return {
            BLS_DASHED: "mediumDashed",
            BLS_DASH_DOT: "mediumDashDot",
            BLS_DASH_DOT_DOT: "mediumDashDotDot",
        }.get(style, "medium")
    if outer >= BORDER_THIN_TWIPS:
        return {
            BLS_DASHED: "dashed",
            BLS_FINE_DASHED: "dashed",
            BLS_DASH_DOT: "dashDot",
            BLS_DASH_DOT_DOT: "dashDotDot",
            BLS_DOTTED: "dotted",
        }.get(style, "thin")
    if outer >= BORDER_HAIR_TWIPS:
        return "hair"
    return None


def xl_rotation(rotate_angle):
    """XclTools::GetXclRotation: Calc 1/100 degree -> OOXML textRotation."""
    degrees = int(rotate_angle) // 100
    if 0 <= degrees <= 90:
        return degrees
    if degrees < 180:
        return 270 - degrees
    if degrees < 270:
        return degrees - 180
    if degrees < 360:
        return 450 - degrees
    return 0


def xl_indent(para_indent_hmm):
    """XclExpCellAlign: indent level = (twips + 100) / 200."""
    return (hmm_to_twips(para_indent_hmm) + 100) // 200


class ModelReader:
    def __init__(self, helper, doc):
        self.helper = helper
        self.doc = doc
        self.grammar = Grammar(doc)
        self.formats = doc.NumberFormats
        self.format_cache = {}
        null = doc.NullDate
        self.null_date = datetime(null.Year, null.Month, null.Day)
        self.char_width_hmm = self._default_digit_width_hmm()
        self.sheet_names = list(doc.Sheets.ElementNames)

    # -------------------------------------------------------------- workbook

    def read(self):
        model = {
            "unsupported": {},
            "named_ranges": self.named_ranges(),
            "sheets": [],
        }
        controller = self.doc.getCurrentController()
        for index, name in enumerate(self.sheet_names):
            sheet = self.doc.Sheets.getByIndex(index)
            model["sheets"].append(self.read_sheet(index, name, sheet, controller))
        return model

    def _default_digit_width_hmm(self):
        """Widest digit of the default font in twips, like XclRoot::SetCharWidth."""
        style = self.doc.StyleFamilies.getByName("CellStyles").getByName("Default")
        descriptor = FontDescriptor()
        descriptor.Name = style.CharFontName
        descriptor.Height = int(round(style.CharHeight * 20))
        descriptor.Weight = style.CharWeight
        font = self.doc.ReferenceDevice.getFont(descriptor)
        width_twips = max(font.getCharWidth(digit) for digit in "0123456789")
        if width_twips <= 0:
            width_twips = 11 * descriptor.Height // 20
        return width_twips * HMM_PER_TWIP

    def named_ranges(self):
        out = []
        sources = [("workbook", None, self.doc.NamedRanges)]
        for name in self.sheet_names:
            sources.append(("sheet", name, self.doc.Sheets.getByName(name).NamedRanges))
        for scope, sheet_name, container in sources:
            for name in container.ElementNames:
                named = container.getByName(name)
                base = named.ReferencePosition
                try:
                    refers_to = self.grammar.tokens_to_ooxml(named.getTokens(), base)
                except Exception:
                    refers_to = named.Content
                out.append(
                    {
                        "name": name,
                        "scope": scope,
                        "sheet": sheet_name,
                        "refers_to": refers_to,
                    }
                )
        return out

    # -------------------------------------------------------------- sheet

    def read_sheet(self, index, name, sheet, controller):
        cursor = sheet.createCursor()
        cursor.gotoEndOfUsedArea(False)
        end = cursor.RangeAddress
        end_col, end_row = end.EndColumn, end.EndRow
        cells = {}
        for row in range(end_row + 1):
            for col in range(end_col + 1):
                entry = self.read_cell(sheet.getCellByPosition(col, row))
                if entry:
                    cells[a1(col, row)] = entry
        return {
            "name": name,
            "cells": cells,
            "row_heights": self.row_heights(sheet, end_row),
            "column_widths": self.column_widths(sheet, end_col),
            "merged_ranges": self.merged_ranges(sheet, end_col, end_row),
            "conditional_formats": self.conditional_formats(sheet),
            "data_validations": self.data_validations(sheet),
            "hyperlinks": self.hyperlinks(sheet, end_col, end_row),
            "images": self.images(sheet),
            "pivot_tables": self.pivot_tables(sheet),
            "comments": self.comments(sheet),
            "freeze_panes": self.freeze_panes(sheet, name, controller),
            "tables": self.tables(sheet, index),
            "page_setup": self.page_setup(sheet),
            "chart_anchors": self.chart_anchors(sheet),
            # Last: probing for a password unprotects the in-memory sheet.
            "sheet_protection": self.sheet_protection(sheet),
        }

    # -------------------------------------------------------------- cells

    def number_format(self, key):
        cached = self.format_cache.get(key)
        if cached is None:
            props = self.formats.getByKey(key)
            cached = (props.FormatString, props.Type)
            self.format_cache[key] = cached
        return cached

    def read_cell(self, cell):
        values = dict(zip(CELL_PROPS, cell.getPropertyValues(CELL_PROPS)))
        format_string, format_type = self.number_format(values["NumberFormat"])
        entry = {
            "format": self.cell_format(values, format_string),
            "border": self.cell_border(values),
        }
        value = self.cell_value(cell, format_type)
        if value is not None:
            entry["value"] = value
        return entry

    def cell_value(self, cell, format_type):
        kind = cell.Type.value
        if kind == "EMPTY":
            return None
        if kind == "TEXT":
            return {"type": "string", "value": cell.String, "formula": None}
        if kind == "FORMULA":
            if cell.getError() != 0:
                return {"type": "error", "value": cell.String, "formula": None}
            formula = "=" + self.grammar.tokens_to_ooxml(cell.getTokens(), cell.CellAddress)
            return {"type": "formula", "value": formula, "formula": formula}
        number = cell.getValue()
        if format_type & _const("com.sun.star.util.NumberFormat.LOGICAL"):
            return {"type": "boolean", "value": number != 0, "formula": None}
        if format_type & _const("com.sun.star.util.NumberFormat.DATE"):
            stamp = self.null_date + timedelta(seconds=round(number * 86400))
            if stamp.hour == 0 and stamp.minute == 0 and stamp.second == 0:
                return {
                    "type": "date",
                    "value": stamp.date().isoformat(),
                    "formula": None,
                }
            return {
                "type": "datetime",
                "value": stamp.isoformat(timespec="seconds"),
                "formula": None,
            }
        return {"type": "number", "value": to_json_number(number), "formula": None}

    @staticmethod
    def cell_format(values, format_string):
        fmt = {
            "bold": values["CharWeight"] >= 150,
            "italic": values["CharPosture"].value in ("ITALIC", "OBLIQUE"),
            "strikethrough": values["CharStrikeout"] not in (0, 3),
            "font_name": values["CharFontName"],
            "font_size": to_json_number(float(values["CharHeight"])),
            "number_format": format_string,
            "wrap": bool(values["IsTextWrapped"]),
        }
        underline = UNDERLINE_NAMES.get(values["CharUnderline"])
        if underline:
            fmt["underline"] = underline
        font_color = color_hex(values["CharColor"])
        if font_color:
            fmt["font_color"] = font_color
        if not values["IsCellBackgroundTransparent"]:
            bg = color_hex(values["CellBackColor"])
            if bg:
                fmt["bg_color"] = bg
        h_align = H_ALIGN_NAMES.get(values["HoriJustify"].value)
        if h_align == "justify" and values["HoriJustifyMethod"] == 1:
            h_align = "distributed"
        if h_align:
            fmt["h_align"] = h_align
        v_align = V_ALIGN_NAMES.get(values["VertJustify"])
        if v_align:
            fmt["v_align"] = v_align
        rotation = xl_rotation(values["RotateAngle"])
        if rotation:
            fmt["rotation"] = rotation
        indent = xl_indent(values["ParaIndent"])
        if indent:
            fmt["indent"] = indent
        return fmt

    @staticmethod
    def cell_border(values):
        border = {}
        for edge, prop in (
            ("top", "TopBorder2"),
            ("bottom", "BottomBorder2"),
            ("left", "LeftBorder2"),
            ("right", "RightBorder2"),
            ("diagonal_up", "DiagonalBLTR2"),
            ("diagonal_down", "DiagonalTLBR2"),
        ):
            line = values[prop]
            style = border_style(line)
            if style:
                border[edge] = {
                    "style": style,
                    "color": color_hex(line.Color) or "#000000",
                }
        return border

    # -------------------------------------------------------------- dimensions

    @staticmethod
    def row_heights(sheet, end_row):
        out = {}
        rows = sheet.Rows
        for row in range(max(end_row + 1, MIN_SCAN_ROWS)):
            props = rows.getByIndex(row)
            if props.OptimalHeight:
                continue
            twips = hmm_to_twips(props.Height)
            out[str(row + 1)] = to_json_number(twips / 20.0)
        return out

    def column_widths(self, sheet, end_col):
        out = {}
        columns = sheet.Columns
        for col in range(max(end_col + 1, MIN_SCAN_COLS)):
            width = columns.getByIndex(col).Width
            chars = int(width / self.char_width_hmm * 100.0 + 0.5) / 100.0
            out[col_letter(col)] = chars
        return out

    # -------------------------------------------------------------- merges

    @staticmethod
    def merged_ranges(sheet, end_col, end_row):
        out = []
        for row in range(end_row + 1):
            for col in range(end_col + 1):
                single = sheet.getCellRangeByPosition(col, row, col, row)
                if not single.getIsMerged():
                    continue
                cursor = sheet.createCursorByRange(single)
                cursor.collapseToMergedArea()
                rng = a1_range(cursor.RangeAddress)
                if ":" in rng and rng not in out:
                    out.append(rng)
        return out

    # -------------------------------------------------------------- CF / DV

    def _ranges_string(self, addresses):
        return " ".join(a1_range(addr) for addr in addresses)

    def conditional_formats(self, sheet):
        out = []
        styles = self.doc.StyleFamilies.getByName("CellStyles")
        for cond_format in sheet.ConditionalFormats.ConditionalFormats:
            addresses = cond_format.Range.RangeAddresses
            if not addresses:
                continue
            range_text = self._ranges_string(addresses)
            first = addresses[0]
            base = CellAddress(first.Sheet, first.StartColumn, first.StartRow)
            for position in range(cond_format.Count):
                entry = cond_format.getByIndex(position)
                info = entry.getPropertySetInfo()
                rule = {
                    "range": range_text,
                    "rule_type": None,
                    "operator": None,
                    "formula": None,
                    "priority": None,
                    "stop_if_true": None,
                    "format": {},
                }
                # The numeric getType() codes differ between entry
                # implementations, so classify by the properties each kind
                # of entry object exposes.
                if info.hasPropertyByName("ColorScaleEntries"):
                    rule["rule_type"] = "colorScale"
                elif info.hasPropertyByName("DataBarEntries"):
                    rule["rule_type"] = "dataBar"
                elif info.hasPropertyByName("IconSetEntries"):
                    rule["rule_type"] = "iconSet"
                elif info.hasPropertyByName("DateType"):
                    rule["rule_type"] = "timePeriod"
                    rule["format"] = self._style_colors(styles, entry.StyleName)
                elif info.hasPropertyByName("Operator"):
                    rule_type, operator = CF_OPERATORS.get(entry.Operator, (None, None))
                    rule["rule_type"] = rule_type
                    rule["operator"] = operator
                    formula = entry.Formula1
                    if formula:
                        try:
                            rule["formula"] = self.grammar.calc_to_ooxml(formula, base, native=True)
                        except Exception:
                            rule["formula"] = formula
                    rule["format"] = self._style_colors(styles, entry.StyleName)
                out.append(rule)
        return out

    @staticmethod
    def _style_colors(styles, style_name):
        if not style_name or not styles.hasByName(style_name):
            return {}
        style = styles.getByName(style_name)
        out = {}
        if not style.IsCellBackgroundTransparent:
            bg = color_hex(style.CellBackColor)
            if bg:
                out["bg_color"] = bg
        direct = uno.Enum("com.sun.star.beans.PropertyState", "DIRECT_VALUE")
        if style.getPropertyState("CharColor") == direct:
            font_color = color_hex(style.CharColor)
            if font_color:
                out["font_color"] = font_color
        return out

    def data_validations(self, sheet):
        groups = {}
        order = []
        unique = sheet.UniqueCellFormatRanges
        for index in range(unique.Count):
            ranges = unique.getByIndex(index)
            validation = ranges.Validation
            if validation.Type.value == "ANY":
                continue
            addresses = ranges.RangeAddresses
            first = addresses[0]
            base = CellAddress(first.Sheet, first.StartColumn, first.StartRow)
            record = self._validation_record(validation, base)
            key = json.dumps(record, sort_keys=True)
            if key not in groups:
                groups[key] = (record, [])
                order.append(key)
            groups[key][1].extend(addresses)
        out = []
        for key in order:
            record, addresses = groups[key]
            container = self.doc.createInstance("com.sun.star.sheet.SheetCellRanges")
            container.addRangeAddresses(tuple(addresses), True)
            merged = sorted(container.RangeAddresses, key=lambda a: (a.StartRow, a.StartColumn))
            out.append({"range": self._ranges_string(merged), **record})
        return out

    def _validation_record(self, validation, base):
        dv_type = DV_TYPES.get(validation.Type.value, validation.Type.value.lower())
        formula1 = self._validation_formula(validation.getFormula1(), base, dv_type == "list")
        formula2 = None
        operator = CONDITION_OPERATORS.get(validation.Operator.value)
        if operator in ("between", "notBetween"):
            formula2 = self._validation_formula(validation.getFormula2(), base, False)
        return {
            "validation_type": dv_type,
            "operator": operator,
            "formula1": formula1,
            "formula2": formula2,
            "allow_blank": bool(validation.IgnoreBlankCells),
            "show_input": bool(validation.ShowInputMessage),
            "show_error": bool(validation.ShowErrorMessage),
            "prompt_title": validation.InputTitle or None,
            "prompt": validation.InputMessage or None,
            "error_title": validation.ErrorTitle or None,
            "error": validation.ErrorMessage or None,
        }

    def _validation_formula(self, text, base, is_list):
        """Print a validation formula (Calc API grammar) in OOXML grammar."""
        if not text:
            return None
        tokens = self.grammar.api.parseFormula(text, base)
        if is_list:
            # A list of string constants is written by Calc's xlsx export as one
            # quoted comma-joined literal (XclExpDV); report it the same way.
            strings = []
            only_strings = True
            for index, token in enumerate(tokens):
                if index % 2 == 0:
                    if token.OpCode == 0 and isinstance(token.Data, str):
                        strings.append(token.Data)
                    else:
                        only_strings = False
                        break
            if only_strings and strings and len(tokens) == 2 * len(strings) - 1:
                return '"' + ",".join(strings) + '"'
        return self.grammar.tokens_to_ooxml(tokens, base)

    # -------------------------------------------------------------- links

    def hyperlinks(self, sheet, end_col, end_row):
        out = []
        for row in range(end_row + 1):
            for col in range(end_col + 1):
                cell = sheet.getCellByPosition(col, row)
                if cell.Type.value == "EMPTY":
                    continue
                found = False
                fields = cell.Text.TextFields.createEnumeration()
                while fields.hasMoreElements():
                    field = fields.nextElement()
                    if not field.getPropertySetInfo().hasPropertyByName("URL"):
                        continue
                    found = True
                    out.append(self._link_record(sheet, a1(col, row), cell, field.URL))
                cell_link = cell.Hyperlink
                if cell_link and not found:
                    out.append(self._link_record(sheet, a1(col, row), cell, cell_link))
        return out

    def _link_record(self, sheet, ref, cell, url):
        internal = url.startswith("#")
        target = url
        if internal:
            # Calc stores internal targets in its own grammar ("Sheet.A1");
            # print them in OOXML grammar. A target Calc cannot parse as a
            # reference (e.g. a defined name) is reported verbatim.
            location = url[1:]
            base = CellAddress(sheet.RangeAddress.Sheet, 0, 0)
            try:
                target = self.grammar.calc_to_ooxml(location, base)
            except Exception:
                target = location
        return {
            "cell": ref,
            "target": target,
            "display": cell.String,
            "tooltip": None,
            "internal": internal,
        }

    # -------------------------------------------------------------- drawings

    @staticmethod
    def _anchor(shape):
        anchor = shape.Anchor
        address = getattr(anchor, "CellAddress", None) if anchor is not None else None
        if address is None:
            return None, None
        kind = "twoCell" if shape.ResizeWithCell else "oneCell"
        return address, kind

    @staticmethod
    def _cell_at(sheet, x, y):
        """Cell holding the shape end point (x, y) in 1/100 mm.

        Uses Calc's own cell positions (accumulated in twips, so they differ
        from summing the rounded 1/100 mm sizes). A point exactly on a cell
        boundary belongs to the preceding cell with a full-size offset, the
        convention Calc's xlsx export uses for the ``xdr:to`` anchor.
        """
        last_col = sheet.Columns.Count - 1
        last_row = sheet.Rows.Count - 1
        col = 0
        while col < last_col and sheet.getCellByPosition(col + 1, 0).Position.X < x:
            col += 1
        row = 0
        while row < last_row and sheet.getCellByPosition(0, row + 1).Position.Y < y:
            row += 1
        return col, row

    def images(self, sheet):
        out = []
        page = sheet.DrawPage
        for index in range(page.Count):
            shape = page.getByIndex(index)
            if shape.ShapeType != "com.sun.star.drawing.GraphicObjectShape":
                continue
            address, kind = self._anchor(shape)
            offset = None
            if address is not None:
                cell = sheet.getCellByPosition(address.Column, address.Row)
                origin = cell.Position
                offset = [
                    (shape.Position.X - origin.X) * EMU_PER_HMM,
                    (shape.Position.Y - origin.Y) * EMU_PER_HMM,
                ]
            out.append(
                {
                    "cell": a1(address.Column, address.Row) if address is not None else None,
                    "path": None,
                    "anchor": kind,
                    "offset": offset,
                    "alt_text": shape.Description or None,
                }
            )
        return out

    def chart_anchors(self, sheet):
        out = []
        shapes = {}
        page = sheet.DrawPage
        for index in range(page.Count):
            shape = page.getByIndex(index)
            if shape.ShapeType == "com.sun.star.drawing.OLE2Shape":
                shapes[shape.PersistName] = shape
        charts = sheet.Charts
        for name in charts.ElementNames:
            chart = charts.getByName(name)
            shape = shapes.get(name)
            chart_type = None
            embedded = chart.EmbeddedObject
            if embedded is not None:
                diagram = embedded.getFirstDiagram()
                if diagram is not None:
                    for system in diagram.getCoordinateSystems():
                        for kind in system.getChartTypes():
                            chart_type = CHART_TYPES.get(kind.ChartType, kind.ChartType)
                            break
                        if chart_type:
                            break
            record = {"type": chart_type, "anchor_type": None, "from": None, "to": None}
            if shape is not None:
                address, kind = self._anchor(shape)
                record["anchor_type"] = kind
                if address is not None:
                    record["from"] = a1(address.Column, address.Row)
                if kind == "twoCell":
                    col, row = self._cell_at(
                        sheet,
                        shape.Position.X + shape.Size.Width,
                        shape.Position.Y + shape.Size.Height,
                    )
                    record["to"] = a1(col, row)
            out.append(record)
        return out

    # -------------------------------------------------------------- objects

    def pivot_tables(self, sheet):
        out = []
        pivots = sheet.DataPilotTables
        for name in pivots.ElementNames:
            pivot = pivots.getByName(name)
            source = pivot.SourceRange
            output = pivot.OutputRange
            out.append(
                {
                    "name": name,
                    "source_range": f"{self.sheet_names[source.Sheet]}!{a1_range(source)}",
                    "target_cell": f"{self.sheet_names[output.Sheet]}!"
                    f"{a1(output.StartColumn, output.StartRow)}",
                }
            )
        return out

    @staticmethod
    def comments(sheet):
        out = []
        for annotation in sheet.Annotations:
            position = annotation.Position
            shape = annotation.AnnotationShape
            text = shape.String if shape is not None else annotation.String
            out.append(
                {
                    "cell": a1(position.Column, position.Row),
                    "text": text,
                    "author": annotation.Author or None,
                    "threaded": False,
                }
            )
        return out

    def freeze_panes(self, sheet, name, controller):
        if controller is not None:
            controller.setActiveSheet(sheet)
            if controller.hasFrozenPanes():
                return {
                    "mode": "freeze",
                    "top_left_cell": a1(controller.getSplitColumn(), controller.getSplitRow()),
                }
        settings = self._view_settings(name)
        if not settings:
            return {}
        h_mode = settings.get("HorizontalSplitMode", 0)
        v_mode = settings.get("VerticalSplitMode", 0)
        if h_mode != 1 and v_mode != 1:
            return {}
        out = {
            "mode": "split",
            "x_split": int(settings.get("HorizontalSplitPositionTwips", 0)) if h_mode == 1 else 0,
            "y_split": int(settings.get("VerticalSplitPositionTwips", 0)) if v_mode == 1 else 0,
            "top_left_cell": a1(
                int(settings.get("PositionRight", 0)),
                int(settings.get("PositionBottom", 0)),
            ),
        }
        pane = ACTIVE_PANES.get(settings.get("ActiveSplitRange"))
        if pane:
            out["active_pane"] = pane
        return out

    def _view_settings(self, sheet_name):
        view_data = self.doc.ViewData
        if view_data is None or view_data.Count == 0:
            return {}
        view = {prop.Name: prop.Value for prop in view_data.getByIndex(0)}
        tables = view.get("Tables")
        if tables is None or not tables.hasByName(sheet_name):
            return {}
        return {prop.Name: prop.Value for prop in tables.getByName(sheet_name)}

    def tables(self, sheet, index):
        out = []
        ranges = self.doc.DatabaseRanges
        for name in ranges.ElementNames:
            db_range = ranges.getByName(name)
            area = db_range.DataArea
            if area.Sheet != index:
                continue
            header = bool(db_range.ContainsHeader)
            columns = []
            if header:
                for col in range(area.StartColumn, area.EndColumn + 1):
                    columns.append(sheet.getCellByPosition(col, area.StartRow).String)
            style = getattr(db_range, "TableStyleName", None)
            out.append(
                {
                    "name": name,
                    "ref": a1_range(area),
                    "header_row": header,
                    "totals_row": bool(getattr(db_range, "TotalsRow", False)),
                    "style": style if style and style != "none" else None,
                    "columns": columns,
                    "autofilter": bool(db_range.AutoFilter),
                }
            )
        return out

    @staticmethod
    def sheet_protection(sheet):
        protected = bool(sheet.isProtected())
        password = False
        if protected:
            # XProtectable has no "has password" query: an empty password
            # unprotects only a sheet that has none.
            try:
                sheet.unprotect("")
            except Exception:
                password = True
        return {
            "protected": protected,
            "password_hash_present": password,
            "format_cells": None,
            "insert_rows": None,
            "select_locked_cells": None,
            "select_unlocked_cells": None,
            "sort": None,
            "auto_filter": None,
        }

    def page_setup(self, sheet):
        style = self.doc.StyleFamilies.getByName("PageStyles").getByName(sheet.PageStyle)
        out = {
            "orientation": "landscape" if style.IsLandscape else "portrait",
            "fit_to_width": None,
            "fit_to_height": None,
            "scale": None,
            "print_title_rows": None,
            "header_center": None,
            "footer_center": None,
        }
        if style.ScaleToPagesX or style.ScaleToPagesY:
            out["fit_to_width"] = int(style.ScaleToPagesX) or None
            out["fit_to_height"] = int(style.ScaleToPagesY) or None
        elif not style.ScaleToPages:
            out["scale"] = int(style.PageScale)
        if sheet.getPrintTitleRows():
            rows = sheet.getTitleRows()
            out["print_title_rows"] = f"${rows.StartRow + 1}:${rows.EndRow + 1}"
        if style.HeaderIsOn:
            out["header_center"] = header_footer_text(style.RightPageHeaderContent.CenterText)
        if style.FooterIsOn:
            out["footer_center"] = header_footer_text(style.RightPageFooterContent.CenterText)
        return out


# Text field service -> OOXML header/footer code (XclExpHFConverter).
HF_FIELD_CODES = (
    ("PageNumber", "&P"),
    ("PageCount", "&N"),
    ("PageCountRange", "&N"),
    ("DateTime", "&D"),
    ("Date", "&D"),
    ("Time", "&T"),
    ("SheetName", "&A"),
    ("FileName", "&F"),
)


def header_footer_text(text):
    parts = []
    paragraphs = text.createEnumeration()
    first = True
    while paragraphs.hasMoreElements():
        paragraph = paragraphs.nextElement()
        if not first:
            parts.append("\n")
        first = False
        portions = paragraph.createEnumeration()
        while portions.hasMoreElements():
            portion = portions.nextElement()
            if portion.TextPortionType == "TextField":
                field = portion.TextField
                code = None
                for service, mapped in HF_FIELD_CODES:
                    if field.supportsService(
                        f"com.sun.star.text.TextField.{service}"
                    ) or field.supportsService(f"com.sun.star.text.textfield.{service}"):
                        code = mapped
                        break
                parts.append(code if code is not None else portion.String)
            else:
                parts.append(portion.String.replace("&", "&&"))
    value = "".join(parts)
    return value or None


# =============================================================================
# Write side
# =============================================================================

H_ALIGN_VALUES = {
    "general": "STANDARD",
    "left": "LEFT",
    "center": "CENTER",
    "right": "RIGHT",
    "justify": "BLOCK",
    "distributed": "BLOCK",
    "fill": "REPEAT",
}
V_ALIGN_VALUES = {"top": 1, "center": 2, "bottom": 3, "justify": 4, "distributed": 4}

# OOXML border style -> (BorderLineStyle, width in 1/100 mm). Widths are the
# ones LibreOffice's own OOXML import assigns (thin 15, medium 35, thick 50
# twips); the export maps them back through lclGetBorderLine.
BORDER_WRITE = {
    "thin": (BLS_SOLID, 26),
    "medium": (BLS_SOLID, 62),
    "thick": (BLS_SOLID, 88),
    "double": (BLS_DOUBLE_THIN, 62),
    "hair": (BLS_SOLID, 2),
    "dotted": (BLS_DOTTED, 26),
    "dashed": (BLS_FINE_DASHED, 26),
    "dashDot": (BLS_DASH_DOT, 26),
    "dashDotDot": (BLS_DASH_DOT_DOT, 26),
    "mediumDashed": (BLS_DASHED, 62),
    "mediumDashDot": (BLS_DASH_DOT, 62),
    "mediumDashDotDot": (BLS_DASH_DOT_DOT, 62),
    "slantDashDot": (BLS_DASH_DOT, 62),
}
BORDER_PROPS = {
    "top": "TopBorder2",
    "bottom": "BottomBorder2",
    "left": "LeftBorder2",
    "right": "RightBorder2",
    "diagonal_up": "DiagonalBLTR2",
    "diagonal_down": "DiagonalTLBR2",
}
UNDERLINE_VALUES = {name: value for value, name in UNDERLINE_NAMES.items()}
UNDERLINE_VALUES.update({"singleAccounting": 1, "doubleAccounting": 2, "none": 0})

ERROR_FORMULAS = {
    "#DIV/0!": "1/0",
    "#N/A": "NA()",
    "#VALUE!": '"text"+1',
    "#REF!": "#REF!",
    "#NAME?": "_undefined_name_",
    "#NUM!": "SQRT(-1)",
    "#NULL!": "A1:A2 B1:B2",
}

CF_OPERATOR_VALUES = {
    "equal": 0,
    "lessThan": 1,
    "greaterThan": 2,
    "lessThanOrEqual": 3,
    "greaterThanOrEqual": 4,
    "notEqual": 5,
    "between": 6,
    "notBetween": 7,
}
DV_TYPE_VALUES = {
    "whole": "WHOLE",
    "decimal": "DECIMAL",
    "date": "DATE",
    "time": "TIME",
    "textLength": "TEXT_LEN",
    "list": "LIST",
    "custom": "CUSTOM",
    "any": "ANY",
    "none": "ANY",
}
CONDITION_OPERATOR_VALUES = {v: k for k, v in CONDITION_OPERATORS.items()}

HF_CODE_SERVICES = {
    "P": "com.sun.star.text.TextField.PageNumber",
    "N": "com.sun.star.text.TextField.PageCount",
    "D": "com.sun.star.text.TextField.Date",
    "T": "com.sun.star.text.TextField.Time",
    "A": "com.sun.star.text.TextField.SheetName",
    "F": "com.sun.star.text.TextField.FileName",
}


class BarEntry(unohelper.Base, XDataBarEntry):
    def __init__(self, entry_type, formula):
        self._type = entry_type
        self._formula = formula

    def getType(self):
        return self._type

    def setType(self, value):
        self._type = value

    def getFormula(self):
        return self._formula

    def setFormula(self, value):
        self._formula = value


class OpWriter:
    def __init__(self, helper, doc):
        self.helper = helper
        self.doc = doc
        self.grammar = Grammar(doc)
        self.default_sheets = list(doc.Sheets.ElementNames)
        self.requested = []
        self.sheets_final = False
        # Empty locale = the document's default language, which is what Calc's
        # UI uses; an explicit locale makes the xlsx export prefix "[$-409]".
        self.locale = Locale()
        self.char_width_hmm = None
        self.style_counter = 0

    def apply(self, ops):
        for op in ops:
            kind = op.get("op")
            if kind == "add_sheet":
                self.add_sheet(str(op["name"]))
                continue
            self.finalize_sheets()
            handler = getattr(self, f"op_{kind}", None)
            if handler is None:
                raise ValueError(f"unknown write op {kind!r}")
            handler(op)
        self.finalize_sheets()

    # -------------------------------------------------------------- sheets

    def add_sheet(self, name):
        if self.sheets_final:
            raise ValueError("add_sheet after other ops is not part of the contract")
        sheets = self.doc.Sheets
        position = len(self.requested)
        if sheets.hasByName(name):
            sheets.moveByName(name, position)
        else:
            sheets.insertNewByName(name, position)
        self.requested.append(name)

    def finalize_sheets(self):
        if self.sheets_final:
            return
        self.sheets_final = True
        if not self.requested:
            return
        for name in self.default_sheets:
            if name not in self.requested and self.doc.Sheets.hasByName(name):
                self.doc.Sheets.removeByName(name)

    def sheet(self, name):
        if not self.doc.Sheets.hasByName(name):
            raise KeyError(f"sheet not found: {name}")
        return self.doc.Sheets.getByName(name)

    def cell(self, op, ref=None):
        sheet = self.sheet(op["sheet"])
        col, row = parse_a1(ref or op["cell"])
        return sheet.getCellByPosition(col, row)

    def sheet_index(self, name):
        return list(self.doc.Sheets.ElementNames).index(name)

    # -------------------------------------------------------------- cells

    def op_cell_value(self, op):
        cell = self.cell(op)
        kind = op.get("type")
        value = op.get("value")
        if kind == "blank" or (kind == "string" and value is None):
            cell.setString("")
        elif kind == "string":
            cell.setString(str(value))
        elif kind == "number":
            cell.setValue(float(value))
        elif kind == "boolean":
            cell.setValue(1.0 if value else 0.0)
            cell.NumberFormat = standard_format(self.doc, "LOGICAL")
        elif kind in ("date", "datetime"):
            stamp = datetime.fromisoformat(str(value))
            null = self.doc.NullDate
            delta = stamp - datetime(null.Year, null.Month, null.Day)
            cell.setValue(delta.days + delta.seconds / 86400.0)
            cell.NumberFormat = standard_format(self.doc, "DATE" if kind == "date" else "DATETIME")
        elif kind == "formula":
            self.set_formula(cell, op.get("formula") or value)
        elif kind == "error":
            formula = ERROR_FORMULAS.get(str(value))
            if formula is None:
                raise ValueError(f"no formula produces error {value!r}")
            self.set_formula(cell, formula)
        else:
            raise ValueError(f"unknown cell value type {kind!r}")

    def set_formula(self, cell, text):
        tokens = self.grammar.ooxml_to_tokens(strip_eq(text), cell.CellAddress)
        cell.setTokens(tokens)

    def number_format_key(self, code):
        formats = self.doc.NumberFormats
        key = formats.queryKey(code, self.locale, False)
        if key == -1:
            key = formats.addNew(code, self.locale)
        return key

    def op_cell_format(self, op):
        cell = self.cell(op)
        fmt = op.get("format") or {}
        if fmt.get("bold") is not None:
            cell.CharWeight = 150.0 if fmt["bold"] else 100.0
        if fmt.get("italic") is not None:
            cell.CharPosture = _enum(
                "com.sun.star.awt.FontSlant", "ITALIC" if fmt["italic"] else "NONE"
            )
        if fmt.get("underline") is not None:
            name = str(fmt["underline"])
            if name not in UNDERLINE_VALUES:
                raise ValueError(f"LibreOffice has no underline kind {name!r}")
            cell.CharUnderline = UNDERLINE_VALUES[name]
        if fmt.get("strikethrough") is not None:
            cell.CharStrikeout = 1 if fmt["strikethrough"] else 0
        if fmt.get("font_name") is not None:
            cell.CharFontName = str(fmt["font_name"])
        if fmt.get("font_size") is not None:
            cell.CharHeight = float(fmt["font_size"])
        if fmt.get("font_color") is not None:
            cell.CharColor = color_int(fmt["font_color"])
        if fmt.get("bg_color") is not None:
            cell.CellBackColor = color_int(fmt["bg_color"])
            cell.IsCellBackgroundTransparent = False
        if fmt.get("number_format") is not None:
            cell.NumberFormat = self.number_format_key(str(fmt["number_format"]))
        if fmt.get("h_align") is not None:
            value = H_ALIGN_VALUES.get(str(fmt["h_align"]))
            if value is None:
                raise ValueError(f"LibreOffice has no horizontal alignment {fmt['h_align']!r}")
            cell.HoriJustify = _enum("com.sun.star.table.CellHoriJustify", value)
            if fmt["h_align"] == "distributed":
                cell.HoriJustifyMethod = 1
        if fmt.get("v_align") is not None:
            value = V_ALIGN_VALUES.get(str(fmt["v_align"]))
            if value is None:
                raise ValueError(f"LibreOffice has no vertical alignment {fmt['v_align']!r}")
            cell.VertJustify = value
            if fmt["v_align"] == "distributed":
                cell.VertJustifyMethod = 1
        if fmt.get("wrap") is not None:
            cell.IsTextWrapped = bool(fmt["wrap"])
        if fmt.get("rotation") is not None:
            rotation = int(fmt["rotation"])
            if rotation == 255:
                cell.Orientation = _enum("com.sun.star.table.CellOrientation", "STACKED")
            elif rotation <= 90:
                cell.RotateAngle = rotation * 100
            else:
                cell.RotateAngle = (360 - (rotation - 90)) * 100
        if fmt.get("indent") is not None:
            # Inverse of the export's (twips + 100) / 200 indent level.
            cell.ParaIndent = int(round(int(fmt["indent"]) * 200 * HMM_PER_TWIP))

    def op_cell_border(self, op):
        cell = self.cell(op)
        for edge, spec in (op.get("border") or {}).items():
            prop = BORDER_PROPS.get(edge)
            if prop is None:
                raise ValueError(f"unknown border edge {edge!r}")
            style = str(spec.get("style") or "thin")
            if style not in BORDER_WRITE:
                raise ValueError(f"LibreOffice has no border style {style!r}")
            line_style, width = BORDER_WRITE[style]
            line = BorderLine2()
            line.LineStyle = line_style
            line.LineWidth = width
            line.Color = color_int(spec.get("color") or "#000000")
            cell.setPropertyValue(prop, line)

    # -------------------------------------------------------------- dimensions

    def op_row_height(self, op):
        sheet = self.sheet(op["sheet"])
        row = sheet.Rows.getByIndex(int(op["row"]) - 1)
        row.Height = int(round(float(op["height"]) * 20 * HMM_PER_TWIP))

    def op_column_width(self, op):
        sheet = self.sheet(op["sheet"])
        col, _ = parse_a1(f"{op['column']}1")
        if self.char_width_hmm is None:
            self.char_width_hmm = ModelReader._default_digit_width_hmm(self)
        column = sheet.Columns.getByIndex(col)
        column.Width = int(round(float(op["width"]) * self.char_width_hmm))

    def op_merge(self, op):
        self.sheet(op["sheet"]).getCellRangeByName(op["range"]).merge(True)

    # -------------------------------------------------------------- CF / DV

    def cell_ranges(self, sheet, text):
        container = self.doc.createInstance("com.sun.star.sheet.SheetCellRanges")
        for part in str(text).replace(",", " ").split():
            container.addRangeAddress(
                sheet.getCellRangeByName(part.replace("$", "")).RangeAddress, False
            )
        return container

    def style_for(self, fmt):
        self.style_counter += 1
        name = f"ExcelBench_CF_{self.style_counter}"
        style = self.doc.createInstance("com.sun.star.style.CellStyle")
        self.doc.StyleFamilies.getByName("CellStyles").insertByName(name, style)
        if fmt.get("bg_color"):
            style.CellBackColor = color_int(fmt["bg_color"])
            style.IsCellBackgroundTransparent = False
        if fmt.get("font_color"):
            style.CharColor = color_int(fmt["font_color"])
        return name

    def op_conditional_format(self, op):
        sheet = self.sheet(op["sheet"])
        rule = op.get("rule") or {}
        rule_type = rule.get("rule_type")
        if rule_type == "colorScale":
            # Not written. A color scale created through createEntry(COLORSCALE)
            # has zero scale entries, and setting ColorScaleEntries indexes the
            # existing entries without growing them (ScColorScaleFormatObj), which
            # aborts soffice (SIGABRT). No other UNO call adds scale entries, so
            # the rule is left out and its test case fails.
            return
        ranges = self.cell_ranges(sheet, rule["range"])
        first = ranges.RangeAddresses[0]
        base = CellAddress(first.Sheet, first.StartColumn, first.StartRow)
        formats = sheet.ConditionalFormats
        format_id = formats.createByRange(ranges)
        target = next(f for f in formats.ConditionalFormats if f.ID == format_id)
        entry_types = "com.sun.star.sheet.ConditionEntryType."
        if rule_type in ("cellIs", "cellIsRule", "expression", "formula"):
            target.createEntry(_const(entry_types + "CONDITION"), target.Count)
            entry = target.getByIndex(target.Count - 1)
            if rule_type in ("cellIs", "cellIsRule"):
                operator = CF_OPERATOR_VALUES.get(rule.get("operator"))
                if operator is None:
                    raise ValueError(f"unknown cellIs operator {rule.get('operator')!r}")
            else:
                operator = 24
            entry.Operator = operator
            formula = rule.get("formula")
            if formula:
                entry.Formula1 = self.grammar.ooxml_to_calc(strip_eq(formula), base, native=True)
            entry.StyleName = self.style_for(rule.get("format") or {})
        elif rule_type == "dataBar":
            target.createEntry(_const(entry_types + "DATABAR"), target.Count)
            entry = target.getByIndex(target.Count - 1)
            kind = "com.sun.star.sheet.DataBarEntryType."
            entry.DataBarEntries = (
                BarEntry(_const(kind + "DATABAR_MIN"), ""),
                BarEntry(_const(kind + "DATABAR_MAX"), ""),
            )
            entry.Color = 0x638EC6
            entry.ShowValue = True
        else:
            raise ValueError(f"unsupported conditional format rule_type {rule_type!r}")
        # stop_if_true and priority: Calc's conditional-format API has no such
        # properties; entries are evaluated in insertion order.

    def op_data_validation(self, op):
        sheet = self.sheet(op["sheet"])
        spec = op.get("validation") or {}
        ranges = self.cell_ranges(sheet, spec["range"])
        first = ranges.RangeAddresses[0]
        base = CellAddress(first.Sheet, first.StartColumn, first.StartRow)
        dv_type = DV_TYPE_VALUES.get(str(spec.get("validation_type") or "any"))
        if dv_type is None:
            raise ValueError(f"unknown validation type {spec.get('validation_type')!r}")
        for address in ranges.RangeAddresses:
            target = sheet.getCellRangeByPosition(
                address.StartColumn, address.StartRow, address.EndColumn, address.EndRow
            )
            validation = target.Validation
            validation.Type = _enum("com.sun.star.sheet.ValidationType", dv_type)
            operator = spec.get("operator")
            if operator is None and spec.get("formula2") is not None:
                operator = "between"
            if dv_type == "CUSTOM":
                validation.Operator = _enum("com.sun.star.sheet.ConditionOperator", "FORMULA")
            elif operator is not None:
                value = CONDITION_OPERATOR_VALUES.get(str(operator))
                if value is None:
                    raise ValueError(f"unknown validation operator {operator!r}")
                validation.Operator = _enum("com.sun.star.sheet.ConditionOperator", value)
            elif dv_type == "LIST":
                validation.Operator = _enum("com.sun.star.sheet.ConditionOperator", "EQUAL")
            # Relative references are anchored at the top-left cell of the first
            # range (OOXML semantics); Calc anchors them at SourcePosition.
            validation.setSourcePosition(base)
            for position, key in ((0, "formula1"), (1, "formula2")):
                text = spec.get(key)
                if text is None:
                    continue
                validation.setTokens(position, self.validation_tokens(str(text), base, dv_type))
            if spec.get("allow_blank") is not None:
                validation.IgnoreBlankCells = bool(spec["allow_blank"])
            if spec.get("show_input") is not None:
                validation.ShowInputMessage = bool(spec["show_input"])
            if spec.get("show_error") is not None:
                validation.ShowErrorMessage = bool(spec["show_error"])
            for key, prop in (
                ("prompt_title", "InputTitle"),
                ("prompt", "InputMessage"),
                ("error_title", "ErrorTitle"),
                ("error", "ErrorMessage"),
            ):
                if spec.get(key) is not None:
                    validation.setPropertyValue(prop, str(spec[key]))
            target.Validation = validation

    def validation_tokens(self, text, base, dv_type):
        text = strip_eq(text)
        if (
            dv_type == "LIST"
            and len(text) >= 2
            and text[0] == '"'
            and text[-1] == '"'
            and '"' not in text[1:-1]
        ):
            # OOXML stores an explicit list as one quoted comma-joined literal;
            # Calc models it as separate string constants (what its import builds).
            items = text[1:-1].split(",")
            calc = ";".join('"' + item + '"' for item in items)
            return self.grammar.api.parseFormula(calc, base)
        return self.grammar.ooxml_to_tokens(text, base)

    # -------------------------------------------------------------- links etc.

    def op_hyperlink(self, op):
        link = op.get("link") or {}
        cell = self.cell(op, link["cell"])
        target = str(link.get("target") or "")
        if link.get("internal"):
            # Calc keeps internal targets as "#Sheet.A1"; translate an OOXML
            # "Sheet!A1" location when it names an existing sheet. Anything else
            # (unknown sheet, defined name) is stored verbatim as the fragment.
            location = target.lstrip("#")
            sheet_part = location.rsplit("!", 1)[0].strip("'") if "!" in location else None
            if sheet_part is not None and self.doc.Sheets.hasByName(sheet_part):
                base = CellAddress(self.sheet_index(op["sheet"]), 0, 0)
                url = "#" + self.grammar.ooxml_to_calc(location, base)
            else:
                url = "#" + location
        else:
            url = target
        display = link.get("display")
        if display is None:
            display = cell.String or target
        cell.setString("")
        text = cell.Text
        field = self.doc.createInstance("com.sun.star.text.TextField.URL")
        field.URL = url
        field.Representation = str(display)
        text.insertTextContent(text.createTextCursor(), field, False)
        # tooltip: the Calc URL text field has no ScreenTip property.

    def op_image(self, op):
        spec = op.get("image") or {}
        sheet = self.sheet(op["sheet"])
        path = spec.get("path")
        if not path or not os.path.isfile(path):
            raise ValueError(f"image path does not exist: {path!r}")
        provider = self.helper.create("com.sun.star.graphic.GraphicProvider")
        graphic = provider.queryGraphic((pv("URL", uno.systemPathToFileUrl(path)),))
        shape = self.doc.createInstance("com.sun.star.drawing.GraphicObjectShape")
        sheet.DrawPage.add(shape)
        shape.Graphic = graphic
        size = graphic.Size100thMM
        if size.Width <= 0 or size.Height <= 0:
            pixels = graphic.SizePixel
            size = Size(int(pixels.Width * 2540 / 96), int(pixels.Height * 2540 / 96))
        shape.Size = size
        col, row = parse_a1(spec["cell"])
        cell = sheet.getCellByPosition(col, row)
        shape.Position = cell.Position
        shape.Anchor = cell
        shape.ResizeWithCell = spec.get("anchor") == "twoCell"

    def op_comment(self, op):
        spec = op.get("comment") or {}
        sheet = self.sheet(op["sheet"])
        col, row = parse_a1(spec["cell"])
        sheet.Annotations.insertNew(
            CellAddress(self.sheet_index(op["sheet"]), col, row),
            str(spec.get("text") or ""),
        )
        # author: XSheetAnnotation exposes Author read-only; new notes carry the
        # profile's user name.

    def op_freeze(self, op):
        settings = op.get("settings") or {}
        sheet = self.sheet(op["sheet"])
        mode = settings.get("mode")
        controller = self.doc.getCurrentController()
        if mode == "freeze":
            col, row = parse_a1(settings.get("top_left_cell") or "A1")
            controller.setActiveSheet(sheet)
            controller.freezeAtPosition(col, row)
        elif mode == "split":
            # OOXML split offsets are twips; XViewSplitable takes pixels at the
            # view zoom (100% => 1440 twips per 96 px).
            x_pixels = round(int(settings.get("x_split") or 0) * 96 / 1440)
            y_pixels = round(int(settings.get("y_split") or 0) * 96 / 1440)
            controller.setActiveSheet(sheet)
            controller.splitAtPosition(x_pixels, y_pixels)
        else:
            raise ValueError(f"unknown freeze mode {mode!r}")

    def op_named_range(self, op):
        spec = op.get("named_range") or {}
        scope = spec.get("scope") or "workbook"
        container = (
            self.sheet(op["sheet"]).NamedRanges if scope == "sheet" else self.doc.NamedRanges
        )
        base = CellAddress(self.sheet_index(op["sheet"]), 0, 0)
        content = self.grammar.ooxml_to_calc(strip_eq(spec["refers_to"]), base)
        container.addNewByName(str(spec["name"]), content, base, 0)

    def op_table(self, op):
        spec = op.get("table") or {}
        sheet = self.sheet(op["sheet"])
        area = sheet.getCellRangeByName(spec["ref"]).RangeAddress
        ranges = self.doc.DatabaseRanges
        ranges.addNewByName(str(spec["name"]), area)
        db_range = ranges.getByName(str(spec["name"]))
        db_range.ContainsHeader = spec.get("header_row") is not False
        if spec.get("totals_row") is not None:
            db_range.TotalsRow = bool(spec["totals_row"])
        if spec.get("style"):
            db_range.TableStyleName = str(spec["style"])
            db_range.UseRowStripes = True
        if spec.get("autofilter"):
            db_range.AutoFilter = True

    def op_protection(self, op):
        settings = op.get("settings") or {}
        sheet = self.sheet(op["sheet"])
        if settings.get("protected") is False:
            if sheet.isProtected():
                sheet.unprotect(str(settings.get("password") or ""))
            return
        sheet.protect(str(settings.get("password") or ""))
        # The granular flags (format_cells, insert_rows, select_*, sort,
        # auto_filter) have no UNO API: XProtectable only takes a password.

    def page_style(self, sheet_name):
        sheet = self.sheet(sheet_name)
        styles = self.doc.StyleFamilies.getByName("PageStyles")
        name = f"ExcelBench_{sheet_name}"
        if not styles.hasByName(name):
            styles.insertByName(name, self.doc.createInstance("com.sun.star.style.PageStyle"))
            sheet.PageStyle = name
        return sheet, styles.getByName(name)

    def op_page_setup(self, op):
        settings = op.get("settings") or {}
        sheet, style = self.page_style(op["sheet"])
        orientation = settings.get("orientation")
        if orientation is not None:
            landscape = orientation == "landscape"
            if bool(style.IsLandscape) != landscape:
                width, height = style.Width, style.Height
                style.IsLandscape = landscape
                style.Width, style.Height = height, width
        if settings.get("fit_to_width") is not None:
            style.ScaleToPagesX = int(settings["fit_to_width"])
        if settings.get("fit_to_height") is not None:
            style.ScaleToPagesY = int(settings["fit_to_height"])
        if settings.get("scale") is not None:
            style.PageScale = int(settings["scale"])
        if settings.get("print_title_rows") is not None:
            _, r1, _, r2 = parse_a1_range(
                "A" + str(settings["print_title_rows"]).replace("$", "").replace(":", ":A")
            )
            sheet.setTitleRows(CellRangeAddress(self.sheet_index(op["sheet"]), 0, r1, 0, r2))
            sheet.setPrintTitleRows(True)
        for key, on_prop, content_prop in (
            ("header_center", "HeaderIsOn", "RightPageHeaderContent"),
            ("footer_center", "FooterIsOn", "RightPageFooterContent"),
        ):
            if settings.get(key) is None:
                continue
            style.setPropertyValue(on_prop, True)
            content = style.getPropertyValue(content_prop)
            self.fill_header_text(content.CenterText, str(settings[key]))
            style.setPropertyValue(content_prop, content)

    def fill_header_text(self, text, value):
        """Write OOXML header/footer text, turning &P/&N/&D/&T/&A/&F into fields."""
        text.setString("")
        cursor = text.createTextCursor()
        buffer = []
        index = 0
        while index < len(value):
            char = value[index]
            if char == "&" and index + 1 < len(value):
                code = value[index + 1]
                if code == "&":
                    buffer.append("&")
                    index += 2
                    continue
                service = HF_CODE_SERVICES.get(code.upper())
                if service is not None:
                    if buffer:
                        text.insertString(cursor, "".join(buffer), False)
                        buffer = []
                    field = self.doc.createInstance(service)
                    text.insertTextContent(cursor, field, False)
                    index += 2
                    continue
            buffer.append(char)
            index += 1
        if buffer:
            text.insertString(cursor, "".join(buffer), False)

    def op_chart(self, op):
        spec = op.get("chart") or {}
        sheet = self.sheet(op["sheet"])
        data_ref = spec.get("data_ref")
        if not data_ref:
            raise ValueError("chart op requires data_ref")
        data_sheet = data_ref.split("!")[0].strip("'") if "!" in data_ref else op["sheet"]
        data_area = (
            self.sheet(data_sheet)
            .getCellRangeByName(data_ref.split("!")[-1].replace("$", ""))
            .RangeAddress
        )
        from_col, from_row = parse_a1(spec.get("from") or "B2")
        to_col, to_row = parse_a1(spec.get("to") or "H12")
        start = sheet.getCellByPosition(from_col, from_row).Position
        end = sheet.getCellByPosition(to_col, to_row).Position
        rect = Rectangle(start.X, start.Y, max(end.X - start.X, 1), max(end.Y - start.Y, 1))
        charts = sheet.Charts
        name = f"Chart{charts.Count + 1}"
        while charts.hasByName(name):
            name += "_"
        charts.addNewByName(name, rect, (data_area,), False, False)
        chart_doc = charts.getByName(name).EmbeddedObject
        chart_type = str(spec.get("type") or "bar").lower()
        if chart_type == "line":
            chart_doc.setDiagram(chart_doc.createInstance("com.sun.star.chart.LineDiagram"))
        elif chart_type != "bar":
            raise ValueError(f"unsupported chart type {chart_type!r}")
        categories = spec.get("categories_ref")
        if categories:
            cat_sheet = categories.split("!")[0].strip("'") if "!" in categories else op["sheet"]
            cat_ref = categories.split("!")[-1].replace("$", "")
            c1, r1, c2, r2 = parse_a1_range(cat_ref)
            representation = f"${cat_sheet}.${col_letter(c1)}${r1 + 1}:${col_letter(c2)}${r2 + 1}"
            provider = chart_doc.getDataProvider()
            sequence = provider.createDataSequenceByRangeRepresentation(representation)
            labeled = self.helper.create("com.sun.star.chart2.data.LabeledDataSequence")
            labeled.setValues(sequence)
            diagram = chart_doc.getFirstDiagram()
            axis = diagram.getCoordinateSystems()[0].getAxisByDimension(0, 0)
            scale = axis.getScaleData()
            scale.Categories = labeled
            axis.setScaleData(scale)
        page = sheet.DrawPage
        for index in range(page.Count):
            shape = page.getByIndex(index)
            if shape.ShapeType == "com.sun.star.drawing.OLE2Shape" and shape.PersistName == name:
                shape.Anchor = sheet.getCellByPosition(from_col, from_row)
                shape.ResizeWithCell = True
                shape.Position = Point(start.X, start.Y)
                break

    def op_pivot(self, op):
        spec = op.get("pivot") or {}
        source = spec.get("source_range") or spec.get("data_range")
        if not source:
            raise ValueError("pivot op requires source_range")
        source_sheet = source.split("!")[0].strip("'") if "!" in source else op["sheet"]
        source_area = (
            self.sheet(source_sheet)
            .getCellRangeByName(source.split("!")[-1].replace("$", ""))
            .RangeAddress
        )
        target = spec.get("target_cell") or spec.get("range")
        if not target:
            raise ValueError("pivot op requires target_cell")
        target_sheet = target.split("!")[0].strip("'") if "!" in target else op["sheet"]
        t_col, t_row = parse_a1(target.split("!")[-1].split(":")[0])
        pivots = self.sheet(target_sheet).DataPilotTables
        descriptor = pivots.createDataPilotDescriptor()
        descriptor.setSourceRange(source_area)
        fields = descriptor.DataPilotFields
        orientation = "com.sun.star.sheet.DataPilotFieldOrientation"
        for key, kind in (
            ("row_fields", "ROW"),
            ("column_fields", "COLUMN"),
            ("filter_fields", "PAGE"),
            ("data_fields", "DATA"),
        ):
            for field_name in spec.get(key) or []:
                field = fields.getByName(str(field_name))
                field.Orientation = _enum(orientation, kind)
                if kind == "DATA":
                    field.Function = _enum("com.sun.star.sheet.GeneralFunction", "SUM")
        name = str(spec.get("name") or f"PivotTable{pivots.Count + 1}")
        pivots.insertNewByName(
            name, CellAddress(self.sheet_index(target_sheet), t_col, t_row), descriptor
        )
