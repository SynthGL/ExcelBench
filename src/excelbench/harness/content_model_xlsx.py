# Vendored verbatim from officelibs-corpus/src/officecorpus/model/xlsx.py
# sha256 17fb55bc132caf91a23f07e1df35423dda7f5e14051837f88dd708ccf7b19e61
# Do not edit here; update the source repo and re-vendor.
"""Spreadsheet (xlsx) content model: what a user of the workbook can observe.

`extract(path)` reads an OOXML spreadsheet package and returns a deterministic,
JSON-serializable model whose top-level keys are `FEATURES`. `compare(before,
after)` lists the differences between two models; an empty list means the
content was preserved. The model never records how a writer serialized the
content (part names, relationship ids, attribute order, shared vs inline
strings, style table layout, default values spelled out or omitted), so a
library that rewrites a workbook into an equivalent form scores exactly the
same as one that copies bytes.

Shared cross-format rule (identical in the xlsx, docx and pptx models)
- Presentation-only lengths compare equal when they differ by at most 0.5 pt.
  Color modifiers (tint, shade, lumMod-like values on a 0..1 or -1..1 scale)
  compare by rendered effect: equal when they differ by at most 1/510, so
  they cannot change an 8-bit channel by a full level. Other
  presentation-only floats that are not lengths (shading percentages,
  gradient geometry, scale factors) compare equal within a relative 1e-9.
  Everything else, including cell values, formula text, number formats, font
  sizes and every non-presentation number, compares exactly: data is data.
- xlsx lengths: column widths are in characters of the workbook default
  font, converted at 1 character = 7 px = 5.25 pt (tolerance 0.5 / 5.25 =
  0.095 characters); row heights (including the default row height) are in
  points (0.5); page margins and header/footer margins are in inches
  (0.5 / 72). Drawing anchor offsets are not part of the model (images and
  charts are anchored by cell), so the 6350 EMU tolerance has nothing to
  apply to.
- xlsx color modifiers: the `~tint:` of any color (an absent tint is 0, so
  `theme:1` equals `theme:1~tint:0.001`). xlsx relative floats: gradient
  fill parameters (degree, left/right/top/bottom, stop positions).
- Content-addressed entries (conditional formats, data validations) whose
  values are equal under these tolerances count as the same entry even
  though their exact hashes differ.
- Equivalent spellings compare equal because the model stores one canonical
  form. xlsx: SpreadsheetML style colors (fonts, fills, borders, dxf,
  conditional-format colors, color scales, data bars, tab colors, rich-text
  runs, indexed palette, theme scheme colors) are 6-digit uppercase RGB,
  because Excel ignores the alpha byte (`001F4E78` == `FF1F4E78` ==
  `1f4e78`); an indexed color equals the explicit RGB its palette entry
  holds; `auto`, indexed 64 and an absent color are all automatic; a tint of
  0 equals no tint. The model records no DrawingML colors, where alpha
  (a:alpha and other srgbClr modifiers) is real and would be kept.

Package rules
- Parts are matched case-insensitively with percent-decoding. Relationships are
  followed by type (the last path segment of the relationship type URI, so
  transitional and strict URIs agree) and resolved target. A relationship id is
  only used to dereference an `r:id` attribute inside the part that owns it;
  no rId value ever enters the model.
- Element and attribute names are compared by local name, so transitional and
  strict namespaces produce the same model.
- A file that is not a zip archive, has no workbook part, or contains XML that
  cannot be parsed raises `ValueError`. Relationships whose target part is
  absent are ignored (nothing observable is behind them).

Equivalence rules
- Absent == default. Every attribute with a documented default is either
  filled with it or dropped when equal to it, so writing a default explicitly
  never differs from omitting it.
- Dicts are keyed (set or map semantics, order ignored); lists are ordered and
  used only where order is user-visible (sheet order, rich-text runs, chart
  plots and series, table columns, comment threads, gradient stops, print
  areas, conditional-format formulas).
- Numbers are parsed to float and compared exactly, except for the
  presentation tolerances in the shared rule above.
- Colors: rgb as uppercase 6-digit RGB (the alpha byte of an 8-digit ARGB is
  dropped, see equivalent spellings), theme colors stay symbolic
  (`theme:N`), indexed colors resolve to RGB
  through the workbook's palette (custom indexedColors entries override the
  ECMA default palette by index; out-of-range indexes stay
  `indexed:N`; 65 is `system-background`), a non-zero tint is appended as
  `~tint:<float>`, and `auto`, indexed 64 (system foreground) and absent all
  mean automatic (None). Referenced theme colors are modeled under `theme`
  so symbolic colors stay comparable.
- Strings decode the OOXML `_xHHHH_` escape. Phonetic runs (`rPh`) are not
  text.
- Formulas: leading/trailing whitespace and one leading `=` are removed,
  everything outside string literals is upper-cased (Excel formulas are
  case-insensitive outside literals), unquoted sheet prefixes are quoted
  (`Sheet1!A1` == `'Sheet1'!A1`), and whitespace outside literals is dropped
  unless it separates two operands (the intersection operator), where it
  becomes one space. String literals and structured-reference brackets are
  kept verbatim apart from upper-casing the brackets. Shared formulas are
  expanded per cell by shifting the anchor formula's relative A1 references,
  so a writer that de-shares formulas compares equal.
- Ranges: `$` dropped, upper-cased, corners ordered, a one-cell range is the
  cell, a full-height range is a column range (`A:A`), a full-width range is a
  row range (`1:1`). Multi-area sqrefs (data validations, conditional
  formats) are canonicalized as the cell-set union decomposed into maximal
  row bands, so any split/merge/reorder of the same cells compares equal.

Feature rules
- workbook: sheet order (list of names), per sheet kind (worksheet,
  chartsheet, dialogsheet, macrosheet, or `missing` if the part is absent)
  and visibility (`visible` default), `date1904`.
- defined_names: keyed `<scope>!<name>` (scope is the sheet name for
  sheet-local names, empty for workbook names) -> normalized refers-to plus
  `hidden` when true. `_xlnm._FilterDatabase` is excluded (derived from the
  auto filter, modeled under auto_filters); `_xlnm.Print_Titles` and
  `_xlnm.Print_Area` are modeled under page_setup; other built-in names stay.
- cells: per sheet, address -> `{"type", "value"}` with type `n` (float), `s`
  (shared, inline and formula-string results all resolve to plain text), `b`
  (bool), `e` (error text). `t="d"` ISO dates become the numeric serial for the
  workbook's date system (1900 system keeps Excel's 1900-02-29 slot). Rich text
  adds `runs`: `[{"text", "font"}]` where each run's font is the cell font
  overridden by the run properties, adjacent runs with equal fonts merged; runs
  are dropped when every run renders in the cell font. Formula cells carry
  their cached result here (missing cached value == no cell entry); unparsable
  values are recorded as type `invalid`.
- formulas: per sheet, address -> `{"text"}` plus `array` (canonical ref) for
  array/dynamic-array anchors, or `data_table` (its attributes) for what-if
  tables.
- cell_styles: `default` is the resolved style of cellXfs[0]. Per sheet,
  address -> resolved style for every cell whose style differs from the style
  an absent cell would have there (the row style if the row has
  customFormat, else the column style, else the default). A style is
  `number_format` (format code; ECMA built-in ids and Excel's en-US codes
  for 41-44 mapped to their codes, other
  built-in ids kept as `builtin:N`; a backslash escape of a character that
  displays literally anyway, such as `\\-`, is dropped), `font` (name, size,
  bold, italic, underline, strike, color, vert_align), `fill` (pattern none
  drops colors, solid keeps only fg; gradients keep type, degree, edges,
  stops), `border`
  (left/right/top/bottom, style none drops color, `start`/`end` == left/right,
  diagonal only when diagonalUp/Down), `alignment` and `protection`, all with
  defaults filled. cellXfs values are used as-is (apply* flags ignored);
  font family/charset/scheme are not modeled. Merged areas: the anchor
  cell holds what the area renders, the anchor style with its border
  replaced by the per-cell perimeter sides (top of the top row, bottom of
  the bottom row, left/right of the outer columns; a uniform side collapses
  to one value); styles of covered cells and interior borders are not
  rendered and are dropped (covered cell values stay under cells). In the
  model, styles are content-addressed keys into `cell_styles.table`;
  `compare` expands them so differences are reported per style field and
  the table itself is never compared.
- merged_cells: per sheet, set of canonical ranges.
- columns: per sheet, maximal runs of adjacent columns with identical
  non-default properties (`width` when given, `hidden`, `outline_level`,
  `collapsed`, `style` when not the default), keyed `A:C`. `compare`
  re-keys both sides onto their common run boundaries first, so runs that
  split differently are compared column range by column range.
- rows: per sheet, row number -> non-default properties: `height` only when
  customHeight (a non-custom `ht` is an autofit result the application
  recomputes, so it is not compared), `hidden`, `outline_level`, `collapsed`,
  `style` when
  customFormat and not the default. Plus `_default` for a custom default row
  height / zeroHeight from sheetFormatPr.
- views: per sheet, first sheet view: frozen pane (`freeze` cols/rows from
  xSplit/ySplit when the pane state is frozen/frozenSplit), right_to_left,
  show_gridlines, show_headers, zoom, tab_color.
- auto_filters: per sheet, `ref` and filter column criteria (canonical dump
  keyed by column id).
- data_validations: per sheet, validations grouped by their parameters (type
  default `none`, operator default `between`, formulas, allow_blank,
  show_dropdown, input/error message flags, texts, error style) with the
  union of their sqrefs. Main and x14 extension validations are one set.
- conditional_formats: per sheet, set of rules (sqref, type, operator,
  formulas, priority rank within the sheet, differential style, stop_if_true,
  type-specific attributes with defaults dropped, color scale / data bar /
  icon set parameters, and linked x14 extension parameters). Priorities are
  replaced by their rank so renumbering is invisible. Only what the rule
  type reads is kept: operator for cellIs and the text rules, text for the
  text rules, timePeriod, rank/percent/bottom for top10,
  aboveAverage/equalAverage/stdDev for aboveAverage; formulas are dropped
  for aboveAverage, top10, duplicate/unique values, color scales, data bars
  and icon sets; cfvo `val` is dropped for min/max thresholds. Unlinked x14
  rules are too. Differential styles keep only what they set: a dxf fill
  without patternType is solid and its color is bgColor (else fgColor), a
  dxf border side without a style sets nothing, a dxf diagonal counts only
  with diagonalUp/Down. Parameter dumps (and auto filter criteria) map
  boolean `true`/`false` to `1`/`0` and drop documented defaults (data bar
  lengths, direction, axis, gradient, border, negative-bar flags; icon set
  name, custom, showValue, percent, reverse; cfvo gte).
- hyperlinks: per sheet, canonical ref -> external `target` and/or in-workbook
  `location` and `tooltip`.
- comments: per sheet, `notes` (address -> author, text) and `threads`
  (address -> ordered list of author display name, text). A legacy note on
  a cell that has a thread is the thread's compatibility shadow and is
  dropped.
- protection: workbook lock_structure/lock_windows; per sheet the
  sheetProtection `sheet` flag (chartsheets: `content`).
- page_setup: per worksheet orientation (`default` == portrait), paper_size,
  fit_to_page with fit_width/fit_height (or scale when not fitting), margins
  with Excel defaults, header/footer texts (`_xHHHH_` decoded, split into
  left/center/right sections by &L/&C/&R; text before any section code is
  center; empty sections dropped) and flags, print_titles
  (`rows`/`cols`), print_area (ordered list of canonical refs).
- images: per sheet, multiset `<anchor>|<sha256 of media bytes>` -> count,
  anchor being the from-cell of drawing anchors, `absolute`, or `background`
  for the sheet background picture. AlternateContent uses its first Choice.
- charts: per sheet, keyed by anchor (`#n` suffix for duplicates; chartsheets
  use `sheet`): plot types (bar charts carry their direction), title
  (`text`, `ref`, or `auto` when a title element has no text, which Excel
  fills automatically), series (name, values, categories, x/y, bubble sizes;
  references normalized and `$`-free) ordered by series order. ChartEx
  charts record their series layout ids.
- tables: per sheet, display name -> ref, columns, header/totals flags,
  style name and stripe flags.
- pivot_tables: per sheet, name -> location and cache source.
- external_links: number of external workbook links.
- vba: presence of a VBA project part.
- theme: `colors`: theme color index -> scheme color (`sysClr` resolves to
  its lastClr RGB; indexes 0-3 map to lt1, dk1, lt2, dk2) for the theme
  indexes the model actually references. Unreferenced theme colors and theme
  fonts are not observable through modeled content.

Not modeled: VML shapes, form controls, header/footer images, slicers,
sparklines, chart formatting beyond types/series/title, calculation chain,
document properties, workbook views, and the default column width (it
depends on font metrics).
"""

from __future__ import annotations

import contextvars
import datetime as _dt
import functools
import hashlib
import itertools
import json
import posixpath
import re
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Any
from urllib.parse import unquote

FEATURES: tuple[str, ...] = (
    "workbook",
    "defined_names",
    "cells",
    "formulas",
    "cell_styles",
    "merged_cells",
    "columns",
    "rows",
    "views",
    "auto_filters",
    "data_validations",
    "conditional_formats",
    "hyperlinks",
    "comments",
    "protection",
    "page_setup",
    "images",
    "charts",
    "tables",
    "pivot_tables",
    "external_links",
    "vba",
    "theme",
)

MAX_ROW = 1048576
MAX_COL = 16384

# ---------------------------------------------------------------------------
# XML / package helpers
# ---------------------------------------------------------------------------


def _ln(tag: Any) -> str:
    return tag.rsplit("}", 1)[-1] if isinstance(tag, str) else ""


def _kids(el: Any, name: str) -> list[Any]:
    return [c for c in el if _ln(c.tag) == name] if el is not None else []


def _kid(el: Any, name: str) -> Any:
    if el is None:
        return None
    for c in el:
        if _ln(c.tag) == name:
            return c
    return None


def _attr(el: Any, local: str) -> str | None:
    """Attribute by local name (any namespace, unqualified first)."""
    if el is None:
        return None
    value = el.get(local)
    if value is not None:
        return value
    for key, val in el.attrib.items():
        if key.startswith("{") and _ln(key) == local:
            return val
    return None


def _rel_attr(el: Any, local: str) -> str | None:
    """A relationship-namespace attribute such as r:id or r:embed."""
    for key, val in el.attrib.items():
        if key.startswith("{") and _ln(key) == local:
            ns = key[1:].split("}", 1)[0]
            if ns.endswith("relationships"):
                return val
    return None


def _bool(value: str | None, default: bool) -> bool:
    if value is None:
        return default
    v = value.strip().lower()
    if v in ("1", "true", "on"):
        return True
    if v in ("0", "false", "off"):
        return False
    return default


def _float(value: str | None, default: float | None = None) -> float | None:
    if value is None:
        return default
    try:
        return float(value)
    except ValueError:
        return default


def _int(value: str | None, default: int | None = None) -> int | None:
    if value is None:
        return default
    try:
        return int(float(value))
    except ValueError:
        return default


_ESCAPE_RE = re.compile(r"_x([0-9A-Fa-f]{4})_")


def _unescape(text: str) -> str:
    return _ESCAPE_RE.sub(lambda m: chr(int(m.group(1), 16)), text)


def _hf_sections(text: str) -> dict[str, str]:
    """Split header/footer text into its &L/&C/&R sections (default: center)."""
    sections = {"left": "", "center": "", "right": ""}
    current = "center"
    i = 0
    while i < len(text):
        if text[i] == "&" and i + 1 < len(text) and text[i + 1] in "LCR":
            current = {"L": "left", "C": "center", "R": "right"}[text[i + 1]]
            i += 2
            continue
        step = 2 if text[i] == "&" and i + 1 < len(text) else 1
        sections[current] += text[i : i + step]
        i += step
    return {k: v for k, v in sections.items() if v}


def _text_of(el: Any) -> str:
    """Plain text of a string item / comment text: t + r/t, skipping rPh."""
    if el is None:
        return ""
    parts: list[str] = []
    for c in el:
        name = _ln(c.tag)
        if name == "t":
            parts.append(c.text or "")
        elif name == "r":
            t = _kid(c, "t")
            parts.append((t.text or "") if t is not None else "")
    if not parts and _ln(el.tag) == "t":
        parts.append(el.text or "")
    return _unescape("".join(parts))


def _key(value: Any) -> str:
    raw = json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=True)
    return hashlib.sha256(raw.encode()).hexdigest()[:16]


class _Rel:
    __slots__ = ("external", "kind", "target")

    def __init__(self, kind: str, target: str, external: bool) -> None:
        self.kind = kind
        self.target = target
        self.external = external


class _Package:
    def __init__(self, path: Path) -> None:
        try:
            archive = zipfile.ZipFile(path)
        except (zipfile.BadZipFile, OSError, ValueError) as exc:
            raise ValueError(f"not a zip archive: {exc}") from exc
        self.parts: dict[str, bytes] = {}
        with archive:
            for info in archive.infolist():
                if info.is_dir():
                    continue
                key = unquote(info.filename.lstrip("/")).lower()
                if key in self.parts:
                    raise ValueError(f"duplicate part name: {info.filename}")
                try:
                    self.parts[key] = archive.read(info)
                except Exception as exc:  # zlib, CRC, encryption
                    raise ValueError(
                        f"unreadable zip member {info.filename}: {exc}"
                    ) from exc
        self._xml: dict[str, Any] = {}
        self._rels: dict[str, dict[str, _Rel]] = {}

    def xml(self, key: str) -> Any:
        if key not in self._xml:
            try:
                self._xml[key] = ET.fromstring(self.parts[key])
            except ET.ParseError as exc:
                raise ValueError(f"malformed XML in {key}: {exc}") from exc
        return self._xml[key]

    def rels(self, source: str) -> dict[str, _Rel]:
        """Relationship id -> relationship, for dereferencing r:id in `source`."""
        if source in self._rels:
            return self._rels[source]
        if source == "":
            rels_key = "_rels/.rels"
        else:
            directory, base = posixpath.split(source)
            rels_key = posixpath.join(directory, "_rels", base + ".rels")
        out: dict[str, _Rel] = {}
        if rels_key in self.parts:
            for child in self.xml(rels_key):
                if _ln(child.tag) != "Relationship":
                    continue
                kind = child.get("Type", "").rstrip("/").rsplit("/", 1)[-1]
                target = child.get("Target", "")
                external = child.get("TargetMode", "Internal") == "External"
                if not external:
                    target = unquote(target.split("#", 1)[0])
                    if target.startswith("/"):
                        target = target.lstrip("/")
                    else:
                        target = posixpath.normpath(
                            posixpath.join(posixpath.dirname(source), target)
                        )
                    target = target.lower()
                out[child.get("Id", "")] = _Rel(kind, target, external)
        self._rels[source] = out
        return out

    def targets(self, source: str, kind: str) -> list[str]:
        """Existing internal parts reached from `source` by relationship type."""
        return sorted(
            r.target
            for r in self.rels(source).values()
            if r.kind == kind and not r.external and r.target in self.parts
        )

    def deref(self, source: str, rid: str | None) -> _Rel | None:
        if not rid:
            return None
        return self.rels(source).get(rid)


# ---------------------------------------------------------------------------
# Cell references and ranges
# ---------------------------------------------------------------------------

_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([0-9]+)$")
_COLS_RE = re.compile(r"^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$")
_ROWS_RE = re.compile(r"^\$?([0-9]+):\$?([0-9]+)$")


@functools.cache
def _col_num(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + ord(ch) - 64
    return n


@functools.cache
def _col_name(n: int) -> str:
    out = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def _parse_cell(text: str) -> tuple[int, int] | None:
    m = _CELL_RE.match(text.strip())
    if not m:
        return None
    return int(m.group(2)), _col_num(m.group(1))


def _parse_range(text: str) -> tuple[int, int, int, int] | None:
    """(r1, c1, r2, c2) for A1, A1:B2, A:B, 1:2; None if not a plain range."""
    text = text.strip()
    if ":" in text:
        a, _, b = text.partition(":")
        pa, pb = _parse_cell(a), _parse_cell(b)
        if pa and pb:
            r1, r2 = sorted((pa[0], pb[0]))
            c1, c2 = sorted((pa[1], pb[1]))
            return r1, c1, r2, c2
        m = _COLS_RE.match(text)
        if m:
            c1, c2 = sorted((_col_num(m.group(1)), _col_num(m.group(2))))
            return 1, c1, MAX_ROW, c2
        m = _ROWS_RE.match(text)
        if m:
            r1, r2 = sorted((int(m.group(1)), int(m.group(2))))
            return r1, 1, r2, MAX_COL
        return None
    p = _parse_cell(text)
    return (p[0], p[1], p[0], p[1]) if p else None


def _render_range(r1: int, c1: int, r2: int, c2: int) -> str:
    if r1 == 1 and r2 == MAX_ROW:
        return f"{_col_name(c1)}:{_col_name(c2)}"
    if c1 == 1 and c2 == MAX_COL:
        return f"{r1}:{r2}"
    if r1 == r2 and c1 == c2:
        return f"{_col_name(c1)}{r1}"
    return f"{_col_name(c1)}{r1}:{_col_name(c2)}{r2}"


def _canon_ref(text: str | None) -> str | None:
    if text is None:
        return None
    rect = _parse_range(text)
    return _render_range(*rect) if rect else text.replace("$", "").strip().upper()


def _canon_sqref(text: str | None) -> str:
    """Canonical cell-set union of a space-separated multi-area reference."""
    rects: list[tuple[int, int, int, int]] = []
    other: list[str] = []
    for token in (text or "").split():
        rect = _parse_range(token)
        if rect:
            rects.append(rect)
        else:
            other.append(token.replace("$", "").upper())
    if not rects:
        return " ".join(sorted(set(other)))
    edges = sorted({r[0] for r in rects} | {r[2] + 1 for r in rects})
    bands: list[list[Any]] = []
    for lo, hi in itertools.pairwise(edges):
        spans = sorted(
            (c1, c2) for r1, c1, r2, c2 in rects if r1 <= lo and r2 >= hi - 1
        )
        merged: list[list[int]] = []
        for c1, c2 in spans:
            if merged and c1 <= merged[-1][1] + 1:
                merged[-1][1] = max(merged[-1][1], c2)
            else:
                merged.append([c1, c2])
        cols = tuple(tuple(m) for m in merged)
        if not cols:
            continue
        if bands and bands[-1][1] == lo - 1 and bands[-1][2] == cols:
            bands[-1][1] = hi - 1
        else:
            bands.append([lo, hi - 1, cols])
    out = sorted((r1, c1, r2, c2) for r1, r2, cols in bands for c1, c2 in cols)
    return " ".join([_render_range(*r) for r in out] + sorted(set(other)))


# ---------------------------------------------------------------------------
# Formulas
# ---------------------------------------------------------------------------

_ERROR_RE = re.compile(
    r"#(?:NULL!|DIV/0!|VALUE!|REF!|NAME\?|NUM!|N/A|GETTING_DATA|SPILL!|CALC!"
    r"|FIELD!|BLOCKED!|UNKNOWN!|CONNECT!|BUSY!|PYTHON!|EXTERNAL!)",
    re.IGNORECASE,
)


def _skip_quoted(text: str, i: int, quote: str) -> int:
    j = i + 1
    n = len(text)
    while j < n:
        if text[j] == quote:
            if j + 1 < n and text[j + 1] == quote:
                j += 2
                continue
            return j + 1
        j += 1
    return n


def _skip_brackets(text: str, i: int) -> int:
    depth = 0
    j = i
    n = len(text)
    while j < n:
        ch = text[j]
        if ch == "'":
            j += 2
            continue
        if ch == "[":
            depth += 1
        elif ch == "]":
            depth -= 1
            if depth == 0:
                return j + 1
        j += 1
    return n


def _operand_end(ch: str) -> bool:
    return ch.isalnum() or ch in "_$.)'\"]?\\%" or ord(ch) > 127


def _operand_start(ch: str) -> bool:
    return ch.isalnum() or ch in "_$.('\"[#\\" or ord(ch) > 127


def _sheet_char(ch: str) -> bool:
    return ch.isalnum() or ch in "_.\\[]:" or ord(ch) > 127


def _normalize_formula(text: str | None) -> str | None:
    if text is None:
        return None
    s = text.strip()
    if s.startswith("="):
        s = s[1:].lstrip()
    out: list[str] = []
    pending_space = False
    i = 0
    n = len(s)

    def emit(token: str) -> None:
        nonlocal pending_space
        if pending_space and out and token:
            prev = out[-1][-1] if out[-1] else ""
            if prev and _operand_end(prev) and _operand_start(token[0]):
                out.append(" ")
        pending_space = False
        out.append(token)

    while i < n:
        ch = s[i]
        if ch == '"':
            j = _skip_quoted(s, i, '"')
            emit(s[i:j])
            i = j
        elif ch == "'":
            j = _skip_quoted(s, i, "'")
            emit(s[i:j].upper())
            i = j
        elif ch == "[":
            j = _skip_brackets(s, i)
            emit(s[i:j].upper())
            i = j
        elif ch.isspace():
            while i < n and s[i].isspace():
                i += 1
            pending_space = True
        elif ch == "#" and _ERROR_RE.match(s, i):
            m = _ERROR_RE.match(s, i)
            emit(m.group(0).upper())
            i = m.end()
        elif ch == "!":
            pending_space = False
            joined = "".join(out)
            if joined and not joined.endswith("'"):
                j = len(joined)
                while j > 0 and _sheet_char(joined[j - 1]):
                    j -= 1
                if j < len(joined) and not (j > 0 and joined[j - 1] == "#"):
                    joined = joined[:j] + "'" + joined[j:] + "'"
            out[:] = [joined, "!"] if joined else ["!"]
            i += 1
        else:
            emit(ch.upper())
            i += 1
    return "".join(out)


_REF_RE = re.compile(
    r"(?<![A-Za-z0-9_.$\\])(?:"
    r"(?P<cd>\$?)(?P<c>[A-Za-z]{1,3})(?P<rd>\$?)(?P<r>[0-9]+)"
    r"|(?P<ad>\$?)(?P<a>[A-Za-z]{1,3}):(?P<bd>\$?)(?P<b>[A-Za-z]{1,3})"
    r"|(?P<xd>\$?)(?P<x>[0-9]+):(?P<yd>\$?)(?P<y>[0-9]+)"
    r")(?![A-Za-z0-9_(.!\\\[])"
)


def _plain_spans(text: str) -> list[tuple[int, int]]:
    spans: list[tuple[int, int]] = []
    start = i = 0
    n = len(text)
    while i < n:
        ch = text[i]
        if ch in "\"'":
            if start < i:
                spans.append((start, i))
            i = _skip_quoted(text, i, ch)
            start = i
        elif ch == "[":
            if start < i:
                spans.append((start, i))
            i = _skip_brackets(text, i)
            start = i
        else:
            i += 1
    if start < n:
        spans.append((start, n))
    return spans


def _translate(text: str, drow: int, dcol: int) -> str:
    """Shift relative A1 references by (drow, dcol), as Excel fills a formula."""
    if not drow and not dcol:
        return text
    spans = _plain_spans(text)
    out: list[str] = []
    last = 0
    for m in _REF_RE.finditer(text):
        if not any(s <= m.start() and m.end() <= e for s, e in spans):
            continue
        if m.group("c") is not None:
            col, row = _col_num(m.group("c")), int(m.group("r"))
            if col > MAX_COL or row > MAX_ROW or row < 1:
                continue
            if not m.group("cd"):
                col += dcol
            if not m.group("rd"):
                row += drow
            if not (1 <= col <= MAX_COL and 1 <= row <= MAX_ROW):
                new = "#REF!"
            else:
                new = f"{m.group('cd')}{_col_name(col)}{m.group('rd')}{row}"
        elif m.group("a") is not None:
            a, b = _col_num(m.group("a")), _col_num(m.group("b"))
            if a > MAX_COL or b > MAX_COL:
                continue
            if not m.group("ad"):
                a += dcol
            if not m.group("bd"):
                b += dcol
            if not (1 <= a <= MAX_COL and 1 <= b <= MAX_COL):
                new = "#REF!"
            else:
                new = f"{m.group('ad')}{_col_name(a)}:{m.group('bd')}{_col_name(b)}"
        else:
            x, y = int(m.group("x")), int(m.group("y"))
            if not m.group("xd"):
                x += drow
            if not m.group("yd"):
                y += drow
            if not (1 <= x <= MAX_ROW and 1 <= y <= MAX_ROW):
                new = "#REF!"
            else:
                new = f"{m.group('xd')}{x}:{m.group('yd')}{y}"
        out.append(text[last : m.start()])
        out.append(new)
        last = m.end()
    out.append(text[last:])
    return "".join(out)


def _split_areas(text: str) -> list[str]:
    """Split a refers-to text on top-level commas (outside quotes/brackets)."""
    areas: list[str] = []
    start = i = 0
    depth = 0
    n = len(text)
    while i < n:
        ch = text[i]
        if ch in "\"'":
            i = _skip_quoted(text, i, ch)
            continue
        if ch == "[":
            i = _skip_brackets(text, i)
            continue
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        elif ch == "," and depth == 0:
            areas.append(text[start:i])
            start = i + 1
        i += 1
    areas.append(text[start:])
    return [a.strip() for a in areas if a.strip()]


def _strip_sheet(area: str) -> str:
    i = 0
    cut = 0
    while i < len(area):
        if area[i] == "'":
            i = _skip_quoted(area, i, "'")
            continue
        if area[i] == "!":
            cut = i + 1
        i += 1
    return area[cut:]


def _chart_ref(text: str | None) -> str | None:
    norm = _normalize_formula(text)
    return norm.replace("$", "") if norm is not None else None


# ---------------------------------------------------------------------------
# Styles
# ---------------------------------------------------------------------------

_BUILTIN_NUMFMTS = {
    0: "General",
    1: "0",
    2: "0.00",
    3: "#,##0",
    4: "#,##0.00",
    9: "0%",
    10: "0.00%",
    11: "0.00E+00",
    12: "# ?/?",
    13: "# ??/??",
    14: "mm-dd-yy",
    15: "d-mmm-yy",
    16: "d-mmm",
    17: "mmm-yy",
    18: "h:mm AM/PM",
    19: "h:mm:ss AM/PM",
    20: "h:mm",
    21: "h:mm:ss",
    22: "m/d/yy h:mm",
    37: "#,##0 ;(#,##0)",
    38: "#,##0 ;[Red](#,##0)",
    39: "#,##0.00;(#,##0.00)",
    40: "#,##0.00;[Red](#,##0.00)",
    # 41-44 are absent from the ECMA table; these are the codes Excel (en-US)
    # writes for them.
    41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
    42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
    43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
    44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
    45: "mm:ss",
    46: "[h]:mm:ss",
    47: "mmss.0",
    48: "##0.0E+0",
    49: "@",
}

# ECMA-376 Part 1, 18.8.27: the default indexed color palette (indexes 0-63).
_DEFAULT_PALETTE = tuple(
    (
        "000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF "
        "000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF "
        "800000 008000 000080 808000 800080 008080 C0C0C0 808080 "
        "9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCFF "
        "000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF "
        "00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 "
        "3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 "
        "003366 339966 003300 333300 993300 993366 333399 333333"
    ).split()
)

_FONT_DEFAULTS = {
    "name": None,
    "size": None,
    "bold": False,
    "italic": False,
    "underline": "none",
    "strike": False,
    "color": None,
    "vert_align": "baseline",
}
_ALIGN_ATTRS = {
    "horizontal": ("horizontal", "general", str),
    "vertical": ("vertical", "bottom", str),
    "wrapText": ("wrap_text", False, bool),
    "shrinkToFit": ("shrink_to_fit", False, bool),
    "indent": ("indent", 0, int),
    "textRotation": ("text_rotation", 0, int),
    "readingOrder": ("reading_order", 0, int),
    "justifyLastLine": ("justify_last_line", False, bool),
}
_SIDES = ("left", "right", "top", "bottom")


_PALETTE: contextvars.ContextVar[tuple[str, ...]] = contextvars.ContextVar(
    "_PALETTE", default=_DEFAULT_PALETTE
)


def _color(el: Any) -> str | None:
    if el is None or _bool(el.get("auto"), False):
        return None
    rgb, theme, indexed = el.get("rgb"), el.get("theme"), el.get("indexed")
    if rgb:
        # SpreadsheetML ignores the alpha byte of ARGB colors when rendering.
        value = rgb.strip().upper()
        if len(value) == 8:
            value = value[2:]
    elif theme is not None:
        value = f"theme:{_int(theme)}"
    elif indexed is not None:
        index = _int(indexed)
        if index == 64:
            return None  # system foreground == automatic
        if index == 65:
            value = "system-background"
        else:
            palette = _PALETTE.get()
            value = (
                palette[index]
                if index is not None and 0 <= index < len(palette)
                else f"indexed:{indexed}"
            )
    else:
        return None
    tint = _float(el.get("tint"), 0.0)
    if tint:
        value += f"~tint:{tint!r}"
    return value


def _font(el: Any, full: bool) -> dict[str, Any]:
    d: dict[str, Any] = dict(_FONT_DEFAULTS) if full else {}
    for c in el if el is not None else ():
        name, val = _ln(c.tag), c.get("val")
        if name in ("name", "rFont"):
            d["name"] = val
        elif name == "sz":
            d["size"] = _float(val)
        elif name in ("b", "i", "strike"):
            d[{"b": "bold", "i": "italic"}.get(name, name)] = _bool(val, True)
        elif name == "u":
            d["underline"] = val or "single"
        elif name == "vertAlign":
            d["vert_align"] = val or "baseline"
        elif name == "color":
            d["color"] = _color(c)
    return d


def _fill(el: Any, full: bool) -> dict[str, Any]:
    if el is None:
        return {"pattern": "none"} if full else {}
    gf = _kid(el, "gradientFill")
    if gf is not None:
        return {
            "gradient": {
                "type": gf.get("type", "linear"),
                "degree": _float(gf.get("degree"), 0.0),
                "left": _float(gf.get("left"), 0.0),
                "right": _float(gf.get("right"), 0.0),
                "top": _float(gf.get("top"), 0.0),
                "bottom": _float(gf.get("bottom"), 0.0),
                "stops": [
                    [_float(s.get("position"), 0.0), _color(_kid(s, "color"))]
                    for s in _kids(gf, "stop")
                ],
            }
        }
    pf = _kid(el, "patternFill")
    if pf is None:
        return {"pattern": "none"} if full else {}
    pattern = pf.get("patternType")
    fg, bg = _color(_kid(pf, "fgColor")), _color(_kid(pf, "bgColor"))
    if not full:
        # Differential (dxf) fills: an absent patternType means solid, and a
        # solid dxf fill paints its bgColor (fgColor when bgColor is absent).
        pattern = pattern or "solid"
        if pattern == "none":
            return {"pattern": "none"}
        if pattern == "solid":
            return {"pattern": "solid", "color": bg if bg is not None else fg}
        return {"pattern": pattern, "fg": fg, "bg": bg}
    pattern = pattern or "none"
    if pattern == "none":
        return {"pattern": "none"}
    if pattern == "solid":
        return {"pattern": "solid", "fg": fg}
    return {"pattern": pattern, "fg": fg, "bg": bg}


def _border(el: Any, full: bool) -> dict[str, Any]:
    sides: dict[str, Any] = {}
    for c in el if el is not None else ():
        name = {"start": "left", "end": "right"}.get(_ln(c.tag), _ln(c.tag))
        if name in (*_SIDES, "diagonal", "vertical", "horizontal"):
            style = c.get("style")
            if style is None and not full:
                continue  # a differential side without style changes nothing
            sides[name] = (
                {"style": "none"}
                if (style or "none") == "none"
                else {"style": style, "color": _color(_kid(c, "color"))}
            )
    up = _bool(_attr(el, "diagonalUp"), False) if el is not None else False
    down = _bool(_attr(el, "diagonalDown"), False) if el is not None else False
    if not full:
        if up or down:
            sides["diagonal_up"], sides["diagonal_down"] = up, down
        else:
            sides.pop("diagonal", None)
        return sides
    out = {s: sides.get(s, {"style": "none"}) for s in _SIDES}
    if up or down:
        out["diagonal"] = {
            "up": up,
            "down": down,
            **sides.get("diagonal", {"style": "none"}),
        }
    return out


def _alignment(el: Any, full: bool) -> dict[str, Any]:
    out: dict[str, Any] = {}
    for attr, (name, default, kind) in _ALIGN_ATTRS.items():
        raw = el.get(attr) if el is not None else None
        if raw is None:
            if full:
                out[name] = default
            continue
        if kind is bool:
            out[name] = _bool(raw, default)
        elif kind is int:
            out[name] = _int(raw, default)
        else:
            out[name] = raw
    return out


def _protection(el: Any, full: bool) -> dict[str, Any]:
    out: dict[str, Any] = {}
    for attr, default in (("locked", True), ("hidden", False)):
        raw = el.get(attr) if el is not None else None
        if raw is not None or full:
            out[attr] = _bool(raw, default)
    return out


_SELF_LITERAL = frozenset("$-+/():!^&'~{}<>= ")


def _numfmt_code(code: str) -> str:
    """Drop backslash escapes of characters that display literally anyway."""
    out: list[str] = []
    i = 0
    n = len(code)
    while i < n:
        ch = code[i]
        if ch == '"':
            j = code.find('"', i + 1)
            j = n if j < 0 else j + 1
            out.append(code[i:j])
            i = j
        elif ch == "\\" and i + 1 < n:
            nxt = code[i + 1]
            out.append(nxt if nxt in _SELF_LITERAL else code[i : i + 2])
            i += 2
        else:
            out.append(ch)
            i += 1
    return "".join(out)


def _builtin_numfmt(nid: int) -> str:
    code = _BUILTIN_NUMFMTS.get(nid)
    return _numfmt_code(code) if code is not None else f"builtin:{nid}"


class _Styles:
    def __init__(self, root: Any) -> None:
        palette = _kids(_kid(_kid(root, "colors"), "indexedColors"), "rgbColor")
        resolved = list(_DEFAULT_PALETTE)
        for i, c in enumerate(palette[: len(resolved)]):
            rgb = (c.get("rgb") or "").strip().upper()
            resolved[i] = rgb[-6:] if len(rgb) >= 6 else resolved[i]
        # Every color parsed from here on (styles, shared strings, sheets) is
        # resolved through this workbook's palette.
        _PALETTE.set(tuple(resolved))
        self.numfmts: dict[int, str] = {}
        for nf in _kids(_kid(root, "numFmts"), "numFmt"):
            nid = _int(nf.get("numFmtId"))
            if nid is not None:
                self.numfmts[nid] = _numfmt_code(nf.get("formatCode", ""))
        self.fonts = [_font(f, True) for f in _kids(_kid(root, "fonts"), "font")]
        self.fills = [_fill(f, True) for f in _kids(_kid(root, "fills"), "fill")]
        self.borders = [
            _border(b, True) for b in _kids(_kid(root, "borders"), "border")
        ]
        self.xfs = _kids(_kid(root, "cellXfs"), "xf")
        self.dxf_els = _kids(_kid(root, "dxfs"), "dxf")
        self.table: dict[str, dict[str, Any]] = {}
        self._keys: dict[int, str] = {}
        self.default_style = self._resolve(0)
        self.default_key = self.key(0)

    def numfmt(self, nid: int) -> str:
        if nid in self.numfmts:
            return self.numfmts[nid]
        return _builtin_numfmt(nid)

    @staticmethod
    def _pick(table: list[Any], index: int | None, default: Any) -> Any:
        if index is None:
            index = 0
        if 0 <= index < len(table):
            return table[index]
        return default if index == 0 and not table else f"invalid:{index}"

    def _resolve(self, index: int) -> dict[str, Any]:
        if not (0 <= index < len(self.xfs)):
            if index == 0:
                xf = None
            else:
                return {"invalid_xf": index}
        else:
            xf = self.xfs[index]
        get = (lambda a: xf.get(a)) if xf is not None else (lambda a: None)
        return {
            "number_format": self.numfmt(_int(get("numFmtId"), 0) or 0),
            "font": self._pick(self.fonts, _int(get("fontId")), dict(_FONT_DEFAULTS)),
            "fill": self._pick(self.fills, _int(get("fillId")), {"pattern": "none"}),
            "border": self._pick(
                self.borders, _int(get("borderId")), _border(None, True)
            ),
            "alignment": _alignment(_kid(xf, "alignment"), True),
            "protection": _protection(_kid(xf, "protection"), True),
        }

    def key(self, index: int) -> str:
        if index not in self._keys:
            self._keys[index] = self.add(self._resolve(index))
        return self._keys[index]

    def add(self, style: dict[str, Any]) -> str:
        k = _key(style)
        self.table.setdefault(k, style)
        return k

    def font(self, index: int) -> dict[str, Any]:
        font = self.table[self.key(index)].get("font")
        return font if isinstance(font, dict) else dict(_FONT_DEFAULTS)

    def dxf(self, index: int | None) -> dict[str, Any] | None:
        if index is None or not (0 <= index < len(self.dxf_els)):
            return None
        return _dxf(self.dxf_els[index])


def _dxf(el: Any) -> dict[str, Any]:
    out: dict[str, Any] = {}
    font = _kid(el, "font")
    if font is not None:
        out["font"] = _font(font, False)
    nf = _kid(el, "numFmt")
    if nf is not None:
        code = nf.get("formatCode")
        nid = _int(nf.get("numFmtId"), 0) or 0
        out["number_format"] = (
            _numfmt_code(code) if code is not None else _builtin_numfmt(nid)
        )
    fill = _kid(el, "fill")
    if fill is not None:
        out["fill"] = _fill(fill, False)
    align = _kid(el, "alignment")
    if align is not None:
        out["alignment"] = _alignment(align, False)
    border = _kid(el, "border")
    if border is not None:
        out["border"] = _border(border, False)
    prot = _kid(el, "protection")
    if prot is not None:
        out["protection"] = _protection(prot, False)
    return out


# ---------------------------------------------------------------------------
# Generic canonical dump (filter criteria, CF parameters)
# ---------------------------------------------------------------------------

_DUMP_DEFAULTS = {
    ("dataBar", "minLength"): "10",
    ("dataBar", "maxLength"): "90",
    ("dataBar", "showValue"): "1",
    ("dataBar", "border"): "0",
    ("dataBar", "gradient"): "1",
    ("dataBar", "direction"): "context",
    ("dataBar", "negativeBarColorSameAsPositive"): "0",
    ("dataBar", "negativeBarBorderColorSameAsPositive"): "1",
    ("dataBar", "axisPosition"): "automatic",
    ("iconSet", "custom"): "0",
    ("iconSet", "iconSet"): "3TrafficLights1",
    ("iconSet", "showValue"): "1",
    ("iconSet", "percent"): "1",
    ("iconSet", "reverse"): "0",
    ("cfvo", "gte"): "1",
    ("filters", "blank"): "0",
    ("customFilters", "and"): "0",
    ("customFilter", "operator"): "equal",
    ("filterColumn", "hiddenButton"): "0",
    ("filterColumn", "showButton"): "1",
}
_COLOR_TAGS = {
    "color",
    "fgColor",
    "bgColor",
    "negativeFillColor",
    "negativeBorderColor",
    "axisColor",
    "borderColor",
    "fillColor",
}
# cfvo types whose `val` is ignored (the threshold is the data's min/max).
_CFVO_NO_VALUE = ("min", "max", "autoMin", "autoMax")


def _dump(el: Any) -> Any:
    tag = _ln(el.tag)
    if tag in _COLOR_TAGS:
        return {"tag": tag, "color": _color(el)}
    attrs: dict[str, str] = {}
    for key, val in el.attrib.items():
        name = _ln(key)
        if key.startswith("{") and key[1:].split("}", 1)[0].endswith("relationships"):
            continue
        if tag == "cfvo" and name == "val" and el.get("type") in _CFVO_NO_VALUE:
            continue
        val = {"true": "1", "false": "0"}.get(val.lower(), val)
        if _DUMP_DEFAULTS.get((tag, name)) == val:
            continue
        attrs[name] = val
    out: dict[str, Any] = {"tag": tag}
    if attrs:
        out["attrs"] = dict(sorted(attrs.items()))
    text = (el.text or "").strip()
    if text:
        out["text"] = text
    children = [_dump(c) for c in el if isinstance(c.tag, str)]
    if children:
        out["children"] = children
    return out


# ---------------------------------------------------------------------------
# Strings and dates
# ---------------------------------------------------------------------------


def _string_item(si: Any) -> tuple[str, list[tuple[str, dict[str, Any]]] | None]:
    runs: list[tuple[str, dict[str, Any]]] = []
    has_runs = False
    for c in si:
        name = _ln(c.tag)
        if name == "t":
            runs.append((_unescape(c.text or ""), {}))
        elif name == "r":
            has_runs = True
            t = _kid(c, "t")
            runs.append(
                (
                    _unescape((t.text or "") if t is not None else ""),
                    _font(_kid(c, "rPr"), False),
                )
            )
    text = "".join(t for t, _ in runs)
    return text, (runs if has_runs else None)


def _effective_runs(
    runs: list[tuple[str, dict[str, Any]]] | None, cell_font: dict[str, Any]
) -> list[dict[str, Any]] | None:
    if not runs:
        return None
    merged: list[dict[str, Any]] = []
    for text, props in runs:
        if not text:
            continue
        font = {**cell_font, **props}
        if merged and merged[-1]["font"] == font:
            merged[-1]["text"] += text
        else:
            merged.append({"text": text, "font": font})
    if all(r["font"] == cell_font for r in merged):
        return None
    return merged


def _date_serial(text: str, date1904: bool) -> float | None:
    raw = text.strip()
    try:
        if "T" in raw or "-" in raw:
            value = _dt.datetime.fromisoformat(raw).replace(tzinfo=None)
        else:
            t = _dt.time.fromisoformat(raw)
            return (
                t.hour * 3600 + t.minute * 60 + t.second + t.microsecond / 1e6
            ) / 86400
    except ValueError:
        return None
    epoch = _dt.datetime(1904, 1, 1) if date1904 else _dt.datetime(1899, 12, 30)
    delta = value - epoch
    serial = delta.days + delta.seconds / 86400 + delta.microseconds / 86400e6
    if not date1904 and value < _dt.datetime(1900, 3, 1):
        serial -= 1
    return serial


# ---------------------------------------------------------------------------
# Extraction
# ---------------------------------------------------------------------------


def extract(path: Path) -> dict[str, Any]:
    pkg = _Package(Path(path))
    wb_part = next(
        (
            r.target
            for r in pkg.rels("").values()
            if r.kind == "officeDocument" and not r.external and r.target in pkg.parts
        ),
        None,
    )
    if wb_part is None:
        raise ValueError("package has no workbook part")
    wb = pkg.xml(wb_part)
    if _ln(wb.tag) != "workbook":
        raise ValueError(f"main part {wb_part} is not a workbook")

    model: dict[str, Any] = {f: {} for f in FEATURES}
    date1904 = _bool(_attr(_kid(wb, "workbookPr"), "date1904"), False)

    styles_parts = pkg.targets(wb_part, "styles")
    styles = _Styles(pkg.xml(styles_parts[0]) if styles_parts else None)

    shared: list[tuple[str, Any]] = []
    for part in pkg.targets(wb_part, "sharedStrings")[:1]:
        shared = [_string_item(si) for si in _kids(pkg.xml(part), "si")]

    persons: dict[str, str] = {}
    for part in pkg.targets(wb_part, "person"):
        for p in _kids(pkg.xml(part), "person"):
            persons[p.get("id", "")] = p.get("displayName", "")

    # Sheets
    sheets: list[tuple[str, str, str | None]] = []  # name, kind, part
    wb_sheets: dict[str, Any] = {}
    for sh in _kids(_kid(wb, "sheets"), "sheet"):
        name = sh.get("name", "")
        rel = pkg.deref(wb_part, _rel_attr(sh, "id"))
        part = (
            rel.target if rel and not rel.external and rel.target in pkg.parts else None
        )
        kind = {"xlMacrosheet": "macrosheet", "xlIntlMacrosheet": "macrosheet"}.get(
            rel.kind if rel else "", rel.kind if rel else "missing"
        )
        if part is None:
            kind = "missing"
        sheets.append((name, kind, part))
        wb_sheets[name] = {"kind": kind, "state": sh.get("state", "visible")}
    sheet_names = [s[0] for s in sheets]
    model["workbook"] = {
        "sheet_order": sheet_names,
        "sheets": wb_sheets,
        "date1904": date1904,
    }

    # Defined names
    print_titles: dict[str, dict[str, str]] = {}
    print_areas: dict[str, list[str]] = {}
    names: dict[str, Any] = {}
    for dn in _kids(_kid(wb, "definedNames"), "definedName"):
        name = dn.get("name", "")
        local = _int(dn.get("localSheetId"))
        scope = ""
        if local is not None:
            scope = sheet_names[local] if 0 <= local < len(sheet_names) else f"#{local}"
        refers = dn.text or ""
        lname = name.lower()
        if lname == "_xlnm._filterdatabase":
            continue
        if lname == "_xlnm.print_titles" and scope:
            titles: dict[str, str] = {}
            for area in _split_areas(refers):
                rect = _parse_range(_strip_sheet(area))
                if rect and rect[1] == 1 and rect[3] == MAX_COL:
                    titles["rows"] = f"{rect[0]}:{rect[2]}"
                elif rect and rect[0] == 1 and rect[2] == MAX_ROW:
                    titles["cols"] = f"{_col_name(rect[1])}:{_col_name(rect[3])}"
                else:
                    titles["other"] = _normalize_formula(area) or ""
            print_titles[scope] = titles
            continue
        if lname == "_xlnm.print_area" and scope:
            print_areas[scope] = [
                _canon_ref(_strip_sheet(a))
                if _parse_range(_strip_sheet(a))
                else _normalize_formula(a) or ""
                for a in _split_areas(refers)
            ]
            continue
        entry: dict[str, Any] = {"refers_to": _normalize_formula(refers)}
        if _bool(dn.get("hidden"), False):
            entry["hidden"] = True
        names[f"{scope}!{name}"] = entry
    model["defined_names"] = names

    # Workbook-level package features
    model["protection"] = {
        "workbook": {
            "lock_structure": _bool(
                _attr(_kid(wb, "workbookProtection"), "lockStructure"), False
            ),
            "lock_windows": _bool(
                _attr(_kid(wb, "workbookProtection"), "lockWindows"), False
            ),
        },
        "sheets": {},
    }
    model["external_links"] = {"count": len(pkg.targets(wb_part, "externalLink"))}
    model["vba"] = {"present": bool(pkg.targets(wb_part, "vbaProject"))}
    theme_colors = _theme_colors(pkg, wb_part)
    model["theme"] = {}

    model["cell_styles"] = {"default": styles.default_key, "table": {}, "sheets": {}}
    for feature in FEATURES:
        if feature in (
            "workbook",
            "defined_names",
            "external_links",
            "vba",
            "theme",
            "protection",
            "cell_styles",
        ):
            continue
        model[feature] = {name: {} for name in sheet_names}
    for name in sheet_names:
        model["cell_styles"]["sheets"][name] = {}
        model["protection"]["sheets"][name] = False

    for name, kind, part in sheets:
        if part is None:
            continue
        root = pkg.xml(part)
        prot = _kid(root, "sheetProtection")
        if _ln(root.tag) == "chartsheet":
            model["protection"]["sheets"][name] = _bool(_attr(prot, "content"), False)
        else:
            model["protection"]["sheets"][name] = _bool(_attr(prot, "sheet"), False)
        _drawings(pkg, part, root, name, model)
        if _kid(root, "sheetData") is None:
            continue
        _worksheet(
            pkg,
            part,
            root,
            name,
            model,
            styles,
            shared,
            persons,
            date1904,
            print_titles.get(name),
            print_areas.get(name),
        )
    model["cell_styles"]["table"] = dict(sorted(styles.table.items()))
    dumped = json.dumps(model)
    used_theme = sorted({int(m) for m in _THEME_RE.findall(dumped)})
    if used_theme:
        model["theme"]["colors"] = {
            str(i): theme_colors.get(_THEME_SLOTS[i]) if i < len(_THEME_SLOTS) else None
            for i in used_theme
        }
    return model


# SpreadsheetML theme color indexes (ECMA-376 Part 1, 18.3.1.15 / 20.1.6.2):
# the first two light/dark pairs are swapped relative to clrScheme order.
_THEME_SLOTS = (
    "lt1",
    "dk1",
    "lt2",
    "dk2",
    "accent1",
    "accent2",
    "accent3",
    "accent4",
    "accent5",
    "accent6",
    "hlink",
    "folHlink",
)


def _theme_colors(pkg: _Package, wb_part: str) -> dict[str, Any]:
    colors: dict[str, Any] = {}
    parts = pkg.targets(wb_part, "theme")
    if not parts:
        return colors
    scheme = _kid(_kid(pkg.xml(parts[0]), "themeElements"), "clrScheme")
    for slot in scheme if scheme is not None else ():
        value = None
        for c in slot:
            tag = _ln(c.tag)
            if tag == "srgbClr":
                value = (c.get("val") or "").upper()
            elif tag == "sysClr":
                last = c.get("lastClr")
                value = last.upper() if last else f"system:{c.get('val')}"
            else:
                value = f"{tag}:{c.get('val')}"
            break
        colors[_ln(slot.tag)] = value
    return colors


_THEME_RE = re.compile(r'"theme:(\d+)')


def _worksheet(
    pkg: _Package,
    part: str,
    root: Any,
    name: str,
    model: dict[str, Any],
    styles: _Styles,
    shared: list[tuple[str, Any]],
    persons: dict[str, str],
    date1904: bool,
    print_titles: dict[str, str] | None,
    print_area: list[str] | None,
) -> None:
    default_key = styles.default_key

    # Columns
    col_props: dict[int, dict[str, Any]] = {}
    col_style: dict[int, str] = {}
    for col in _kids(_kid(root, "cols"), "col"):
        lo, hi = _int(col.get("min")), _int(col.get("max"))
        if lo is None or hi is None:
            continue
        props: dict[str, Any] = {}
        width = _float(col.get("width"))
        if width is not None:
            props["width"] = width
        if _bool(col.get("hidden"), False):
            props["hidden"] = True
        level = _int(col.get("outlineLevel"), 0)
        if level:
            props["outline_level"] = level
        if _bool(col.get("collapsed"), False):
            props["collapsed"] = True
        skey = styles.key(_int(col.get("style"), 0) or 0)
        if skey != default_key:
            props["style"] = skey
        for c in range(max(lo, 1), min(hi, MAX_COL) + 1):
            if props:
                col_props[c] = props
            if skey != default_key:
                col_style[c] = skey
    columns: dict[str, Any] = {}
    run_start = prev = None
    for c in sorted(col_props) + [None]:
        if run_start is not None and (
            c is None or c != prev + 1 or col_props[c] != col_props[run_start]
        ):
            columns[f"{_col_name(run_start)}:{_col_name(prev)}"] = col_props[run_start]
            run_start = None
        if c is not None and run_start is None:
            run_start = c
        prev = c
    model["columns"][name] = columns

    # Rows and cells
    rows: dict[str, Any] = {}
    row_styles: dict[int, str] = {}
    fmt = _kid(root, "sheetFormatPr")
    if fmt is not None:
        default_row: dict[str, Any] = {}
        if _bool(fmt.get("customHeight"), False) and fmt.get("defaultRowHeight"):
            default_row["height"] = _float(fmt.get("defaultRowHeight"))
        if _bool(fmt.get("zeroHeight"), False):
            default_row["hidden"] = True
        if default_row:
            rows["_default"] = default_row
    cells: dict[str, Any] = {}
    formulas: dict[str, Any] = {}
    cell_styles: dict[str, str] = {}
    shared_anchor: dict[str, tuple[int, int, str]] = {}
    shared_users: list[tuple[str, int, int, str]] = []
    row_num = 0
    for row in _kids(_kid(root, "sheetData"), "row"):
        row_num = _int(row.get("r"), row_num + 1) or row_num + 1
        props = {}
        if _bool(row.get("customHeight"), False) and row.get("ht") is not None:
            props["height"] = _float(row.get("ht"))
        if _bool(row.get("hidden"), False):
            props["hidden"] = True
        level = _int(row.get("outlineLevel"), 0)
        if level:
            props["outline_level"] = level
        if _bool(row.get("collapsed"), False):
            props["collapsed"] = True
        row_key = None
        if _bool(row.get("customFormat"), False):
            row_key = styles.key(_int(row.get("s"), 0) or 0)
            row_styles[row_num] = row_key
            if row_key != default_key:
                props["style"] = row_key
        if props:
            rows[str(row_num)] = props
        col_num = 0
        for c in _kids(row, "c"):
            ref = c.get("r")
            parsed = _parse_cell(ref) if ref else None
            if parsed:
                row_here, col_num = parsed
            else:
                row_here, col_num = row_num, col_num + 1
            addr = f"{_col_name(col_num)}{row_here}"
            s_index = _int(c.get("s"), 0) or 0
            skey = styles.key(s_index)
            background = (
                row_key if row_key is not None else col_style.get(col_num, default_key)
            )
            if skey != background:
                cell_styles[addr] = skey
            value = _cell_value(c, shared, styles.font(s_index), date1904)
            if value is not None:
                cells[addr] = value
            f = _kid(c, "f")
            if f is None:
                continue
            ftype = f.get("t", "normal")
            text = f.text if f.text and f.text.strip() else None
            if ftype == "shared":
                si = f.get("si", "")
                if text is not None and si not in shared_anchor:
                    shared_anchor[si] = (row_here, col_num, text)
                    formulas[addr] = {"text": _normalize_formula(text)}
                elif text is not None:
                    formulas[addr] = {"text": _normalize_formula(text)}
                else:
                    shared_users.append((addr, row_here, col_num, si))
            elif ftype == "array":
                entry: dict[str, Any] = {"text": _normalize_formula(text)}
                entry["array"] = _canon_ref(f.get("ref") or addr)
                formulas[addr] = entry
            elif ftype == "dataTable":
                attrs = {
                    k: v
                    for k, v in sorted(f.attrib.items())
                    if k in ("ref", "dt2D", "dtr", "del1", "del2", "r1", "r2")
                }
                if "ref" in attrs:
                    attrs["ref"] = _canon_ref(attrs["ref"]) or ""
                formulas[addr] = {"text": _normalize_formula(text), "data_table": attrs}
            elif text is not None:
                formulas[addr] = {"text": _normalize_formula(text)}
    for addr, r, col, si in shared_users:
        anchor = shared_anchor.get(si)
        if anchor is None:
            formulas[addr] = {"text": None, "shared_without_anchor": True}
            continue
        formulas[addr] = {
            "text": _normalize_formula(
                _translate(anchor[2], r - anchor[0], col - anchor[1])
            )
        }
    model["rows"][name] = rows
    model["cells"][name] = cells
    model["formulas"][name] = formulas
    merges = [
        rect
        for rect in (
            _parse_range(m.get("ref") or "")
            for m in _kids(_kid(root, "mergeCells"), "mergeCell")
        )
        if rect
    ]
    if merges:
        _merge_styles(merges, cell_styles, row_styles, col_style, styles)
    model["cell_styles"]["sheets"][name] = cell_styles

    # Merged cells
    model["merged_cells"][name] = {_render_range(*rect): True for rect in merges}

    # Views
    views: dict[str, Any] = {}
    view = _kid(_kid(root, "sheetViews"), "sheetView")
    if view is not None:
        pane = _kid(view, "pane")
        if pane is not None and pane.get("state") in ("frozen", "frozenSplit"):
            views["freeze"] = {
                "cols": _int(pane.get("xSplit"), 0),
                "rows": _int(pane.get("ySplit"), 0),
            }
    views["right_to_left"] = _bool(_attr(view, "rightToLeft"), False)
    views["show_gridlines"] = _bool(_attr(view, "showGridLines"), True)
    views["show_headers"] = _bool(_attr(view, "showRowColHeaders"), True)
    views["zoom"] = _int(_attr(view, "zoomScale"), 100)
    sheet_pr = _kid(root, "sheetPr")
    tab = _color(_kid(sheet_pr, "tabColor"))
    if tab is not None:
        views["tab_color"] = tab
    model["views"][name] = views

    # Auto filter
    af = _kid(root, "autoFilter")
    if af is not None:
        entry = {"ref": _canon_ref(af.get("ref"))}
        cols = {fc.get("colId", ""): _dump(fc) for fc in _kids(af, "filterColumn")}
        if cols:
            entry["columns"] = cols
        model["auto_filters"][name] = entry

    ext_dv: list[Any] = []
    ext_cf: list[Any] = []
    for ext in _kids(_kid(root, "extLst"), "ext"):
        for c in ext:
            if _ln(c.tag) == "dataValidations":
                ext_dv.extend(_kids(c, "dataValidation"))
            elif _ln(c.tag) == "conditionalFormattings":
                ext_cf.extend(_kids(c, "conditionalFormatting"))

    model["data_validations"][name] = _validations(
        _kids(_kid(root, "dataValidations"), "dataValidation"), ext_dv
    )
    model["conditional_formats"][name] = _conditional_formats(
        _kids(root, "conditionalFormatting"), ext_cf, styles
    )

    # Hyperlinks
    links: dict[str, Any] = {}
    for h in _kids(_kid(root, "hyperlinks"), "hyperlink"):
        entry = {}
        rel = pkg.deref(part, _rel_attr(h, "id"))
        if rel is not None:
            entry["target"] = rel.target
        if h.get("location"):
            entry["location"] = h.get("location")
        if h.get("tooltip"):
            entry["tooltip"] = h.get("tooltip")
        links[_canon_ref(h.get("ref", "")) or ""] = entry
    model["hyperlinks"][name] = links

    # Comments
    notes: dict[str, Any] = {}
    for cpart in pkg.targets(part, "comments"):
        croot = pkg.xml(cpart)
        authors = [a.text or "" for a in _kids(_kid(croot, "authors"), "author")]
        for cm in _kids(_kid(croot, "commentList"), "comment"):
            aid = _int(cm.get("authorId"))
            notes[_canon_ref(cm.get("ref", "")) or ""] = {
                "author": authors[aid]
                if aid is not None and 0 <= aid < len(authors)
                else None,
                "text": _text_of(_kid(cm, "text")),
            }
    threads: dict[str, list[Any]] = {}
    for tpart in pkg.targets(part, "threadedComment"):
        for tc in _kids(pkg.xml(tpart), "threadedComment"):
            threads.setdefault(_canon_ref(tc.get("ref", "")) or "", []).append(
                {
                    "author": persons.get(tc.get("personId", "")),
                    "text": _unescape(_kid(tc, "text").text or "")
                    if _kid(tc, "text") is not None
                    else "",
                }
            )
    for addr in threads:
        notes.pop(addr, None)
    model["comments"][name] = {"notes": notes, "threads": threads}

    # Page setup
    ps = _kid(root, "pageSetup")
    fit = _bool(_attr(_kid(sheet_pr, "pageSetUpPr"), "fitToPage"), False)
    orientation = _attr(ps, "orientation") or "default"
    setup: dict[str, Any] = {
        "orientation": "portrait" if orientation == "default" else orientation,
        "paper_size": _int(_attr(ps, "paperSize"), 1),
        "fit_to_page": fit,
    }
    if fit:
        setup["fit_width"] = _int(_attr(ps, "fitToWidth"), 1)
        setup["fit_height"] = _int(_attr(ps, "fitToHeight"), 1)
    else:
        setup["scale"] = _int(_attr(ps, "scale"), 100)
    pm = _kid(root, "pageMargins")
    setup["margins"] = {
        side: _float(_attr(pm, side), default)
        for side, default in (
            ("left", 0.7),
            ("right", 0.7),
            ("top", 0.75),
            ("bottom", 0.75),
            ("header", 0.3),
            ("footer", 0.3),
        )
    }
    hf = _kid(root, "headerFooter")
    if hf is not None:
        hf_out: dict[str, Any] = {}
        for c in hf:
            sections = _hf_sections(_unescape(c.text or ""))
            if sections:
                hf_out[_ln(c.tag)] = sections
        for flag in ("differentOddEven", "differentFirst"):
            if _bool(hf.get(flag), False):
                hf_out[flag] = True
        if hf_out:
            setup["header_footer"] = hf_out
    if print_titles:
        setup["print_titles"] = print_titles
    if print_area:
        setup["print_area"] = print_area
    model["page_setup"][name] = setup

    # Tables
    tables: dict[str, Any] = {}
    for tpart in pkg.targets(part, "table"):
        t = pkg.xml(tpart)
        info = _kid(t, "tableStyleInfo")
        entry = {
            "ref": _canon_ref(t.get("ref")),
            "columns": [
                _unescape(tc.get("name", ""))
                for tc in _kids(_kid(t, "tableColumns"), "tableColumn")
            ],
            "header_row": (_int(t.get("headerRowCount"), 1) or 0) > 0,
            "totals_row": (_int(t.get("totalsRowCount"), 0) or 0) > 0,
        }
        if info is not None:
            entry["style"] = {
                "name": info.get("name"),
                **{
                    flag: _bool(info.get(flag), False)
                    for flag in (
                        "showFirstColumn",
                        "showLastColumn",
                        "showRowStripes",
                        "showColumnStripes",
                    )
                },
            }
        tables[t.get("displayName") or t.get("name") or tpart] = entry
    model["tables"][name] = tables

    # Pivot tables
    pivots: dict[str, Any] = {}
    for ppart in pkg.targets(part, "pivotTable"):
        p = pkg.xml(ppart)
        source: dict[str, Any] = {}
        for cpart in pkg.targets(ppart, "pivotCacheDefinition")[:1]:
            cs = _kid(pkg.xml(cpart), "cacheSource")
            source["type"] = _attr(cs, "type") or "worksheet"
            ws = _kid(cs, "worksheetSource")
            if ws is not None:
                for attr in ("sheet", "name"):
                    if ws.get(attr):
                        source[attr] = ws.get(attr)
                if ws.get("ref"):
                    source["ref"] = _canon_ref(ws.get("ref"))
        pivots[p.get("name") or ppart] = {
            "location": _canon_ref(_attr(_kid(p, "location"), "ref")),
            "source": source,
        }
    model["pivot_tables"][name] = pivots


def _merge_styles(
    merges: list[tuple[int, int, int, int]],
    cell_styles: dict[str, str],
    row_styles: dict[int, str],
    col_style: dict[int, str],
    styles: _Styles,
) -> None:
    """Replace the styles of merged cells by what a merged area shows.

    A merged area renders the anchor (top-left) cell's style, except that its
    border is the per-cell perimeter: top sides of the top row, bottom sides
    of the bottom row, left sides of the left column, right sides of the right
    column (a side list collapses to one value when uniform). Styles of
    covered cells and interior borders are not rendered and are dropped.
    """

    def effective(r: int, c: int) -> dict[str, Any]:
        key = cell_styles.get(f"{_col_name(c)}{r}")
        if key is None:
            key = row_styles.get(r) or col_style.get(c, styles.default_key)
        return styles.table[key]

    def side(r: int, c: int, name: str) -> Any:
        border = effective(r, c).get("border")
        return (
            border.get(name, {"style": "none"}) if isinstance(border, dict) else border
        )

    def collapse(values: list[Any]) -> Any:
        return values[0] if all(v == values[0] for v in values) else values

    composed: dict[str, str] = {}
    for r1, c1, r2, c2 in merges:
        anchor = effective(r1, c1)
        border = {
            "top": collapse([side(r1, c, "top") for c in range(c1, c2 + 1)]),
            "bottom": collapse([side(r2, c, "bottom") for c in range(c1, c2 + 1)]),
            "left": collapse([side(r, c1, "left") for r in range(r1, r2 + 1)]),
            "right": collapse([side(r, c2, "right") for r in range(r1, r2 + 1)]),
        }
        anchor_border = anchor.get("border")
        if isinstance(anchor_border, dict) and "diagonal" in anchor_border:
            border["diagonal"] = anchor_border["diagonal"]
        composed[f"{_col_name(c1)}{r1}"] = styles.add({**anchor, "border": border})
    short: dict[int, list[tuple[int, int]]] = {}
    tall: list[tuple[int, int, int, int]] = []
    for r1, c1, r2, c2 in merges:
        if r2 - r1 > 10000:
            tall.append((r1, c1, r2, c2))
            continue
        for r in range(r1, r2 + 1):
            short.setdefault(r, []).append((c1, c2))
    for addr in list(cell_styles):
        r, c = _parse_cell(addr) or (0, 0)
        if any(c1 <= c <= c2 for c1, c2 in short.get(r, ())) or any(
            r1 <= r <= r2 and c1 <= c <= c2 for r1, c1, r2, c2 in tall
        ):
            del cell_styles[addr]
    cell_styles.update(composed)


def _cell_value(
    c: Any, shared: list[tuple[str, Any]], cell_font: dict[str, Any], date1904: bool
) -> dict[str, Any] | None:
    t = c.get("t", "n")
    v = _kid(c, "v")
    raw = v.text if v is not None else None
    if t == "inlineStr":
        is_el = _kid(c, "is")
        if is_el is None:
            return None if raw is None else {"type": "s", "value": _unescape(raw)}
        text, runs = _string_item(is_el)
        out: dict[str, Any] = {"type": "s", "value": text}
        eff = _effective_runs(runs, cell_font)
        if eff:
            out["runs"] = eff
        return out
    if raw is None:
        return None
    if t == "s":
        idx = _int(raw)
        if idx is None or not (0 <= idx < len(shared)):
            return {"type": "invalid", "value": f"sst:{raw}"}
        text, runs = shared[idx]
        out = {"type": "s", "value": text}
        eff = _effective_runs(runs, cell_font)
        if eff:
            out["runs"] = eff
        return out
    if t == "str":
        return {"type": "s", "value": _unescape(raw)}
    if t == "b":
        return {"type": "b", "value": _bool(raw, False)}
    if t == "e":
        return {"type": "e", "value": raw.strip()}
    if t == "d":
        serial = _date_serial(raw, date1904)
        if serial is None:
            return {"type": "invalid", "value": raw}
        return {"type": "n", "value": float(serial)}
    if not raw.strip():
        return None
    number = _float(raw)
    if number is None:
        return {"type": "invalid", "value": raw}
    return {"type": "n", "value": number}


def _formula_texts(el: Any) -> list[str | None]:
    """formula / formula1 / formula2 / xm:f children, normalized, in order."""
    out: list[str | None] = []
    for c in el:
        tag = _ln(c.tag)
        if tag in ("formula", "f"):
            out.append(_normalize_formula(c.text or ""))
    return out


def _dv_formula(el: Any, which: str) -> str | None:
    f = _kid(el, which)
    if f is None:
        return None
    inner = _kid(f, "f")
    return _normalize_formula((inner.text if inner is not None else f.text) or "")


# Validation types that use `operator`; formula2 is read only by these types
# with between/notBetween.
_DV_OPERATOR_TYPES = ("whole", "decimal", "date", "time", "textLength")


def _validations(main: list[Any], ext: list[Any]) -> dict[str, Any]:
    groups: dict[str, dict[str, Any]] = {}
    for dv in [*main, *ext]:
        dv_type = dv.get("type", "none")
        operator = dv.get("operator", "between")
        params: dict[str, Any] = {
            "type": dv_type,
            "operator": operator if dv_type in _DV_OPERATOR_TYPES else None,
            "formula1": _dv_formula(dv, "formula1") if dv_type != "none" else None,
            "formula2": (
                _dv_formula(dv, "formula2")
                if dv_type in _DV_OPERATOR_TYPES
                and operator in ("between", "notBetween")
                else None
            ),
            "allow_blank": _bool(dv.get("allowBlank"), False),
            "show_dropdown": not _bool(dv.get("showDropDown"), False),
            "show_input_message": _bool(dv.get("showInputMessage"), False),
            "show_error_message": _bool(dv.get("showErrorMessage"), False),
            "error_style": dv.get("errorStyle", "stop"),
            "error_title": dv.get("errorTitle"),
            "error": dv.get("error"),
            "prompt_title": dv.get("promptTitle"),
            "prompt": dv.get("prompt"),
        }
        params = {k: v for k, v in params.items() if v is not None}
        sqref = dv.get("sqref")
        if sqref is None:
            sq = _kid(dv, "sqref")
            sqref = sq.text if sq is not None else ""
        key = _key(params)
        group = groups.setdefault(key, {**params, "_areas": []})
        group["_areas"].append(sqref or "")
    out: dict[str, Any] = {}
    for key in sorted(groups):
        group = groups[key]
        areas = group.pop("_areas")
        group["sqref"] = _canon_sqref(" ".join(areas))
        out[key] = group
    return out


_CF_ATTR_DEFAULTS = {
    "percent": False,
    "bottom": False,
    "aboveAverage": True,
    "equalAverage": False,
}
_CF_TEXT_TYPES = ("containsText", "notContainsText", "beginsWith", "endsWith")
# Attributes each rule type reads (ECMA-376 18.3.1.10); others are ignored.
_CF_TYPE_ATTRS = {
    **{t: ("text",) for t in _CF_TEXT_TYPES},
    "timePeriod": ("timePeriod",),
    "top10": ("rank", "percent", "bottom"),
    "aboveAverage": ("aboveAverage", "equalAverage", "stdDev"),
}
# Rule types whose condition is not a formula; formula children are ignored.
_CF_NO_FORMULA_TYPES = (
    "aboveAverage",
    "top10",
    "duplicateValues",
    "uniqueValues",
    "colorScale",
    "dataBar",
    "iconSet",
)


def _cf_rule(rule: Any, sqref: str, styles: _Styles) -> dict[str, Any]:
    rtype = rule.get("type")
    entry: dict[str, Any] = {"sqref": _canon_sqref(sqref), "type": rtype}
    if rule.get("operator") and rtype in ("cellIs", *_CF_TEXT_TYPES):
        entry["operator"] = rule.get("operator")
    formulas = _formula_texts(rule)
    if formulas and rtype not in _CF_NO_FORMULA_TYPES:
        entry["formulas"] = formulas
    if _bool(rule.get("stopIfTrue"), False):
        entry["stop_if_true"] = True
    attrs: dict[str, Any] = {}
    for attr in _CF_TYPE_ATTRS.get(rtype or "", ()):
        if attr in _CF_ATTR_DEFAULTS:
            val = _bool(rule.get(attr), _CF_ATTR_DEFAULTS[attr])
            if val != _CF_ATTR_DEFAULTS[attr]:
                attrs[attr] = val
        elif rule.get(attr) is not None:
            attrs[attr] = rule.get(attr)
    if attrs:
        entry["attrs"] = attrs
    dxf = styles.dxf(_int(rule.get("dxfId")))
    inline_dxf = _kid(rule, "dxf")
    if inline_dxf is not None:
        dxf = _dxf(inline_dxf)
    if dxf:
        entry["dxf"] = dxf
    params = [
        _dump(c) for c in rule if _ln(c.tag) in ("colorScale", "dataBar", "iconSet")
    ]
    if params:
        entry["params"] = params
    return entry


def _conditional_formats(
    main: list[Any], ext: list[Any], styles: _Styles
) -> dict[str, Any]:
    ext_rules: dict[str, tuple[Any, str]] = {}
    unlinked: list[tuple[Any, str]] = []
    for cf in ext:
        sq = _kid(cf, "sqref")
        sqref = sq.text if sq is not None else ""
        for rule in _kids(cf, "cfRule"):
            rid = rule.get("id")
            if rid:
                ext_rules[rid] = (rule, sqref)
            unlinked.append((rule, sqref))
    linked: set[str] = set()
    collected: list[tuple[int, int, dict[str, Any]]] = []
    order = 0
    for cf in main:
        sqref = cf.get("sqref", "")
        for rule in _kids(cf, "cfRule"):
            entry = _cf_rule(rule, sqref, styles)
            for ext_el in _kids(_kid(rule, "extLst"), "ext"):
                ident = _kid(ext_el, "id")
                if ident is not None and ident.text in ext_rules:
                    linked.add(ident.text)
                    ext_rule = ext_rules[ident.text][0]
                    entry["x14"] = [_dump(c) for c in ext_rule if _ln(c.tag) != "dxf"]
            priority = _int(rule.get("priority"))
            collected.append(
                (priority if priority is not None else 1 << 30, order, entry)
            )
            order += 1
    for rule, sqref in unlinked:
        if rule.get("id") in linked:
            continue
        entry = _cf_rule(rule, sqref, styles)
        priority = _int(rule.get("priority"))
        collected.append((priority if priority is not None else 1 << 30, order, entry))
        order += 1
    out: dict[str, Any] = {}
    for rank, (_, _, entry) in enumerate(
        sorted(collected, key=lambda x: (x[0], x[1])), 1
    ):
        entry["rank"] = rank
        out[_key(entry)] = entry
    return dict(sorted(out.items()))


# ---------------------------------------------------------------------------
# Drawings: images and charts
# ---------------------------------------------------------------------------


def _choose(alt: Any) -> list[Any]:
    """Children of the first Choice (else Fallback) of mc:AlternateContent."""
    choice = _kid(alt, "Choice")
    if choice is None:
        choice = _kid(alt, "Fallback")
    return list(choice) if choice is not None else []


def _anchors(root: Any) -> list[Any]:
    out: list[Any] = []
    for child in root:
        tag = _ln(child.tag)
        if tag == "AlternateContent":
            out.extend(_anchors(_choose_el(child)))
        elif tag in ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor"):
            out.append(child)
    return out


def _choose_el(alt: Any) -> Any:
    holder = ET.Element("holder")
    holder.extend(_choose(alt))
    return holder


def _walk_objects(el: Any) -> list[tuple[str, Any]]:
    """('pic', blip) and ('chart', chart ref element) in document order."""
    found: list[tuple[str, Any]] = []
    for c in el:
        tag = _ln(c.tag)
        if tag == "AlternateContent":
            found.extend(_walk_objects(_choose_el(c)))
        elif tag == "pic":
            blip = next((d for d in c.iter() if _ln(d.tag) == "blip"), None)
            if blip is not None:
                found.append(("pic", blip))
        elif tag == "graphicFrame":
            data = next((d for d in c.iter() if _ln(d.tag) == "graphicData"), None)
            for d in data if data is not None else ():
                if _ln(d.tag) == "chart":
                    found.append(("chart", d))
        elif tag == "grpSp":
            found.extend(_walk_objects(c))
    return found


def _anchor_name(anchor: Any) -> str:
    frm = _kid(anchor, "from")
    if frm is None:
        return "absolute"
    col = (
        _int((_kid(frm, "col").text if _kid(frm, "col") is not None else None), 0) or 0
    )
    row = (
        _int((_kid(frm, "row").text if _kid(frm, "row") is not None else None), 0) or 0
    )
    return f"{_col_name(col + 1)}{row + 1}"


def _image_token(pkg: _Package, source: str, rid: str | None) -> str | None:
    rel = pkg.deref(source, rid)
    if rel is None:
        return None
    if rel.external:
        return f"external:{rel.target}"
    data = pkg.parts.get(rel.target)
    return hashlib.sha256(data).hexdigest() if data is not None else None


def _drawings(
    pkg: _Package, part: str, root: Any, name: str, model: dict[str, Any]
) -> None:
    images: dict[str, int] = {}
    charts: dict[str, Any] = {}
    is_chartsheet = _ln(root.tag) == "chartsheet"
    picture = _kid(root, "picture")
    if picture is not None:
        token = _image_token(pkg, part, _rel_attr(picture, "id"))
        if token:
            images[f"background|{token}"] = images.get(f"background|{token}", 0) + 1
    drawing = _kid(root, "drawing")
    dparts = []
    if drawing is not None:
        rel = pkg.deref(part, _rel_attr(drawing, "id"))
        if rel and not rel.external and rel.target in pkg.parts:
            dparts.append(rel.target)
    for dpart in dparts:
        droot = pkg.xml(dpart)
        for anchor in _anchors(droot):
            where = "sheet" if is_chartsheet else _anchor_name(anchor)
            for kind, el in _walk_objects(anchor):
                if kind == "pic":
                    token = _image_token(
                        pkg, dpart, _rel_attr(el, "embed") or _rel_attr(el, "link")
                    )
                    if token:
                        key = f"{where}|{token}"
                        images[key] = images.get(key, 0) + 1
                else:
                    rel = pkg.deref(dpart, _rel_attr(el, "id"))
                    if rel is None or rel.external or rel.target not in pkg.parts:
                        continue
                    key = where
                    n = 2
                    while key in charts:
                        key = f"{where}#{n}"
                        n += 1
                    charts[key] = _chart(pkg.xml(rel.target))
    model["images"][name] = dict(sorted(images.items()))
    model["charts"][name] = charts


_STRING_FORMULA_RE = re.compile(r'^=?\s*"((?:[^"]|"")*)"$')


def _ref_or_literal(el: Any) -> Any:
    if el is None:
        return None
    for d in el.iter():
        if _ln(d.tag) == "f" and d.text:
            literal = _STRING_FORMULA_RE.match(d.text.strip())
            if literal:
                return {"literal": [literal.group(1).replace('""', '"')]}
            return {"ref": _chart_ref(d.text)}
    values = [_unescape(v.text or "") for v in el.iter() if _ln(v.tag) == "v"]
    if values:
        return {"literal": values}
    return None


def _chart(root: Any) -> dict[str, Any]:
    ns = root.tag[1:].split("}", 1)[0] if root.tag.startswith("{") else ""
    if "chartex" in ns:
        layouts = [s.get("layoutId") for s in root.iter() if _ln(s.tag) == "series"]
        return {"format": "chartex", "types": layouts}
    chart = _kid(root, "chart")
    out: dict[str, Any] = {}
    title = _kid(chart, "title")
    if title is not None:
        tx = _kid(title, "tx")
        texts = [
            t.text or ""
            for t in (tx.iter() if tx is not None else ())
            if _ln(t.tag) == "t"
        ]
        if texts:
            out["title"] = {"text": "".join(texts)}
        else:
            out["title"] = _ref_or_literal(tx) or {"auto": True}
    plot = _kid(chart, "plotArea")
    types: list[str] = []
    series: list[tuple[int, int, dict[str, Any]]] = []
    seq = 0
    for p in plot if plot is not None else ():
        tag = _ln(p.tag)
        if not tag.endswith("Chart"):
            continue
        bar_dir = _attr(_kid(p, "barDir"), "val")
        ptype = (
            f"{tag}/{bar_dir or 'col'}" if tag in ("barChart", "bar3DChart") else tag
        )
        types.append(ptype)
        for ser in _kids(p, "ser"):
            entry: dict[str, Any] = {"plot": ptype}
            for field, child in (
                ("name", "tx"),
                ("values", "val"),
                ("categories", "cat"),
                ("x", "xVal"),
                ("y", "yVal"),
                ("bubble_sizes", "bubbleSize"),
            ):
                value = _ref_or_literal(_kid(ser, child))
                if value is not None:
                    entry[field] = value
            order = _int(_attr(_kid(ser, "order"), "val"), seq)
            series.append((order if order is not None else seq, seq, entry))
            seq += 1
    out["types"] = types
    out["series"] = [e for _, _, e in sorted(series, key=lambda x: (x[0], x[1]))]
    return out


# ---------------------------------------------------------------------------
# Comparison
# ---------------------------------------------------------------------------


def _expand_styles(model: dict[str, Any]) -> dict[str, Any]:
    """Replace content-addressed style keys with the style dicts they name."""
    styles = model.get("cell_styles", {})
    table = styles.get("table", {})

    def resolve(key: Any) -> Any:
        return table.get(key, key) if isinstance(key, str) else key

    out = dict(model)
    out["cell_styles"] = {
        "default": resolve(styles.get("default")),
        "sheets": {
            sheet: {addr: resolve(k) for addr, k in cells.items()}
            for sheet, cells in styles.get("sheets", {}).items()
        },
    }
    for feature in ("columns", "rows"):
        out[feature] = {
            sheet: {
                span: (
                    {**props, "style": resolve(props["style"])}
                    if "style" in props
                    else props
                )
                for span, props in entries.items()
            }
            for sheet, entries in model.get(feature, {}).items()
        }
    return out


def _escape(key: str) -> str:
    return key.replace("~", "~0").replace("/", "~1")


# Shared presentation tolerances (see the module docstring). Each pattern is
# matched against the full JSON Pointer of a numeric leaf.
_LENGTH_TOLERANCES = (
    # column width in characters: 1 character = 7 px = 5.25 pt
    (re.compile(r"^/columns/[^/]+/[^/]+/width$"), 0.5 / 5.25),
    # row height in points
    (re.compile(r"^/rows/[^/]+/[^/]+/height$"), 0.5),
    # page and header/footer margins in inches
    (re.compile(r"^/page_setup/[^/]+/margins/[^/]+$"), 0.5 / 72),
)
_RELATIVE_FLOAT_PATHS = re.compile(r"/fill/gradient/")
_RELATIVE_TOLERANCE = 1e-9
# A color modifier on a -1..1 scale changes an 8-bit channel by at most
# 255 * delta * 2 levels; below 1/510 it cannot move a channel a full level.
_MODIFIER_TOLERANCE = 1 / 510
_TINT_RE = re.compile(r"^(.*)~tint:([^~]+)$")
_CONTENT_KEY_RE = re.compile(r"^[0-9a-f]{16}$")


def _close(x: float, y: float) -> bool:
    return abs(x - y) <= _RELATIVE_TOLERANCE * max(abs(x), abs(y))


def _split_tint(color: str) -> tuple[str, float | None]:
    """Color token -> (base, tint); an absent tint is 0, a bad one None."""
    m = _TINT_RE.match(color)
    if not m:
        return color, 0.0
    try:
        return m.group(1), float(m.group(2))
    except ValueError:
        return m.group(1), None


def _leaf_equal(x: Any, y: Any, path: str) -> bool:
    if type(x) is type(y) and x == y:
        return True
    numbers = (int, float)
    if (
        isinstance(x, numbers)
        and isinstance(y, numbers)
        and not isinstance(x, bool)
        and not isinstance(y, bool)
    ):
        for pattern, tolerance in _LENGTH_TOLERANCES:
            if pattern.match(path):
                return abs(x - y) <= tolerance
        return bool(_RELATIVE_FLOAT_PATHS.search(path)) and _close(x, y)
    if isinstance(x, str) and isinstance(y, str) and ("~tint:" in x or "~tint:" in y):
        bx, tx = _split_tint(x)
        by, ty = _split_tint(y)
        return (
            bx == by
            and tx is not None
            and ty is not None
            and abs(tx - ty) <= _MODIFIER_TOLERANCE
        )
    return False


def _pair_content_keys(
    x: dict[str, Any], y: dict[str, Any], path: str, feature: str
) -> set[str]:
    """Content-addressed keys present on one side only whose values are equal
    under the presentation tolerances (the hash itself is exact)."""
    only_x = [k for k in x if k not in y and _CONTENT_KEY_RE.match(str(k))]
    only_y = [k for k in y if k not in x and _CONTENT_KEY_RE.match(str(k))]
    paired: set[str] = set()
    for kx in only_x:
        for ky in only_y:
            if ky in paired:
                continue
            probe: list[dict[str, Any]] = []
            _diff(x[kx], y[ky], f"{path}/{kx}", feature, probe)
            if not probe:
                paired.update((kx, ky))
                break
    return paired


def _align_runs(
    x: dict[str, Any], y: dict[str, Any]
) -> tuple[dict[str, Any], dict[str, Any]]:
    """Re-key two column-run maps onto their common partition, so runs split
    differently (e.g. by widths equal within tolerance) compare per column."""

    def spans(d: dict[str, Any]) -> list[tuple[int, int, Any]]:
        out = []
        for key, props in d.items():
            m = _SPAN_RE.match(str(key))
            if m:
                out.append((_col_num(m.group(1)), _col_num(m.group(2)), props))
        return out

    sx, sy = spans(x), spans(y)
    cuts = sorted({lo for lo, _, _ in sx + sy} | {hi + 1 for _, hi, _ in sx + sy})

    def rekey(d: dict[str, Any], runs: list[tuple[int, int, Any]]) -> dict[str, Any]:
        out = {k: v for k, v in d.items() if not _SPAN_RE.match(str(k))}
        for lo, hi, props in runs:
            inside = [c for c in cuts if lo <= c <= hi + 1]
            for s, e in itertools.pairwise(inside):
                out[f"{_col_name(s)}:{_col_name(e - 1)}"] = props
        return out

    return rekey(x, sx), rekey(y, sy)


_SPAN_RE = re.compile(r"^([A-Z]{1,3}):([A-Z]{1,3})$")


def _diff(x: Any, y: Any, path: str, feature: str, out: list[dict[str, Any]]) -> None:
    if isinstance(x, dict) and isinstance(y, dict):
        paired = _pair_content_keys(x, y, path, feature)
        for key in sorted(set(x) | set(y), key=str):
            if key in paired:
                continue
            sub = f"{path}/{_escape(str(key))}"
            if key not in y:
                out.append(
                    {
                        "feature": feature,
                        "path": sub,
                        "kind": "missing",
                        "before": x[key],
                        "after": None,
                    }
                )
            elif key not in x:
                out.append(
                    {
                        "feature": feature,
                        "path": sub,
                        "kind": "added",
                        "before": None,
                        "after": y[key],
                    }
                )
            else:
                _diff(x[key], y[key], sub, feature, out)
        return
    if isinstance(x, list) and isinstance(y, list):
        for i in range(max(len(x), len(y))):
            sub = f"{path}/{i}"
            if i >= len(y):
                out.append(
                    {
                        "feature": feature,
                        "path": sub,
                        "kind": "missing",
                        "before": x[i],
                        "after": None,
                    }
                )
            elif i >= len(x):
                out.append(
                    {
                        "feature": feature,
                        "path": sub,
                        "kind": "added",
                        "before": None,
                        "after": y[i],
                    }
                )
            else:
                _diff(x[i], y[i], sub, feature, out)
        return
    if not _leaf_equal(x, y, path):
        out.append(
            {
                "feature": feature,
                "path": path,
                "kind": "changed",
                "before": x,
                "after": y,
            }
        )


def compare(before: dict[str, Any], after: dict[str, Any]) -> list[dict[str, Any]]:
    b, a = _expand_styles(before), _expand_styles(after)
    diffs: list[dict[str, Any]] = []
    for feature in FEATURES:
        path = f"/{feature}"
        if feature not in a and feature in b:
            diffs.append(
                {
                    "feature": feature,
                    "path": path,
                    "kind": "missing",
                    "before": b[feature],
                    "after": None,
                }
            )
        elif feature not in b and feature in a:
            diffs.append(
                {
                    "feature": feature,
                    "path": path,
                    "kind": "added",
                    "before": None,
                    "after": a[feature],
                }
            )
        elif feature in a:
            x, y = b[feature], a[feature]
            if feature == "columns":
                x, y = dict(x), dict(y)
                for sheet in set(x) & set(y):
                    x[sheet], y[sheet] = _align_runs(x[sheet], y[sheet])
            _diff(x, y, path, feature, diffs)
    return diffs
