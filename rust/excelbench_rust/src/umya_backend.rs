use pyo3::exceptions::{PyIOError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::PyDict;

use std::path::Path;

use chrono::{Duration, NaiveDate, NaiveDateTime, NaiveTime};

use std::str::FromStr;

use umya_spreadsheet::{new_file, reader, writer, NumberingFormat, Spreadsheet};
use umya_spreadsheet::structs::{
    EnumTrait, HorizontalAlignmentValues, PatternValues, VerticalAlignmentValues,
};

use crate::util::{a1_to_row_col, cell_blank, cell_with_value, parse_iso_date, parse_iso_datetime};

fn looks_like_date_format(code: &str) -> bool {
    // Heuristic: date formats typically include year + day tokens.
    let lc = code.to_ascii_lowercase();
    lc.contains('y') && lc.contains('d')
}

fn excel_serial_to_naive_datetime(serial: f64) -> Option<NaiveDateTime> {
    // Excel 1900 date system, with the standard 1900 leap-year bug adjustment.
    let epoch = NaiveDate::from_ymd_opt(1899, 12, 30)?.and_time(NaiveTime::MIN);
    let mut f = serial;
    if f < 60.0 {
        f += 1.0;
    }
    let total_ms = (f * 86_400_000.0).round() as i64;
    epoch.checked_add_signed(Duration::milliseconds(total_ms))
}

fn naive_datetime_to_excel_serial(dt: NaiveDateTime) -> Option<f64> {
    let epoch = NaiveDate::from_ymd_opt(1899, 12, 30)?.and_time(NaiveTime::MIN);
    let delta = dt - epoch;
    let total_ms = delta.num_milliseconds();
    Some(total_ms as f64 / 86_400_000.0)
}

// ---------------------------------------------------------------------------
// Color helpers: ARGB <-> hex
// ---------------------------------------------------------------------------

/// Convert ARGB "FFRRGGBB" or "RRGGBB" to "#RRGGBB".
fn argb_to_hex(argb: &str) -> String {
    let s = argb.trim();
    if s.len() == 8 {
        // "FFRRGGBB" → "#RRGGBB"
        format!("#{}", &s[2..])
    } else if s.len() == 6 {
        format!("#{s}")
    } else if s.starts_with('#') {
        s.to_string()
    } else {
        format!("#{s}")
    }
}

/// Convert "#RRGGBB" to "FFRRGGBB" ARGB.
fn hex_to_argb(hex: &str) -> String {
    let s = hex.strip_prefix('#').unwrap_or(hex);
    format!("FF{s}")
}

/// Map umya border style string to our canonical style names.
fn umya_border_style_to_str(style: &str) -> &'static str {
    match style.to_ascii_lowercase().as_str() {
        "thin" => "thin",
        "medium" => "medium",
        "thick" => "thick",
        "double" => "double",
        "dashed" => "dashed",
        "dotted" => "dotted",
        "hair" => "hair",
        "mediumdashed" | "mediumDashed" => "mediumDashed",
        "dashdot" | "dashDot" => "dashDot",
        "mediumdashdot" | "mediumDashDot" => "mediumDashDot",
        "dashdotdot" | "dashDotDot" => "dashDotDot",
        "mediumdashdotdot" | "mediumDashDotDot" => "mediumDashDotDot",
        "slantdashdot" | "slantDashDot" => "slantDashDot",
        _ => "none",
    }
}

fn col_letter_to_u32(col_str: &str) -> Result<u32, String> {
    let mut col: u32 = 0;
    for ch in col_str.chars() {
        if !ch.is_ascii_alphabetic() {
            return Err(format!("Invalid column letter: {col_str}"));
        }
        let uc = ch.to_ascii_uppercase() as u8;
        col = col * 26 + (uc - b'A' + 1) as u32;
    }
    if col == 0 {
        return Err(format!("Invalid column letter: {col_str}"));
    }
    Ok(col)
}

#[pyclass(unsendable)]
pub struct UmyaBook {
    book: Spreadsheet,
    saved: bool,
}

#[pymethods]
impl UmyaBook {
    #[new]
    pub fn new() -> Self {
        let mut book = new_file();
        // Match other adapters: start without a default sheet.
        let _ = book.remove_sheet_by_name("Sheet1");
        Self { book, saved: false }
    }

    #[staticmethod]
    pub fn open(path: &str) -> PyResult<Self> {
        let p = Path::new(path);
        let book = reader::xlsx::read(p)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to open workbook: {e}")))?;
        Ok(Self { book, saved: false })
    }

    pub fn sheet_names(&self) -> PyResult<Vec<String>> {
        let mut names: Vec<String> = Vec::new();
        for sheet in self.book.get_sheet_collection().iter() {
            names.push(sheet.get_name().to_string());
        }
        Ok(names)
    }

    pub fn add_sheet(&mut self, name: &str) -> PyResult<()> {
        self.book
            .new_sheet(name)
            .map_err(|e| PyErr::new::<PyValueError, _>(format!("Failed to add sheet: {e}")))?;
        Ok(())
    }

    // =========================================================================
    // Read operations
    // =========================================================================

    pub fn read_cell_value(&self, py: Python<'_>, sheet: &str, a1: &str) -> PyResult<PyObject> {
        let ws = self
            .book
            .get_sheet_by_name(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let (row0, col0) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
        let coord = (col0 + 1, row0 + 1);

        let cell = match ws.get_cell(coord) {
            Some(c) => c,
            None => return cell_blank(py),
        };

        // Formula wins over value.
        let formula = cell.get_formula();
        if !formula.is_empty() {
            // Map well-known error formulas to error tokens (similar to OpenpyxlAdapter).
            let norm = if formula.starts_with('=') {
                formula.to_string()
            } else {
                format!("={formula}")
            };
            let token = match norm.as_str() {
                "=1/0" => Some("#DIV/0!"),
                "=NA()" => Some("#N/A"),
                "=\"text\"+1" => Some("#VALUE!"),
                _ => None,
            };
            if let Some(t) = token {
                return cell_with_value(py, "error", t);
            }

            let d = PyDict::new_bound(py);
            d.set_item("type", "formula")?;
            d.set_item("formula", formula.to_string())?;
            d.set_item("value", formula.to_string())?;
            return Ok(d.into());
        }

        // Numeric typed access.
        if let Some(f) = cell.get_value_number() {
            if let Some(nf) = cell.get_style().get_number_format() {
                let code = nf.get_format_code();
                if looks_like_date_format(code) {
                    if let Some(ndt) = excel_serial_to_naive_datetime(f) {
                        let midnight = NaiveTime::from_hms_opt(0, 0, 0).unwrap();
                        if ndt.time() == midnight {
                            let s = ndt.date().format("%Y-%m-%d").to_string();
                            return cell_with_value(py, "date", s);
                        }
                        let s = ndt.format("%Y-%m-%dT%H:%M:%S").to_string();
                        return cell_with_value(py, "datetime", s);
                    }
                }
            }

            return cell_with_value(py, "number", f);
        }

        // TODO: improve typing (beyond Tier 1). For now treat as error/bool/string.
        let raw = cell
            .get_value()
            .into_owned()
            .replace("\r\n", "\n")
            .replace('\r', "\n");

        // Errors
        if raw == "#N/A" || (raw.starts_with('#') && raw.ends_with('!')) {
            return cell_with_value(py, "error", raw);
        }
        // Boolean
        if raw.eq_ignore_ascii_case("true") {
            return cell_with_value(py, "boolean", true);
        }
        if raw.eq_ignore_ascii_case("false") {
            return cell_with_value(py, "boolean", false);
        }

        if raw.is_empty() {
            return cell_blank(py);
        }

        cell_with_value(py, "string", raw)
    }

    pub fn read_cell_format(
        &self,
        py: Python<'_>,
        sheet: &str,
        a1: &str,
    ) -> PyResult<PyObject> {
        let ws = self
            .book
            .get_sheet_by_name(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let (row0, col0) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
        let coord = (col0 + 1, row0 + 1);

        let d = PyDict::new_bound(py);

        let cell = match ws.get_cell(coord) {
            Some(c) => c,
            None => return Ok(d.into()),
        };

        let style = cell.get_style();

        // Font properties
        if let Some(font) = style.get_font() {
            if *font.get_bold() {
                d.set_item("bold", true)?;
            }
            if *font.get_italic() {
                d.set_item("italic", true)?;
            }
            {
                let ul = font.get_underline();
                if !ul.is_empty() && ul != "none" {
                    d.set_item("underline", ul.to_string())?;
                }
            }
            if *font.get_strikethrough() {
                d.set_item("strikethrough", true)?;
            }
            {
                let name = font.get_name();
                if !name.is_empty() {
                    d.set_item("font_name", name.to_string())?;
                }
            }
            {
                let size = *font.get_size();
                if size > 0.0 {
                    d.set_item("font_size", size)?;
                }
            }
            {
                let argb = font.get_color().get_argb();
                if !argb.is_empty() {
                    let hex = argb_to_hex(argb);
                    if hex != "#000000" {
                        d.set_item("font_color", hex)?;
                    }
                }
            }
        }

        // Fill / background color
        if let Some(fill) = style.get_fill() {
            if let Some(pf) = fill.get_pattern_fill() {
                if let Some(fg) = pf.get_foreground_color() {
                    let argb = fg.get_argb();
                    if !argb.is_empty() {
                        let hex = argb_to_hex(argb);
                        d.set_item("bg_color", hex)?;
                    }
                }
            }
        }

        // Number format
        if let Some(nf) = style.get_number_format() {
            let code = nf.get_format_code();
            if !code.is_empty() && code != "General" {
                d.set_item("number_format", code.to_string())?;
            }
        }

        // Alignment
        if let Some(align) = style.get_alignment() {
            let h = align.get_horizontal().get_value_string();
            if !h.is_empty() && h != "general" {
                d.set_item("h_align", h.to_string())?;
            }
            let v = align.get_vertical().get_value_string();
            if !v.is_empty() && v != "bottom" {
                d.set_item("v_align", v.to_string())?;
            }
            if *align.get_wrap_text() {
                d.set_item("wrap", true)?;
            }
            let rot = *align.get_text_rotation();
            if rot != 0 {
                d.set_item("rotation", rot)?;
            }
        }

        Ok(d.into())
    }

    pub fn read_cell_border(
        &self,
        py: Python<'_>,
        sheet: &str,
        a1: &str,
    ) -> PyResult<PyObject> {
        let ws = self
            .book
            .get_sheet_by_name(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let (row0, col0) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
        let coord = (col0 + 1, row0 + 1);

        let d = PyDict::new_bound(py);

        let cell = match ws.get_cell(coord) {
            Some(c) => c,
            None => return Ok(d.into()),
        };

        let style = cell.get_style();
        if let Some(borders) = style.get_borders() {
            // Helper closure to read one edge. Returns None if style is "none" or empty.
            let read_edge = |e: &umya_spreadsheet::structs::Border| -> Option<(String, String)> {
                let style_str = e.get_border_style();
                if style_str.is_empty() || style_str == "none" {
                    return None;
                }
                let argb = e.get_color().get_argb();
                let color_str = if argb.is_empty() {
                    "#000000".to_string()
                } else {
                    argb_to_hex(argb)
                };
                Some((
                    umya_border_style_to_str(style_str).to_string(),
                    color_str,
                ))
            };

            if let Some((s, c)) = read_edge(borders.get_top()) {
                let edge = PyDict::new_bound(py);
                edge.set_item("style", s)?;
                edge.set_item("color", c)?;
                d.set_item("top", edge)?;
            }
            if let Some((s, c)) = read_edge(borders.get_bottom()) {
                let edge = PyDict::new_bound(py);
                edge.set_item("style", s)?;
                edge.set_item("color", c)?;
                d.set_item("bottom", edge)?;
            }
            if let Some((s, c)) = read_edge(borders.get_left()) {
                let edge = PyDict::new_bound(py);
                edge.set_item("style", s)?;
                edge.set_item("color", c)?;
                d.set_item("left", edge)?;
            }
            if let Some((s, c)) = read_edge(borders.get_right()) {
                let edge = PyDict::new_bound(py);
                edge.set_item("style", s)?;
                edge.set_item("color", c)?;
                d.set_item("right", edge)?;
            }
            if let Some((s, c)) = read_edge(borders.get_diagonal()) {
                let edge = PyDict::new_bound(py);
                edge.set_item("style", s)?;
                edge.set_item("color", c)?;
                // umya doesn't distinguish up/down diagonal — use diagonal_up.
                d.set_item("diagonal_up", edge)?;
            }
        }

        Ok(d.into())
    }

    pub fn read_row_height(&self, sheet: &str, row: u32) -> PyResult<Option<f64>> {
        let ws = self
            .book
            .get_sheet_by_name(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        // umya uses 1-based row index.
        if let Some(rd) = ws.get_row_dimension(&(row + 1)) {
            let h = rd.get_height();
            if h > &0.0 {
                return Ok(Some(*h));
            }
        }
        Ok(None)
    }

    pub fn read_column_width(&self, sheet: &str, col_str: &str) -> PyResult<Option<f64>> {
        let ws = self
            .book
            .get_sheet_by_name(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let col_idx =
            col_letter_to_u32(col_str).map_err(|e| PyErr::new::<PyValueError, _>(e))?;

        if let Some(cd) = ws.get_column_dimension_by_number(&col_idx) {
            let w = cd.get_width();
            if w > &0.0 {
                return Ok(Some(*w));
            }
        }
        Ok(None)
    }

    // =========================================================================
    // Write operations
    // =========================================================================

    pub fn write_cell_value(
        &mut self,
        sheet: &str,
        a1: &str,
        payload: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let dict = payload
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("payload must be a dict"))?;
        let type_obj = dict
            .get_item("type")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("payload missing 'type'"))?;
        let type_str: String = type_obj.extract()?;

        match type_str.as_str() {
            "blank" => {
                // no-op
                Ok(())
            }
            "string" => {
                let v = dict.get_item("value")?;
                let s = match v {
                    Some(v) => v.extract::<String>()?,
                    None => String::new(),
                };
                ws.get_cell_mut(a1).set_value_string(s);
                Ok(())
            }
            "number" => {
                let v = dict.get_item("value")?.ok_or_else(|| {
                    PyErr::new::<PyValueError, _>("number payload missing 'value'")
                })?;
                let f = v.extract::<f64>()?;
                ws.get_cell_mut(a1).set_value_number(f);
                Ok(())
            }
            "boolean" => {
                let v = dict.get_item("value")?.ok_or_else(|| {
                    PyErr::new::<PyValueError, _>("boolean payload missing 'value'")
                })?;
                let b = v.extract::<bool>()?;
                ws.get_cell_mut(a1).set_value_bool(b);
                Ok(())
            }
            "formula" => {
                let v = if let Some(v) = dict.get_item("formula")? {
                    v
                } else if let Some(v) = dict.get_item("value")? {
                    v
                } else {
                    return Err(PyErr::new::<PyValueError, _>(
                        "formula payload missing 'formula'",
                    ));
                };
                let formula = v.extract::<String>()?;
                let f = formula.strip_prefix('=').unwrap_or(&formula);
                ws.get_cell_mut(a1).set_formula(f);
                Ok(())
            }
            "error" => {
                let v = dict.get_item("value")?.ok_or_else(|| {
                    PyErr::new::<PyValueError, _>("error payload missing 'value'")
                })?;
                let token = v.extract::<String>()?;
                // Prefer formulas for errors that OpenpyxlAdapter can recognize.
                // For other error tokens, write the literal token as a string.
                let formula = match token.as_str() {
                    "#DIV/0!" => Some("1/0"),
                    "#N/A" => Some("NA()"),
                    "#VALUE!" => Some("\"text\"+1"),
                    _ => None,
                };
                if let Some(f) = formula {
                    ws.get_cell_mut(a1).set_formula(f);
                } else {
                    ws.get_cell_mut(a1).set_value_string(token);
                }
                Ok(())
            }
            "date" => {
                let v = dict
                    .get_item("value")?
                    .ok_or_else(|| PyErr::new::<PyValueError, _>("date payload missing 'value'"))?;
                let s = v.extract::<String>()?;
                let d = parse_iso_date(&s)
                    .ok_or_else(|| PyErr::new::<PyValueError, _>("Invalid ISO date"))?;
                let dt = d.and_time(NaiveTime::from_hms_opt(0, 0, 0).unwrap());
                let serial = naive_datetime_to_excel_serial(dt)
                    .ok_or_else(|| PyErr::new::<PyValueError, _>("Failed to convert date"))?;

                ws.get_cell_mut(a1).set_value_number(serial);
                ws.get_style_mut(a1)
                    .get_number_format_mut()
                    .set_format_code(NumberingFormat::FORMAT_DATE_YYYYMMDD);
                Ok(())
            }
            "datetime" => {
                let v = dict.get_item("value")?.ok_or_else(|| {
                    PyErr::new::<PyValueError, _>("datetime payload missing 'value'")
                })?;
                let s = v.extract::<String>()?;
                let dt = parse_iso_datetime(&s)
                    .ok_or_else(|| PyErr::new::<PyValueError, _>("Invalid ISO datetime"))?;
                let serial = naive_datetime_to_excel_serial(dt)
                    .ok_or_else(|| PyErr::new::<PyValueError, _>("Failed to convert datetime"))?;

                ws.get_cell_mut(a1).set_value_number(serial);
                ws.get_style_mut(a1)
                    .get_number_format_mut()
                    .set_format_code("yyyy-mm-dd h:mm:ss");
                Ok(())
            }
            other => Err(PyErr::new::<PyValueError, _>(format!(
                "Unsupported cell type: {other}"
            ))),
        }
    }

    pub fn write_cell_format(
        &mut self,
        sheet: &str,
        a1: &str,
        format_dict: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let dict = format_dict
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("format_dict must be a dict"))?;

        let style = ws.get_style_mut(a1);

        // Font properties
        {
            let font = style.get_font_mut();

            if let Some(bold) = dict.get_item("bold")?.and_then(|v| v.extract::<bool>().ok()) {
                font.set_bold(bold);
            }
            if let Some(italic) = dict.get_item("italic")?.and_then(|v| v.extract::<bool>().ok()) {
                font.set_italic(italic);
            }
            if let Some(ul) = dict
                .get_item("underline")?
                .and_then(|v| v.extract::<String>().ok())
            {
                font.set_underline(ul);
            }
            if let Some(st) = dict
                .get_item("strikethrough")?
                .and_then(|v| v.extract::<bool>().ok())
            {
                font.set_strikethrough(st);
            }
            if let Some(name) = dict
                .get_item("font_name")?
                .and_then(|v| v.extract::<String>().ok())
            {
                font.set_name(name);
            }
            if let Some(size) = dict
                .get_item("font_size")?
                .and_then(|v| v.extract::<f64>().ok())
            {
                font.set_size(size);
            }
            if let Some(color) = dict
                .get_item("font_color")?
                .and_then(|v| v.extract::<String>().ok())
            {
                font.get_color_mut().set_argb(hex_to_argb(&color));
            }
        }

        // Background color via pattern fill
        if let Some(bg) = dict
            .get_item("bg_color")?
            .and_then(|v| v.extract::<String>().ok())
        {
            let fill = style.get_fill_mut();
            let pf = fill.get_pattern_fill_mut();
            pf.set_pattern_type(PatternValues::Solid);
            pf.get_foreground_color_mut().set_argb(hex_to_argb(&bg));
        }

        // Number format
        if let Some(nf) = dict
            .get_item("number_format")?
            .and_then(|v| v.extract::<String>().ok())
        {
            style.get_number_format_mut().set_format_code(nf);
        }

        // Alignment
        if let Some(h) = dict
            .get_item("h_align")?
            .and_then(|v| v.extract::<String>().ok())
        {
            if let Ok(ha) = HorizontalAlignmentValues::from_str(&h) {
                style.get_alignment_mut().set_horizontal(ha);
            }
        }
        if let Some(v) = dict
            .get_item("v_align")?
            .and_then(|v| v.extract::<String>().ok())
        {
            if let Ok(va) = VerticalAlignmentValues::from_str(&v) {
                style.get_alignment_mut().set_vertical(va);
            }
        }
        if let Some(wrap) = dict.get_item("wrap")?.and_then(|v| v.extract::<bool>().ok()) {
            style.get_alignment_mut().set_wrap_text(wrap);
        }
        if let Some(rot) = dict
            .get_item("rotation")?
            .and_then(|v| v.extract::<i64>().ok())
        {
            style.get_alignment_mut().set_text_rotation(rot as u32);
        }

        Ok(())
    }

    pub fn write_cell_border(
        &mut self,
        sheet: &str,
        a1: &str,
        border_dict: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let dict = border_dict
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("border_dict must be a dict"))?;

        let style = ws.get_style_mut(a1);
        let borders = style.get_borders_mut();

        // Helper to apply one edge.
        fn apply_edge(
            edge: &mut umya_spreadsheet::structs::Border,
            sub: &Bound<'_, PyDict>,
        ) -> PyResult<()> {
            if let Some(s) = sub
                .get_item("style")?
                .and_then(|v| v.extract::<String>().ok())
            {
                edge.set_border_style(s);
            }
            if let Some(c) = sub
                .get_item("color")?
                .and_then(|v| v.extract::<String>().ok())
            {
                edge.get_color_mut().set_argb(hex_to_argb(&c));
            }
            Ok(())
        }

        if let Some(sub) = dict.get_item("top")? {
            if let Ok(d) = sub.downcast::<PyDict>() {
                apply_edge(borders.get_top_mut(), d)?;
            }
        }
        if let Some(sub) = dict.get_item("bottom")? {
            if let Ok(d) = sub.downcast::<PyDict>() {
                apply_edge(borders.get_bottom_mut(), d)?;
            }
        }
        if let Some(sub) = dict.get_item("left")? {
            if let Ok(d) = sub.downcast::<PyDict>() {
                apply_edge(borders.get_left_mut(), d)?;
            }
        }
        if let Some(sub) = dict.get_item("right")? {
            if let Ok(d) = sub.downcast::<PyDict>() {
                apply_edge(borders.get_right_mut(), d)?;
            }
        }
        if let Some(sub) = dict.get_item("diagonal_up")? {
            if let Ok(d) = sub.downcast::<PyDict>() {
                apply_edge(borders.get_diagonal_mut(), d)?;
            }
        }
        if let Some(sub) = dict.get_item("diagonal_down")? {
            if let Ok(d) = sub.downcast::<PyDict>() {
                apply_edge(borders.get_diagonal_mut(), d)?;
            }
        }

        Ok(())
    }

    pub fn set_row_height(&mut self, sheet: &str, row: u32, height: f64) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        // umya uses 1-based row index.
        ws.get_row_dimension_mut(&(row + 1)).set_height(height);
        Ok(())
    }

    pub fn set_column_width(&mut self, sheet: &str, col_str: &str, width: f64) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let col_idx =
            col_letter_to_u32(col_str).map_err(|e| PyErr::new::<PyValueError, _>(e))?;

        ws.get_column_dimension_by_number_mut(&col_idx)
            .set_width(width);
        Ok(())
    }

    pub fn save(&mut self, path: &str) -> PyResult<()> {
        if self.saved {
            return Err(PyErr::new::<PyValueError, _>(
                "Workbook already saved (UmyaBook is consumed-on-save)",
            ));
        }
        self.saved = true;

        let p = Path::new(path);
        writer::xlsx::write(&self.book, p)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to save workbook: {e}")))
    }
}
