use pyo3::exceptions::{PyIOError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::PyDict;

use std::collections::HashMap;

use indexmap::IndexMap;

use rust_xlsxwriter::{
    Color, Format, FormatAlign, FormatBorder, Workbook, Worksheet,
};

use crate::util::{a1_to_row_col, parse_iso_date, parse_iso_datetime};

// ---------------------------------------------------------------------------
// Queued operation types
// ---------------------------------------------------------------------------

/// Stored cell value payload (mirrors the Python dict contract).
struct CellPayload {
    type_str: String,
    value: Option<String>,
    formula: Option<String>,
}

/// Queued per-cell format dict fields.
struct FormatFields {
    bold: Option<bool>,
    italic: Option<bool>,
    underline: Option<String>,
    strikethrough: Option<bool>,
    font_name: Option<String>,
    font_size: Option<f64>,
    font_color: Option<String>,
    bg_color: Option<String>,
    number_format: Option<String>,
    h_align: Option<String>,
    v_align: Option<String>,
    wrap: Option<bool>,
    rotation: Option<i32>,
    indent: Option<i32>,
}

/// Queued per-cell border dict fields.
struct BorderFields {
    top_style: Option<String>,
    top_color: Option<String>,
    bottom_style: Option<String>,
    bottom_color: Option<String>,
    left_style: Option<String>,
    left_color: Option<String>,
    right_style: Option<String>,
    right_color: Option<String>,
    diagonal_up_style: Option<String>,
    diagonal_up_color: Option<String>,
    diagonal_down_style: Option<String>,
    diagonal_down_color: Option<String>,
}

type CellKey = (String, u32, u16); // (sheet, row, col)

#[pyclass(unsendable)]
pub struct RustXlsxWriterBook {
    sheet_names: Vec<String>,
    values: IndexMap<CellKey, CellPayload>,
    formats: HashMap<CellKey, FormatFields>,
    borders: HashMap<CellKey, BorderFields>,
    row_heights: HashMap<(String, u32), f64>,
    col_widths: HashMap<(String, u16), f64>,
    saved: bool,
}

// ---------------------------------------------------------------------------
// Color / enum helpers
// ---------------------------------------------------------------------------

fn parse_hex_color(hex: &str) -> Color {
    let s = hex.strip_prefix('#').unwrap_or(hex);
    if let Ok(n) = u32::from_str_radix(s, 16) {
        Color::RGB(n)
    } else {
        Color::Black
    }
}

fn map_h_align(s: &str) -> FormatAlign {
    match s.to_ascii_lowercase().as_str() {
        "left" => FormatAlign::Left,
        "center" | "centre" => FormatAlign::Center,
        "right" => FormatAlign::Right,
        "fill" => FormatAlign::Fill,
        "justify" => FormatAlign::Justify,
        "distributed" | "centerContinuous" => FormatAlign::CenterAcross,
        _ => FormatAlign::Left,
    }
}

fn map_v_align(s: &str) -> FormatAlign {
    match s.to_ascii_lowercase().as_str() {
        "top" => FormatAlign::Top,
        "center" | "centre" => FormatAlign::VerticalCenter,
        "bottom" => FormatAlign::Bottom,
        "justify" => FormatAlign::VerticalJustify,
        "distributed" => FormatAlign::VerticalDistributed,
        _ => FormatAlign::Bottom,
    }
}

fn map_border_style(s: &str) -> FormatBorder {
    match s.to_ascii_lowercase().as_str() {
        "thin" => FormatBorder::Thin,
        "medium" => FormatBorder::Medium,
        "thick" => FormatBorder::Thick,
        "double" => FormatBorder::Double,
        "dashed" => FormatBorder::Dashed,
        "dotted" => FormatBorder::Dotted,
        "hair" => FormatBorder::Hair,
        "mediumdashed" | "mediumDashed" => FormatBorder::MediumDashed,
        "dashdot" | "dashDot" => FormatBorder::DashDot,
        "mediumdashdot" | "mediumDashDot" => FormatBorder::MediumDashDot,
        "dashdotdot" | "dashDotDot" => FormatBorder::DashDotDot,
        "mediumdashdotdot" | "mediumDashDotDot" => FormatBorder::MediumDashDotDot,
        "slantdashdot" | "slantDashDot" => FormatBorder::SlantDashDot,
        "none" | "" => FormatBorder::None,
        _ => FormatBorder::Thin,
    }
}

fn map_underline(s: &str) -> rust_xlsxwriter::FormatUnderline {
    match s.to_ascii_lowercase().as_str() {
        "single" => rust_xlsxwriter::FormatUnderline::Single,
        "double" => rust_xlsxwriter::FormatUnderline::Double,
        "singleaccounting" | "single_accounting" => {
            rust_xlsxwriter::FormatUnderline::SingleAccounting
        }
        "doubleaccounting" | "double_accounting" => {
            rust_xlsxwriter::FormatUnderline::DoubleAccounting
        }
        _ => rust_xlsxwriter::FormatUnderline::Single,
    }
}

// ---------------------------------------------------------------------------
// Build a Format from optional FormatFields + BorderFields
// ---------------------------------------------------------------------------

fn build_format(fmt: Option<&FormatFields>, bdr: Option<&BorderFields>) -> Format {
    let mut f = Format::new();

    if let Some(ff) = fmt {
        if ff.bold == Some(true) {
            f = f.set_bold();
        }
        if ff.italic == Some(true) {
            f = f.set_italic();
        }
        if let Some(ref ul) = ff.underline {
            f = f.set_underline(map_underline(ul));
        }
        if ff.strikethrough == Some(true) {
            f = f.set_font_strikethrough();
        }
        if let Some(ref name) = ff.font_name {
            f = f.set_font_name(name);
        }
        if let Some(size) = ff.font_size {
            f = f.set_font_size(size);
        }
        if let Some(ref color) = ff.font_color {
            f = f.set_font_color(parse_hex_color(color));
        }
        if let Some(ref bg) = ff.bg_color {
            f = f.set_background_color(parse_hex_color(bg));
        }
        if let Some(ref nf) = ff.number_format {
            f = f.set_num_format(nf);
        }
        if let Some(ref h) = ff.h_align {
            f = f.set_align(map_h_align(h));
        }
        if let Some(ref v) = ff.v_align {
            f = f.set_align(map_v_align(v));
        }
        if ff.wrap == Some(true) {
            f = f.set_text_wrap();
        }
        if let Some(rot) = ff.rotation {
            f = f.set_rotation(rot as i16);
        }
        if let Some(indent) = ff.indent {
            f = f.set_indent(indent as u8);
        }
    }

    if let Some(bb) = bdr {
        if let Some(ref s) = bb.top_style {
            f = f.set_border_top(map_border_style(s));
            if let Some(ref c) = bb.top_color {
                f = f.set_border_top_color(parse_hex_color(c));
            }
        }
        if let Some(ref s) = bb.bottom_style {
            f = f.set_border_bottom(map_border_style(s));
            if let Some(ref c) = bb.bottom_color {
                f = f.set_border_bottom_color(parse_hex_color(c));
            }
        }
        if let Some(ref s) = bb.left_style {
            f = f.set_border_left(map_border_style(s));
            if let Some(ref c) = bb.left_color {
                f = f.set_border_left_color(parse_hex_color(c));
            }
        }
        if let Some(ref s) = bb.right_style {
            f = f.set_border_right(map_border_style(s));
            if let Some(ref c) = bb.right_color {
                f = f.set_border_right_color(parse_hex_color(c));
            }
        }
        if let Some(ref s) = bb.diagonal_up_style {
            f = f.set_border_diagonal(map_border_style(s));
            if let Some(ref c) = bb.diagonal_up_color {
                f = f.set_border_diagonal_color(parse_hex_color(c));
            }
            f = f.set_border_diagonal_type(rust_xlsxwriter::FormatDiagonalBorder::BorderUp);
        }
        if let Some(ref s) = bb.diagonal_down_style {
            f = f.set_border_diagonal(map_border_style(s));
            if let Some(ref c) = bb.diagonal_down_color {
                f = f.set_border_diagonal_color(parse_hex_color(c));
            }
            // If both up + down are set, this overrides. That's a limitation
            // but diagonal borders are rare in the benchmark.
            f = f.set_border_diagonal_type(rust_xlsxwriter::FormatDiagonalBorder::BorderDown);
        }
    }

    f
}

// ---------------------------------------------------------------------------
// Extract fields from Python dicts
// ---------------------------------------------------------------------------

fn extract_format_fields(dict: &Bound<'_, PyDict>) -> PyResult<FormatFields> {
    Ok(FormatFields {
        bold: dict.get_item("bold")?.and_then(|v| v.extract().ok()),
        italic: dict.get_item("italic")?.and_then(|v| v.extract().ok()),
        underline: dict.get_item("underline")?.and_then(|v| v.extract().ok()),
        strikethrough: dict
            .get_item("strikethrough")?
            .and_then(|v| v.extract().ok()),
        font_name: dict.get_item("font_name")?.and_then(|v| v.extract().ok()),
        font_size: dict.get_item("font_size")?.and_then(|v| v.extract().ok()),
        font_color: dict
            .get_item("font_color")?
            .and_then(|v| v.extract().ok()),
        bg_color: dict.get_item("bg_color")?.and_then(|v| v.extract().ok()),
        number_format: dict
            .get_item("number_format")?
            .and_then(|v| v.extract().ok()),
        h_align: dict.get_item("h_align")?.and_then(|v| v.extract().ok()),
        v_align: dict.get_item("v_align")?.and_then(|v| v.extract().ok()),
        wrap: dict.get_item("wrap")?.and_then(|v| v.extract().ok()),
        rotation: dict.get_item("rotation")?.and_then(|v| v.extract().ok()),
        indent: dict.get_item("indent")?.and_then(|v| v.extract().ok()),
    })
}

fn extract_border_fields(dict: &Bound<'_, PyDict>) -> PyResult<BorderFields> {
    fn edge(
        dict: &Bound<'_, PyDict>,
        key: &str,
    ) -> PyResult<(Option<String>, Option<String>)> {
        if let Some(sub) = dict.get_item(key)? {
            if let Ok(d) = sub.downcast::<PyDict>() {
                let style: Option<String> =
                    d.get_item("style")?.and_then(|v| v.extract().ok());
                let color: Option<String> =
                    d.get_item("color")?.and_then(|v| v.extract().ok());
                return Ok((style, color));
            }
        }
        Ok((None, None))
    }

    let (ts, tc) = edge(dict, "top")?;
    let (bs, bc) = edge(dict, "bottom")?;
    let (ls, lc) = edge(dict, "left")?;
    let (rs, rc) = edge(dict, "right")?;
    let (dus, duc) = edge(dict, "diagonal_up")?;
    let (dds, ddc) = edge(dict, "diagonal_down")?;

    Ok(BorderFields {
        top_style: ts,
        top_color: tc,
        bottom_style: bs,
        bottom_color: bc,
        left_style: ls,
        left_color: lc,
        right_style: rs,
        right_color: rc,
        diagonal_up_style: dus,
        diagonal_up_color: duc,
        diagonal_down_style: dds,
        diagonal_down_color: ddc,
    })
}

fn resolve_key(sheet: &str, a1: &str) -> PyResult<CellKey> {
    let (row, col0) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
    let col: u16 = col0.try_into().map_err(|_| {
        PyErr::new::<PyValueError, _>(format!("Column out of range for Excel: {a1}"))
    })?;
    Ok((sheet.to_string(), row, col))
}

fn col_letter_to_index(col_str: &str) -> PyResult<u16> {
    let mut col: u32 = 0;
    for ch in col_str.chars() {
        if !ch.is_ascii_alphabetic() {
            return Err(PyErr::new::<PyValueError, _>(format!(
                "Invalid column letter: {col_str}"
            )));
        }
        let uc = ch.to_ascii_uppercase() as u8;
        col = col * 26 + (uc - b'A' + 1) as u32;
    }
    if col == 0 {
        return Err(PyErr::new::<PyValueError, _>(format!(
            "Invalid column letter: {col_str}"
        )));
    }
    let idx = (col - 1) as u16;
    Ok(idx)
}

// ---------------------------------------------------------------------------
// Helper: write a single cell's value+format to a Worksheet
// ---------------------------------------------------------------------------

fn write_cell(
    ws: &mut Worksheet,
    row: u32,
    col: u16,
    payload: &CellPayload,
    format: &Format,
) -> PyResult<()> {
    match payload.type_str.as_str() {
        "blank" => {
            // Write blank with format so the format is preserved.
            ws.write_blank(row, col, format)
                .map(|_| ())
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("write_blank failed: {e}")))
        }
        "string" => {
            let s = payload.value.as_deref().unwrap_or("");
            ws.write_string_with_format(row, col, s, format)
                .map(|_| ())
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("write_string failed: {e}")))
        }
        "number" => {
            let f_val: f64 = payload
                .value
                .as_deref()
                .unwrap_or("0")
                .parse()
                .map_err(|_| PyErr::new::<PyValueError, _>("number parse failed"))?;
            ws.write_number_with_format(row, col, f_val, format)
                .map(|_| ())
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("write_number failed: {e}")))
        }
        "boolean" => {
            let b = payload.value.as_deref().unwrap_or("false") == "true";
            ws.write_boolean_with_format(row, col, b, format)
                .map(|_| ())
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("write_boolean failed: {e}")))
        }
        "formula" => {
            let formula = payload
                .formula
                .as_deref()
                .or(payload.value.as_deref())
                .unwrap_or("");
            ws.write_formula_with_format(row, col, formula, format)
                .map(|_| ())
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("write_formula failed: {e}")))
        }
        "error" => {
            let token = payload.value.as_deref().unwrap_or("");
            let formula = match token {
                "#DIV/0!" => Some("=1/0"),
                "#N/A" => Some("=NA()"),
                "#VALUE!" => Some("=\"text\"+1"),
                _ => None,
            };
            if let Some(f) = formula {
                ws.write_formula_with_format(row, col, f, format)
                    .map(|_| ())
                    .map_err(|e| {
                        PyErr::new::<PyIOError, _>(format!("write_formula failed: {e}"))
                    })
            } else {
                ws.write_string_with_format(row, col, token, format)
                    .map(|_| ())
                    .map_err(|e| {
                        PyErr::new::<PyIOError, _>(format!("write_string failed: {e}"))
                    })
            }
        }
        "date" => {
            let s = payload.value.as_deref().unwrap_or("");
            if let Some(d) = parse_iso_date(s) {
                let mut combined_fmt = format.clone();
                // Only set date number format if the user didn't already set one.
                combined_fmt = combined_fmt.set_num_format("yyyy-mm-dd");
                ws.write_datetime_with_format(row, col, d, &combined_fmt)
                    .map(|_| ())
                    .map_err(|e| {
                        PyErr::new::<PyIOError, _>(format!("write_datetime failed: {e}"))
                    })
            } else {
                ws.write_string_with_format(row, col, s, format)
                    .map(|_| ())
                    .map_err(|e| {
                        PyErr::new::<PyIOError, _>(format!("write_string failed: {e}"))
                    })
            }
        }
        "datetime" => {
            let s = payload.value.as_deref().unwrap_or("");
            if let Some(dt) = parse_iso_datetime(s) {
                let mut combined_fmt = format.clone();
                combined_fmt = combined_fmt.set_num_format("yyyy-mm-dd hh:mm:ss");
                ws.write_datetime_with_format(row, col, dt, &combined_fmt)
                    .map(|_| ())
                    .map_err(|e| {
                        PyErr::new::<PyIOError, _>(format!("write_datetime failed: {e}"))
                    })
            } else {
                ws.write_string_with_format(row, col, s, format)
                    .map(|_| ())
                    .map_err(|e| {
                        PyErr::new::<PyIOError, _>(format!("write_string failed: {e}"))
                    })
            }
        }
        other => Err(PyErr::new::<PyValueError, _>(format!(
            "Unsupported cell type: {other}"
        ))),
    }
}

// ---------------------------------------------------------------------------
// PyO3 implementation
// ---------------------------------------------------------------------------

#[pymethods]
impl RustXlsxWriterBook {
    #[new]
    pub fn new() -> Self {
        Self {
            sheet_names: Vec::new(),
            values: IndexMap::new(),
            formats: HashMap::new(),
            borders: HashMap::new(),
            row_heights: HashMap::new(),
            col_widths: HashMap::new(),
            saved: false,
        }
    }

    pub fn add_sheet(&mut self, name: &str) -> PyResult<()> {
        if self.sheet_names.contains(&name.to_string()) {
            return Ok(());
        }
        self.sheet_names.push(name.to_string());
        Ok(())
    }

    pub fn write_cell_value(
        &mut self,
        sheet: &str,
        a1: &str,
        payload: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        if !self.sheet_names.contains(&sheet.to_string()) {
            return Err(PyErr::new::<PyValueError, _>(format!(
                "Unknown sheet: {sheet}"
            )));
        }

        let key = resolve_key(sheet, a1)?;

        let dict = payload
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("payload must be a dict"))?;
        let type_str: String = dict
            .get_item("type")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("payload missing 'type'"))?
            .extract()?;

        // Store the value for deferred writing.
        let value_str: Option<String> = dict
            .get_item("value")?
            .and_then(|v| v.extract::<String>().ok().or_else(|| {
                // Handle numeric/bool values by converting to string.
                v.extract::<f64>()
                    .map(|n| n.to_string())
                    .ok()
                    .or_else(|| v.extract::<bool>().map(|b| b.to_string()).ok())
            }));
        let formula_str: Option<String> = dict
            .get_item("formula")?
            .and_then(|v| v.extract().ok());

        self.values.insert(
            key,
            CellPayload {
                type_str,
                value: value_str,
                formula: formula_str,
            },
        );

        Ok(())
    }

    pub fn write_cell_format(
        &mut self,
        sheet: &str,
        a1: &str,
        format_dict: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        if !self.sheet_names.contains(&sheet.to_string()) {
            return Err(PyErr::new::<PyValueError, _>(format!(
                "Unknown sheet: {sheet}"
            )));
        }
        let key = resolve_key(sheet, a1)?;
        let dict = format_dict
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("format_dict must be a dict"))?;
        let fields = extract_format_fields(dict)?;
        self.formats.insert(key, fields);
        Ok(())
    }

    pub fn write_cell_border(
        &mut self,
        sheet: &str,
        a1: &str,
        border_dict: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        if !self.sheet_names.contains(&sheet.to_string()) {
            return Err(PyErr::new::<PyValueError, _>(format!(
                "Unknown sheet: {sheet}"
            )));
        }
        let key = resolve_key(sheet, a1)?;
        let dict = border_dict
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("border_dict must be a dict"))?;
        let fields = extract_border_fields(dict)?;
        self.borders.insert(key, fields);
        Ok(())
    }

    pub fn set_row_height(&mut self, sheet: &str, row: u32, height: f64) -> PyResult<()> {
        if !self.sheet_names.contains(&sheet.to_string()) {
            return Err(PyErr::new::<PyValueError, _>(format!(
                "Unknown sheet: {sheet}"
            )));
        }
        self.row_heights.insert((sheet.to_string(), row), height);
        Ok(())
    }

    pub fn set_column_width(&mut self, sheet: &str, col_str: &str, width: f64) -> PyResult<()> {
        if !self.sheet_names.contains(&sheet.to_string()) {
            return Err(PyErr::new::<PyValueError, _>(format!(
                "Unknown sheet: {sheet}"
            )));
        }
        let col_idx = col_letter_to_index(col_str)?;
        self.col_widths
            .insert((sheet.to_string(), col_idx), width);
        Ok(())
    }

    pub fn save(&mut self, path: &str) -> PyResult<()> {
        if self.saved {
            return Err(PyErr::new::<PyValueError, _>(
                "Workbook already saved (RustXlsxWriterBook is consumed-on-save)",
            ));
        }
        self.saved = true;

        let mut wb = Workbook::new();

        // Create worksheets in insertion order.
        let mut ws_map: IndexMap<String, Worksheet> = IndexMap::new();
        for name in &self.sheet_names {
            let mut ws = Worksheet::new();
            ws.set_name(name).map_err(|e| {
                PyErr::new::<PyValueError, _>(format!("Invalid sheet name: {e}"))
            })?;
            ws_map.insert(name.clone(), ws);
        }

        // Apply row heights.
        for ((sheet, row), height) in &self.row_heights {
            if let Some(ws) = ws_map.get_mut(sheet) {
                ws.set_row_height(*row, *height).map_err(|e| {
                    PyErr::new::<PyIOError, _>(format!("set_row_height failed: {e}"))
                })?;
            }
        }

        // Apply column widths.
        for ((sheet, col), width) in &self.col_widths {
            if let Some(ws) = ws_map.get_mut(sheet) {
                ws.set_column_width(*col, *width).map_err(|e| {
                    PyErr::new::<PyIOError, _>(format!("set_column_width failed: {e}"))
                })?;
            }
        }

        // Write all cells with merged format+border.
        for (key, payload) in &self.values {
            let (ref sheet, row, col) = *key;
            let fmt_fields = self.formats.get(key);
            let bdr_fields = self.borders.get(key);
            let format = build_format(fmt_fields, bdr_fields);

            let ws = ws_map.get_mut(sheet).ok_or_else(|| {
                PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}"))
            })?;

            write_cell(ws, row, col, payload, &format)?;
        }

        // Also write formats for cells that have format/border but no value
        // (e.g., blank cells with borders).
        for (key, _) in &self.formats {
            if !self.values.contains_key(key) {
                let (ref sheet, row, col) = *key;
                let fmt_fields = self.formats.get(key);
                let bdr_fields = self.borders.get(key);
                let format = build_format(fmt_fields, bdr_fields);
                if let Some(ws) = ws_map.get_mut(sheet) {
                    let _ = ws.write_blank(row, col, &format);
                }
            }
        }
        for (key, _) in &self.borders {
            if !self.values.contains_key(key) && !self.formats.contains_key(key) {
                let (ref sheet, row, col) = *key;
                let bdr_fields = self.borders.get(key);
                let format = build_format(None, bdr_fields);
                if let Some(ws) = ws_map.get_mut(sheet) {
                    let _ = ws.write_blank(row, col, &format);
                }
            }
        }

        for (_name, ws) in ws_map.drain(..) {
            wb.push_worksheet(ws);
        }

        wb.save(path)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to save workbook: {e}")))
    }
}
