use pyo3::exceptions::{PyIOError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::PyDict;

use chrono::{NaiveDate, NaiveDateTime};
use indexmap::IndexMap;
use rust_xlsxwriter::{
    Color, ConditionalFormat3ColorScale, ConditionalFormatCell, ConditionalFormatCellRule,
    ConditionalFormatDataBar, ConditionalFormatFormula, DataValidation, DataValidationRule, Format,
    FormatAlign, FormatBorder, FormatDiagonalBorder, FormatPattern, FormatUnderline, Formula,
    Image, Note, Url, Workbook, Worksheet,
};

use std::collections::{HashMap, HashSet};

use crate::util::{a1_to_row_col, parse_iso_date, parse_iso_datetime};

#[derive(Clone, Debug)]
enum CellPayload {
    Blank,
    String(String),
    Number(f64),
    Boolean(bool),
    Formula(String),
    Error(String),
    Date(NaiveDate),
    DateTime(NaiveDateTime),
}

#[derive(Clone, Debug, Default)]
struct CellFormatSpec {
    bold: Option<bool>,
    italic: Option<bool>,
    underline: Option<String>,
    strikethrough: Option<bool>,
    font_name: Option<String>,
    font_size: Option<f64>,
    font_color: Option<u32>,
    bg_color: Option<u32>,
    number_format: Option<String>,
    h_align: Option<String>,
    v_align: Option<String>,
    wrap: Option<bool>,
    rotation: Option<i16>,
    indent: Option<u8>,
}

#[derive(Clone, Debug, Default)]
struct BorderEdgeSpec {
    style: Option<String>,
    color: Option<u32>,
}

#[derive(Clone, Debug, Default)]
struct BorderSpec {
    top: Option<BorderEdgeSpec>,
    bottom: Option<BorderEdgeSpec>,
    left: Option<BorderEdgeSpec>,
    right: Option<BorderEdgeSpec>,
    diagonal_up: Option<BorderEdgeSpec>,
    diagonal_down: Option<BorderEdgeSpec>,
}

#[derive(Clone, Debug, Default)]
struct SheetState {
    cells: HashMap<(u32, u16), CellPayload>,
    formats: HashMap<(u32, u16), CellFormatSpec>,
    borders: HashMap<(u32, u16), BorderSpec>,
    row_heights: HashMap<u32, f64>, // Excel row number (1-based)
    col_widths: HashMap<u16, f64>,  // 0-based column index

    merges: Vec<MergeSpec>,
    freeze: Option<FreezeSpec>,
    conditional_formats: Vec<ConditionalFormatSpec>,
    data_validations: Vec<DataValidationSpec>,
    hyperlinks: Vec<HyperlinkSpec>,
    images: Vec<ImageSpec>,
    notes: Vec<NoteSpec>,
}

#[derive(Clone, Debug)]
struct MergeSpec {
    first_row: u32,
    first_col: u16,
    last_row: u32,
    last_col: u16,
}

#[derive(Clone, Debug)]
struct FreezeSpec {
    mode: String,
    top_left_cell: Option<(u32, u16)>,
    x_split: Option<u32>,
    y_split: Option<u32>,
}

#[derive(Clone, Debug)]
enum ConditionalFormatKind {
    CellIs {
        operator: String,
        formula: String,
        bg_color: Option<u32>,
        font_color: Option<u32>,
        stop_if_true: bool,
    },
    Expression {
        formula: String,
        bg_color: Option<u32>,
        font_color: Option<u32>,
        stop_if_true: bool,
    },
    DataBar,
    ColorScale,
}

#[derive(Clone, Debug)]
struct ConditionalFormatSpec {
    first_row: u32,
    first_col: u16,
    last_row: u32,
    last_col: u16,
    kind: ConditionalFormatKind,
}

#[derive(Clone, Debug)]
struct DataValidationSpec {
    first_row: u32,
    first_col: u16,
    last_row: u32,
    last_col: u16,
    validation_type: String,
    operator: Option<String>,
    formula1: Option<String>,
    formula2: Option<String>,
    allow_blank: Option<bool>,
    prompt_title: Option<String>,
    prompt: Option<String>,
    error_title: Option<String>,
    error: Option<String>,
}

#[derive(Clone, Debug)]
struct HyperlinkSpec {
    row: u32,
    col: u16,
    target: String,
    display: Option<String>,
    tooltip: Option<String>,
    internal: bool,
}

#[derive(Clone, Debug)]
struct ImageSpec {
    row: u32,
    col: u16,
    path: String,
    x_offset: u32,
    y_offset: u32,
}

#[derive(Clone, Debug)]
struct NoteSpec {
    row: u32,
    col: u16,
    text: String,
    author: Option<String>,
}

fn parse_rgb_color(s: &str) -> Option<u32> {
    let s = s.trim();
    let s = s.strip_prefix('#').unwrap_or(s);
    let s = s.strip_prefix("0x").unwrap_or(s);
    if s.len() != 6 {
        return None;
    }
    u32::from_str_radix(s, 16).ok()
}

fn normalize_a1(s: &str) -> String {
    s.replace('$', "")
}

fn parse_a1_cell(a1: &str) -> PyResult<(u32, u16)> {
    let a1 = normalize_a1(a1);
    let (row0, col0) = a1_to_row_col(&a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
    let col: u16 = col0.try_into().map_err(|_| {
        PyErr::new::<PyValueError, _>(format!("Column out of range for Excel: {a1}"))
    })?;
    Ok((row0, col))
}

fn parse_a1_range(range: &str) -> PyResult<((u32, u16), (u32, u16))> {
    let range = normalize_a1(range);
    if let Some((a, b)) = range.split_once(':') {
        let (r1, c1) = parse_a1_cell(a)?;
        let (r2, c2) = parse_a1_cell(b)?;
        let first_row = r1.min(r2);
        let last_row = r1.max(r2);
        let first_col = c1.min(c2);
        let last_col = c1.max(c2);
        Ok(((first_row, first_col), (last_row, last_col)))
    } else {
        let (r, c) = parse_a1_cell(&range)?;
        Ok(((r, c), (r, c)))
    }
}

fn build_cf_format(bg_color: Option<u32>, font_color: Option<u32>) -> Option<Format> {
    let mut used = false;
    let mut fmt = Format::new();
    if let Some(rgb) = bg_color {
        // Conditional format fills are stored as foreground colors (dxf fills).
        fmt = fmt
            .set_foreground_color(Color::RGB(rgb))
            .set_pattern(FormatPattern::Solid);
        used = true;
    }
    if let Some(rgb) = font_color {
        fmt = fmt.set_font_color(Color::RGB(rgb));
        used = true;
    }
    if used {
        Some(fmt)
    } else {
        None
    }
}

fn parse_cf_operator_rule(operator: &str, formula: &str) -> Option<ConditionalFormatCellRule<f64>> {
    let v = formula.trim().parse::<f64>().ok()?;
    match operator {
        "greaterThan" => Some(ConditionalFormatCellRule::GreaterThan(v)),
        "greaterThanOrEqual" => Some(ConditionalFormatCellRule::GreaterThanOrEqualTo(v)),
        "lessThan" => Some(ConditionalFormatCellRule::LessThan(v)),
        "lessThanOrEqual" => Some(ConditionalFormatCellRule::LessThanOrEqualTo(v)),
        "equal" => Some(ConditionalFormatCellRule::EqualTo(v)),
        "notEqual" => Some(ConditionalFormatCellRule::NotEqualTo(v)),
        _ => None,
    }
}

fn format_underline_from_str(s: &str) -> Option<FormatUnderline> {
    match s {
        "single" => Some(FormatUnderline::Single),
        "double" => Some(FormatUnderline::Double),
        _ => None,
    }
}

fn format_h_align_from_str(s: &str) -> Option<FormatAlign> {
    match s {
        "general" => Some(FormatAlign::General),
        "left" => Some(FormatAlign::Left),
        "center" => Some(FormatAlign::Center),
        "right" => Some(FormatAlign::Right),
        "fill" => Some(FormatAlign::Fill),
        "justify" => Some(FormatAlign::Justify),
        "centerAcross" => Some(FormatAlign::CenterAcross),
        "distributed" => Some(FormatAlign::Distributed),
        _ => None,
    }
}

fn format_v_align_from_str(s: &str) -> Option<FormatAlign> {
    match s {
        "top" => Some(FormatAlign::Top),
        "center" | "vcenter" | "verticalCenter" | "centerContinuous" => {
            Some(FormatAlign::VerticalCenter)
        }
        "bottom" => Some(FormatAlign::Bottom),
        "verticalJustify" => Some(FormatAlign::VerticalJustify),
        "verticalDistributed" => Some(FormatAlign::VerticalDistributed),
        _ => None,
    }
}

fn format_border_from_str(s: &str) -> Option<FormatBorder> {
    match s {
        "none" => Some(FormatBorder::None),
        "thin" => Some(FormatBorder::Thin),
        "medium" => Some(FormatBorder::Medium),
        "thick" => Some(FormatBorder::Thick),
        "double" => Some(FormatBorder::Double),
        "dashed" => Some(FormatBorder::Dashed),
        "dotted" => Some(FormatBorder::Dotted),
        "hair" => Some(FormatBorder::Hair),
        "mediumDashed" => Some(FormatBorder::MediumDashed),
        "dashDot" => Some(FormatBorder::DashDot),
        "mediumDashDot" => Some(FormatBorder::MediumDashDot),
        "dashDotDot" => Some(FormatBorder::DashDotDot),
        "mediumDashDotDot" => Some(FormatBorder::MediumDashDotDot),
        "slantDashDot" => Some(FormatBorder::SlantDashDot),
        _ => None,
    }
}

fn build_format(
    cell_type: &CellPayload,
    fmt_spec: Option<&CellFormatSpec>,
    border_spec: Option<&BorderSpec>,
) -> Option<Format> {
    let mut used = false;
    let mut fmt = Format::new();

    if let Some(spec) = fmt_spec {
        if spec.bold == Some(true) {
            fmt = fmt.set_bold();
            used = true;
        }
        if spec.italic == Some(true) {
            fmt = fmt.set_italic();
            used = true;
        }
        if spec.strikethrough == Some(true) {
            fmt = fmt.set_font_strikethrough();
            used = true;
        }
        if let Some(u) = &spec.underline {
            if let Some(ul) = format_underline_from_str(u) {
                fmt = fmt.set_underline(ul);
                used = true;
            }
        }
        if let Some(name) = &spec.font_name {
            fmt = fmt.set_font_name(name);
            used = true;
        }
        if let Some(sz) = spec.font_size {
            fmt = fmt.set_font_size(sz);
            used = true;
        }
        if let Some(rgb) = spec.font_color {
            fmt = fmt.set_font_color(Color::RGB(rgb));
            used = true;
        }
        if let Some(rgb) = spec.bg_color {
            fmt = fmt
                .set_background_color(Color::RGB(rgb))
                .set_pattern(FormatPattern::Solid);
            used = true;
        }
        if let Some(nf) = &spec.number_format {
            fmt = fmt.set_num_format(nf);
            used = true;
        }
        if let Some(a) = &spec.h_align {
            if let Some(align) = format_h_align_from_str(a) {
                fmt = fmt.set_align(align);
                used = true;
            }
        }
        if let Some(a) = &spec.v_align {
            if let Some(align) = format_v_align_from_str(a) {
                fmt = fmt.set_align(align);
                used = true;
            }
        }
        if let Some(w) = spec.wrap {
            if w {
                fmt = fmt.set_text_wrap();
            } else {
                fmt = fmt.unset_text_wrap();
            }
            used = true;
        }
        if let Some(rot) = spec.rotation {
            fmt = fmt.set_rotation(rot);
            used = true;
        }
        if let Some(indent) = spec.indent {
            fmt = fmt.set_indent(indent);
            used = true;
        }
    }

    // Date/datetime semantics: ensure we always have a num format even if the
    // harness didn't explicitly set one.
    let has_num_format = fmt_spec.and_then(|s| s.number_format.as_ref()).is_some();
    match cell_type {
        CellPayload::Date(_) if !has_num_format => {
            fmt = fmt.set_num_format("yyyy-mm-dd");
            used = true;
        }
        CellPayload::DateTime(_) if !has_num_format => {
            fmt = fmt.set_num_format("yyyy-mm-dd hh:mm:ss");
            used = true;
        }
        _ => {}
    }

    if let Some(b) = border_spec {
        if let Some(edge) = &b.top {
            if let Some(style) = edge.style.as_deref().and_then(format_border_from_str) {
                if style != FormatBorder::None {
                    fmt = fmt.set_border_top(style);
                    used = true;
                }
            }
            if let Some(rgb) = edge.color {
                fmt = fmt.set_border_top_color(Color::RGB(rgb));
                used = true;
            }
        }
        if let Some(edge) = &b.bottom {
            if let Some(style) = edge.style.as_deref().and_then(format_border_from_str) {
                if style != FormatBorder::None {
                    fmt = fmt.set_border_bottom(style);
                    used = true;
                }
            }
            if let Some(rgb) = edge.color {
                fmt = fmt.set_border_bottom_color(Color::RGB(rgb));
                used = true;
            }
        }
        if let Some(edge) = &b.left {
            if let Some(style) = edge.style.as_deref().and_then(format_border_from_str) {
                if style != FormatBorder::None {
                    fmt = fmt.set_border_left(style);
                    used = true;
                }
            }
            if let Some(rgb) = edge.color {
                fmt = fmt.set_border_left_color(Color::RGB(rgb));
                used = true;
            }
        }
        if let Some(edge) = &b.right {
            if let Some(style) = edge.style.as_deref().and_then(format_border_from_str) {
                if style != FormatBorder::None {
                    fmt = fmt.set_border_right(style);
                    used = true;
                }
            }
            if let Some(rgb) = edge.color {
                fmt = fmt.set_border_right_color(Color::RGB(rgb));
                used = true;
            }
        }

        let diag_up = b.diagonal_up.as_ref();
        let diag_down = b.diagonal_down.as_ref();
        if diag_up.is_some() || diag_down.is_some() {
            let edge = diag_up.or(diag_down);
            if let Some(edge) = edge {
                if let Some(style) = edge.style.as_deref().and_then(format_border_from_str) {
                    if style != FormatBorder::None {
                        fmt = fmt.set_border_diagonal(style);
                    }
                }
                if let Some(rgb) = edge.color {
                    fmt = fmt.set_border_diagonal_color(Color::RGB(rgb));
                }
            }

            let diag_type = match (diag_up.is_some(), diag_down.is_some()) {
                (true, true) => FormatDiagonalBorder::BorderUpDown,
                (true, false) => FormatDiagonalBorder::BorderUp,
                (false, true) => FormatDiagonalBorder::BorderDown,
                _ => FormatDiagonalBorder::BorderUpDown,
            };
            fmt = fmt.set_border_diagonal_type(diag_type);
            used = true;
        }
    }

    if used {
        Some(fmt)
    } else {
        None
    }
}

fn parse_cell_value_payload(dict: &Bound<'_, PyDict>) -> PyResult<CellPayload> {
    let type_obj = dict
        .get_item("type")?
        .ok_or_else(|| PyErr::new::<PyValueError, _>("payload missing 'type'"))?;
    let type_str: String = type_obj.extract()?;

    match type_str.as_str() {
        "blank" => Ok(CellPayload::Blank),
        "string" => {
            let v = dict.get_item("value")?;
            let s = match v {
                Some(v) => v.extract::<String>()?,
                None => String::new(),
            };
            Ok(CellPayload::String(s))
        }
        "number" => {
            let v = dict
                .get_item("value")?
                .ok_or_else(|| PyErr::new::<PyValueError, _>("number payload missing 'value'"))?;
            Ok(CellPayload::Number(v.extract::<f64>()?))
        }
        "boolean" => {
            let v = dict
                .get_item("value")?
                .ok_or_else(|| PyErr::new::<PyValueError, _>("boolean payload missing 'value'"))?;
            Ok(CellPayload::Boolean(v.extract::<bool>()?))
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
            Ok(CellPayload::Formula(v.extract::<String>()?))
        }
        "error" => {
            let v = dict
                .get_item("value")?
                .ok_or_else(|| PyErr::new::<PyValueError, _>("error payload missing 'value'"))?;
            Ok(CellPayload::Error(v.extract::<String>()?))
        }
        "date" => {
            let v = dict
                .get_item("value")?
                .ok_or_else(|| PyErr::new::<PyValueError, _>("date payload missing 'value'"))?;
            let s = v.extract::<String>()?;
            if let Some(d) = parse_iso_date(&s) {
                Ok(CellPayload::Date(d))
            } else {
                Ok(CellPayload::String(s))
            }
        }
        "datetime" => {
            let v = dict
                .get_item("value")?
                .ok_or_else(|| PyErr::new::<PyValueError, _>("datetime payload missing 'value'"))?;
            let s = v.extract::<String>()?;
            if let Some(dt) = parse_iso_datetime(&s) {
                Ok(CellPayload::DateTime(dt))
            } else {
                Ok(CellPayload::String(s))
            }
        }
        other => Err(PyErr::new::<PyValueError, _>(format!(
            "Unsupported cell type: {other}"
        ))),
    }
}

fn parse_cell_format_payload(dict: &Bound<'_, PyDict>) -> PyResult<CellFormatSpec> {
    let mut spec = CellFormatSpec::default();

    if let Some(v) = dict.get_item("bold")? {
        spec.bold = Some(v.extract::<bool>()?);
    }
    if let Some(v) = dict.get_item("italic")? {
        spec.italic = Some(v.extract::<bool>()?);
    }
    if let Some(v) = dict.get_item("underline")? {
        spec.underline = Some(v.extract::<String>()?);
    }
    if let Some(v) = dict.get_item("strikethrough")? {
        spec.strikethrough = Some(v.extract::<bool>()?);
    }
    if let Some(v) = dict.get_item("font_name")? {
        spec.font_name = Some(v.extract::<String>()?);
    }
    if let Some(v) = dict.get_item("font_size")? {
        spec.font_size = Some(v.extract::<f64>()?);
    }
    if let Some(v) = dict.get_item("font_color")? {
        let s = v.extract::<String>()?;
        spec.font_color = parse_rgb_color(&s);
    }
    if let Some(v) = dict.get_item("bg_color")? {
        let s = v.extract::<String>()?;
        spec.bg_color = parse_rgb_color(&s);
    }
    if let Some(v) = dict.get_item("number_format")? {
        spec.number_format = Some(v.extract::<String>()?);
    }
    if let Some(v) = dict.get_item("h_align")? {
        spec.h_align = Some(v.extract::<String>()?);
    }
    if let Some(v) = dict.get_item("v_align")? {
        spec.v_align = Some(v.extract::<String>()?);
    }
    if let Some(v) = dict.get_item("wrap")? {
        spec.wrap = Some(v.extract::<bool>()?);
    }
    if let Some(v) = dict.get_item("rotation")? {
        let i = v.extract::<i64>()?;
        spec.rotation = Some(i as i16);
    }
    if let Some(v) = dict.get_item("indent")? {
        let i = v.extract::<i64>()?;
        if i >= 0 {
            spec.indent = Some(i as u8);
        }
    }

    Ok(spec)
}

fn parse_border_edge_payload(dict: &Bound<'_, PyDict>) -> PyResult<BorderEdgeSpec> {
    let mut edge = BorderEdgeSpec::default();
    if let Some(v) = dict.get_item("style")? {
        edge.style = Some(v.extract::<String>()?);
    }
    if let Some(v) = dict.get_item("color")? {
        let s = v.extract::<String>()?;
        edge.color = parse_rgb_color(&s);
    }
    Ok(edge)
}

fn parse_border_payload(dict: &Bound<'_, PyDict>) -> PyResult<BorderSpec> {
    let mut border = BorderSpec::default();

    if let Some(v) = dict.get_item("top")? {
        let d = v.downcast::<PyDict>()?;
        border.top = Some(parse_border_edge_payload(&d)?);
    }
    if let Some(v) = dict.get_item("bottom")? {
        let d = v.downcast::<PyDict>()?;
        border.bottom = Some(parse_border_edge_payload(&d)?);
    }
    if let Some(v) = dict.get_item("left")? {
        let d = v.downcast::<PyDict>()?;
        border.left = Some(parse_border_edge_payload(&d)?);
    }
    if let Some(v) = dict.get_item("right")? {
        let d = v.downcast::<PyDict>()?;
        border.right = Some(parse_border_edge_payload(&d)?);
    }
    if let Some(v) = dict.get_item("diagonal_up")? {
        let d = v.downcast::<PyDict>()?;
        border.diagonal_up = Some(parse_border_edge_payload(&d)?);
    }
    if let Some(v) = dict.get_item("diagonal_down")? {
        let d = v.downcast::<PyDict>()?;
        border.diagonal_down = Some(parse_border_edge_payload(&d)?);
    }

    Ok(border)
}

#[pyclass(unsendable)]
pub struct RustXlsxWriterBook {
    sheets: IndexMap<String, SheetState>,
    saved: bool,
}

#[pymethods]
impl RustXlsxWriterBook {
    #[new]
    pub fn new() -> Self {
        Self {
            sheets: IndexMap::new(),
            saved: false,
        }
    }

    pub fn add_sheet(&mut self, name: &str) -> PyResult<()> {
        if self.sheets.contains_key(name) {
            return Ok(());
        }

        // Validate sheet name using rust_xlsxwriter.
        let mut ws = Worksheet::new();
        ws.set_name(name)
            .map_err(|e| PyErr::new::<PyValueError, _>(format!("Invalid sheet name: {e}")))?;

        self.sheets.insert(name.to_string(), SheetState::default());
        Ok(())
    }

    pub fn write_cell_value(
        &mut self,
        sheet: &str,
        a1: &str,
        payload: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let sheet_state = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let (row0, col0) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
        let col: u16 = col0.try_into().map_err(|_| {
            PyErr::new::<PyValueError, _>(format!("Column out of range for Excel: {a1}"))
        })?;

        let dict = payload
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("payload must be a dict"))?;
        let cell = parse_cell_value_payload(&dict)?;
        sheet_state.cells.insert((row0, col), cell);
        Ok(())
    }

    pub fn write_cell_format(
        &mut self,
        sheet: &str,
        a1: &str,
        payload: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let sheet_state = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;
        let (row0, col0) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
        let col: u16 = col0.try_into().map_err(|_| {
            PyErr::new::<PyValueError, _>(format!("Column out of range for Excel: {a1}"))
        })?;

        let dict = payload
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("payload must be a dict"))?;
        sheet_state
            .formats
            .insert((row0, col), parse_cell_format_payload(&dict)?);
        Ok(())
    }

    pub fn write_cell_border(
        &mut self,
        sheet: &str,
        a1: &str,
        payload: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let sheet_state = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;
        let (row0, col0) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
        let col: u16 = col0.try_into().map_err(|_| {
            PyErr::new::<PyValueError, _>(format!("Column out of range for Excel: {a1}"))
        })?;

        let dict = payload
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("payload must be a dict"))?;
        sheet_state
            .borders
            .insert((row0, col), parse_border_payload(&dict)?);
        Ok(())
    }

    pub fn set_row_height(&mut self, sheet: &str, row: u32, height: f64) -> PyResult<()> {
        let sheet_state = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;
        sheet_state.row_heights.insert(row, height);
        Ok(())
    }

    pub fn set_column_width(&mut self, sheet: &str, column: &str, width: f64) -> PyResult<()> {
        let sheet_state = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let a1 = format!("{column}1");
        let (_row0, col0) = a1_to_row_col(&a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
        let col: u16 = col0
            .try_into()
            .map_err(|_| PyErr::new::<PyValueError, _>(format!("Column out of range: {column}")))?;
        sheet_state.col_widths.insert(col, width);
        Ok(())
    }

    pub fn merge_cells(&mut self, sheet: &str, cell_range: &str) -> PyResult<()> {
        let sheet_state = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let ((r1, c1), (r2, c2)) = parse_a1_range(cell_range)?;
        if r1 == r2 && c1 == c2 {
            // Excel doesn't allow single-cell merges; treat as no-op.
            return Ok(());
        }
        sheet_state.merges.push(MergeSpec {
            first_row: r1,
            first_col: c1,
            last_row: r2,
            last_col: c2,
        });
        Ok(())
    }

    pub fn add_conditional_format(
        &mut self,
        sheet: &str,
        payload: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let sheet_state = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let outer = payload
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("payload must be a dict"))?;
        let cf_any = outer
            .get_item("cf_rule")?
            .unwrap_or_else(|| outer.clone().into_any());
        let cf = cf_any
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("cf_rule must be a dict"))?;

        let range = cf
            .get_item("range")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("cf_rule missing 'range'"))?
            .extract::<String>()?;
        let rule_type = cf
            .get_item("rule_type")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("cf_rule missing 'rule_type'"))?
            .extract::<String>()?;

        let stop_if_true = cf
            .get_item("stop_if_true")?
            .and_then(|v| v.extract::<bool>().ok())
            .unwrap_or(false);

        let fmt_any = cf.get_item("format")?;
        let mut bg_color: Option<u32> = None;
        let mut font_color: Option<u32> = None;
        if let Some(fmt_any) = fmt_any {
            if let Ok(fmt) = fmt_any.downcast::<PyDict>() {
                if let Some(v) = fmt.get_item("bg_color")? {
                    let s = v.extract::<String>()?;
                    bg_color = parse_rgb_color(&s);
                }
                if let Some(v) = fmt.get_item("font_color")? {
                    let s = v.extract::<String>()?;
                    font_color = parse_rgb_color(&s);
                }
            }
        }

        let ((r1, c1), (r2, c2)) = parse_a1_range(&range)?;

        let kind = match rule_type.as_str() {
            "cellIs" | "cellIsRule" => {
                let operator = cf
                    .get_item("operator")?
                    .and_then(|v| v.extract::<String>().ok())
                    .unwrap_or_else(|| "equal".to_string());
                let formula = cf
                    .get_item("formula")?
                    .and_then(|v| v.extract::<String>().ok())
                    .unwrap_or_default();
                ConditionalFormatKind::CellIs {
                    operator,
                    formula,
                    bg_color,
                    font_color,
                    stop_if_true,
                }
            }
            "expression" | "formula" => {
                let formula = cf
                    .get_item("formula")?
                    .and_then(|v| v.extract::<String>().ok())
                    .unwrap_or_default();
                ConditionalFormatKind::Expression {
                    formula,
                    bg_color,
                    font_color,
                    stop_if_true,
                }
            }
            "dataBar" => ConditionalFormatKind::DataBar,
            "colorScale" => ConditionalFormatKind::ColorScale,
            _ => {
                // Unsupported rule types are a no-op for this backend.
                return Ok(());
            }
        };

        sheet_state.conditional_formats.push(ConditionalFormatSpec {
            first_row: r1,
            first_col: c1,
            last_row: r2,
            last_col: c2,
            kind,
        });
        Ok(())
    }

    pub fn add_data_validation(&mut self, sheet: &str, payload: &Bound<'_, PyAny>) -> PyResult<()> {
        let sheet_state = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let outer = payload
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("payload must be a dict"))?;
        let v_any = outer
            .get_item("validation")?
            .unwrap_or_else(|| outer.clone().into_any());
        let v = v_any
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("validation must be a dict"))?;

        let range = v
            .get_item("range")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("validation missing 'range'"))?
            .extract::<String>()?;
        let validation_type = v
            .get_item("validation_type")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("validation missing 'validation_type'"))?
            .extract::<String>()?;

        let ((r1, c1), (r2, c2)) = parse_a1_range(&range)?;

        let spec = DataValidationSpec {
            first_row: r1,
            first_col: c1,
            last_row: r2,
            last_col: c2,
            validation_type,
            operator: v
                .get_item("operator")?
                .and_then(|x| x.extract::<String>().ok()),
            formula1: v
                .get_item("formula1")?
                .and_then(|x| x.extract::<String>().ok()),
            formula2: v
                .get_item("formula2")?
                .and_then(|x| x.extract::<String>().ok()),
            allow_blank: v
                .get_item("allow_blank")?
                .and_then(|x| x.extract::<bool>().ok()),
            prompt_title: v
                .get_item("prompt_title")?
                .and_then(|x| x.extract::<String>().ok()),
            prompt: v
                .get_item("prompt")?
                .and_then(|x| x.extract::<String>().ok()),
            error_title: v
                .get_item("error_title")?
                .and_then(|x| x.extract::<String>().ok()),
            error: v
                .get_item("error")?
                .and_then(|x| x.extract::<String>().ok()),
        };

        sheet_state.data_validations.push(spec);
        Ok(())
    }

    pub fn add_hyperlink(&mut self, sheet: &str, payload: &Bound<'_, PyAny>) -> PyResult<()> {
        let sheet_state = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let outer = payload
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("payload must be a dict"))?;
        let link_any = outer
            .get_item("hyperlink")?
            .unwrap_or_else(|| outer.clone().into_any());
        let link = link_any
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("hyperlink must be a dict"))?;

        let cell = link
            .get_item("cell")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("hyperlink missing 'cell'"))?
            .extract::<String>()?;
        let target = link
            .get_item("target")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("hyperlink missing 'target'"))?
            .extract::<String>()?;
        let display = link
            .get_item("display")?
            .and_then(|v| v.extract::<String>().ok());
        let tooltip = link
            .get_item("tooltip")?
            .and_then(|v| v.extract::<String>().ok());
        let internal = link
            .get_item("internal")?
            .and_then(|v| v.extract::<bool>().ok())
            .unwrap_or(false);

        let (row, col) = parse_a1_cell(&cell)?;
        sheet_state.hyperlinks.push(HyperlinkSpec {
            row,
            col,
            target,
            display,
            tooltip,
            internal,
        });
        Ok(())
    }

    pub fn add_image(&mut self, sheet: &str, payload: &Bound<'_, PyAny>) -> PyResult<()> {
        let sheet_state = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let outer = payload
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("payload must be a dict"))?;
        let img_any = outer
            .get_item("image")?
            .unwrap_or_else(|| outer.clone().into_any());
        let img = img_any
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("image must be a dict"))?;

        let cell = img
            .get_item("cell")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("image missing 'cell'"))?
            .extract::<String>()?;
        let path = img
            .get_item("path")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("image missing 'path'"))?
            .extract::<String>()?;

        let mut x_offset: u32 = 0;
        let mut y_offset: u32 = 0;
        if let Some(v) = img.get_item("offset")? {
            if let Ok(pair) = v.extract::<(u32, u32)>() {
                x_offset = pair.0;
                y_offset = pair.1;
            }
        }

        let (row, col) = parse_a1_cell(&cell)?;
        sheet_state.images.push(ImageSpec {
            row,
            col,
            path,
            x_offset,
            y_offset,
        });
        Ok(())
    }

    pub fn add_comment(&mut self, sheet: &str, payload: &Bound<'_, PyAny>) -> PyResult<()> {
        let sheet_state = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let outer = payload
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("payload must be a dict"))?;
        let note_any = outer
            .get_item("comment")?
            .unwrap_or_else(|| outer.clone().into_any());
        let note = note_any
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("comment must be a dict"))?;

        let cell = note
            .get_item("cell")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("comment missing 'cell'"))?
            .extract::<String>()?;
        let text = note
            .get_item("text")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("comment missing 'text'"))?
            .extract::<String>()?;
        let author = note
            .get_item("author")?
            .and_then(|v| v.extract::<String>().ok());

        let (row, col) = parse_a1_cell(&cell)?;
        sheet_state.notes.push(NoteSpec {
            row,
            col,
            text,
            author,
        });
        Ok(())
    }

    pub fn set_freeze_panes(&mut self, sheet: &str, payload: &Bound<'_, PyAny>) -> PyResult<()> {
        let sheet_state = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let outer = payload
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("payload must be a dict"))?;
        let cfg_any = outer
            .get_item("freeze")?
            .unwrap_or_else(|| outer.clone().into_any());
        let cfg = cfg_any
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("freeze must be a dict"))?;

        let mode = cfg
            .get_item("mode")?
            .and_then(|v| v.extract::<String>().ok())
            .unwrap_or_else(|| "freeze".to_string());
        let top_left_cell = cfg
            .get_item("top_left_cell")?
            .and_then(|v| v.extract::<String>().ok())
            .and_then(|s| parse_a1_cell(&s).ok());
        let x_split = cfg
            .get_item("x_split")?
            .and_then(|v| v.extract::<u32>().ok());
        let y_split = cfg
            .get_item("y_split")?
            .and_then(|v| v.extract::<u32>().ok());

        sheet_state.freeze = Some(FreezeSpec {
            mode,
            top_left_cell,
            x_split,
            y_split,
        });
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

        for (name, state) in self.sheets.drain(..) {
            let mut ws = Worksheet::new();
            ws.set_name(&name)
                .map_err(|e| PyErr::new::<PyValueError, _>(format!("Invalid sheet name: {e}")))?;

            for (row1, height) in state.row_heights {
                if row1 == 0 {
                    continue;
                }
                ws.set_row_height(row1 - 1, height).map_err(|e| {
                    PyErr::new::<PyIOError, _>(format!("set_row_height failed: {e}"))
                })?;
            }
            for (col0, width) in state.col_widths {
                ws.set_column_width(col0, width).map_err(|e| {
                    PyErr::new::<PyIOError, _>(format!("set_column_width failed: {e}"))
                })?;
            }

            // Freeze panes (split panes aren't supported by rust_xlsxwriter).
            if let Some(freeze) = &state.freeze {
                if freeze.mode == "freeze" {
                    if let Some((row0, col0)) = freeze.top_left_cell {
                        ws.set_freeze_panes(row0, col0).map_err(|e| {
                            PyErr::new::<PyIOError, _>(format!("set_freeze_panes failed: {e}"))
                        })?;
                    }
                }
            }

            // Apply merged ranges via merge_range().
            let mut merged_cells: HashSet<(u32, u16)> = HashSet::new();
            for m in &state.merges {
                for r in m.first_row..=m.last_row {
                    for c in m.first_col..=m.last_col {
                        merged_cells.insert((r, c));
                    }
                }

                let cell = state
                    .cells
                    .get(&(m.first_row, m.first_col))
                    .unwrap_or(&CellPayload::Blank);

                let value = match cell {
                    CellPayload::Blank => "".to_string(),
                    CellPayload::String(s) => s.clone(),
                    CellPayload::Number(n) => n.to_string(),
                    CellPayload::Boolean(b) => b.to_string(),
                    CellPayload::Formula(f) => f.clone(),
                    CellPayload::Error(t) => t.clone(),
                    CellPayload::Date(d) => d.format("%Y-%m-%d").to_string(),
                    CellPayload::DateTime(dt) => dt.format("%Y-%m-%dT%H:%M:%S").to_string(),
                };

                let fmt = build_format(
                    cell,
                    state.formats.get(&(m.first_row, m.first_col)),
                    state.borders.get(&(m.first_row, m.first_col)),
                );
                let default_fmt = Format::new();
                let fmt_ref = fmt.as_ref().unwrap_or(&default_fmt);

                ws.merge_range(
                    m.first_row,
                    m.first_col,
                    m.last_row,
                    m.last_col,
                    &value,
                    fmt_ref,
                )
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("merge_range failed: {e}")))?;
            }

            // Conditional formats.
            for spec in &state.conditional_formats {
                match &spec.kind {
                    ConditionalFormatKind::CellIs {
                        operator,
                        formula,
                        bg_color,
                        font_color,
                        stop_if_true,
                    } => {
                        let rule = match parse_cf_operator_rule(operator, formula) {
                            Some(r) => r,
                            None => continue,
                        };

                        let mut cf = ConditionalFormatCell::new().set_rule(rule);
                        if *stop_if_true {
                            cf = cf.set_stop_if_true(true);
                        }
                        if let Some(fmt) = build_cf_format(*bg_color, *font_color) {
                            cf = cf.set_format(fmt);
                        }
                        ws.add_conditional_format(
                            spec.first_row,
                            spec.first_col,
                            spec.last_row,
                            spec.last_col,
                            &cf,
                        )
                        .map_err(|e| {
                            PyErr::new::<PyIOError, _>(format!(
                                "add_conditional_format failed: {e}"
                            ))
                        })?;
                    }
                    ConditionalFormatKind::Expression {
                        formula,
                        bg_color,
                        font_color,
                        stop_if_true,
                    } => {
                        let rule = formula.trim();
                        let rule = rule.strip_prefix('=').unwrap_or(rule);
                        let mut cf = ConditionalFormatFormula::new().set_rule(rule);
                        if *stop_if_true {
                            cf = cf.set_stop_if_true(true);
                        }
                        if let Some(fmt) = build_cf_format(*bg_color, *font_color) {
                            cf = cf.set_format(fmt);
                        }
                        ws.add_conditional_format(
                            spec.first_row,
                            spec.first_col,
                            spec.last_row,
                            spec.last_col,
                            &cf,
                        )
                        .map_err(|e| {
                            PyErr::new::<PyIOError, _>(format!(
                                "add_conditional_format failed: {e}"
                            ))
                        })?;
                    }
                    ConditionalFormatKind::DataBar => {
                        let cf = ConditionalFormatDataBar::new();
                        ws.add_conditional_format(
                            spec.first_row,
                            spec.first_col,
                            spec.last_row,
                            spec.last_col,
                            &cf,
                        )
                        .map_err(|e| {
                            PyErr::new::<PyIOError, _>(format!(
                                "add_conditional_format failed: {e}"
                            ))
                        })?;
                    }
                    ConditionalFormatKind::ColorScale => {
                        let cf = ConditionalFormat3ColorScale::new();
                        ws.add_conditional_format(
                            spec.first_row,
                            spec.first_col,
                            spec.last_row,
                            spec.last_col,
                            &cf,
                        )
                        .map_err(|e| {
                            PyErr::new::<PyIOError, _>(format!(
                                "add_conditional_format failed: {e}"
                            ))
                        })?;
                    }
                }
            }

            // Data validations.
            for spec in &state.data_validations {
                let mut dv = DataValidation::new();

                match spec.validation_type.as_str() {
                    "list" => {
                        if let Some(f1) = &spec.formula1 {
                            let f1 = f1.trim();
                            if f1.starts_with('"') && f1.ends_with('"') {
                                let inner = &f1[1..f1.len() - 1];
                                let parts: Vec<&str> = inner.split(',').collect();
                                dv = dv.allow_list_strings(&parts).map_err(|e| {
                                    PyErr::new::<PyValueError, _>(format!(
                                        "allow_list_strings failed: {e}"
                                    ))
                                })?;
                            } else {
                                dv = dv.allow_list_formula(Formula::new(f1));
                            }
                        }
                    }
                    "custom" => {
                        if let Some(f1) = &spec.formula1 {
                            dv = dv.allow_custom(Formula::new(f1));
                        }
                    }
                    "whole" => {
                        let op = spec
                            .operator
                            .clone()
                            .unwrap_or_else(|| "between".to_string());
                        let f1 = spec
                            .formula1
                            .as_deref()
                            .unwrap_or("0")
                            .trim()
                            .parse::<i32>()
                            .unwrap_or(0);
                        let f2 = spec
                            .formula2
                            .as_deref()
                            .unwrap_or("0")
                            .trim()
                            .parse::<i32>()
                            .unwrap_or(0);
                        let rule = match op.as_str() {
                            "between" => DataValidationRule::Between(f1, f2),
                            "notBetween" => DataValidationRule::NotBetween(f1, f2),
                            "greaterThan" => DataValidationRule::GreaterThan(f1),
                            "greaterThanOrEqual" => DataValidationRule::GreaterThanOrEqualTo(f1),
                            "lessThan" => DataValidationRule::LessThan(f1),
                            "lessThanOrEqual" => DataValidationRule::LessThanOrEqualTo(f1),
                            "equal" => DataValidationRule::EqualTo(f1),
                            "notEqual" => DataValidationRule::NotEqualTo(f1),
                            _ => DataValidationRule::Between(f1, f2),
                        };
                        dv = dv.allow_whole_number(rule);
                    }
                    _ => {
                        // Unsupported types are ignored.
                        continue;
                    }
                }

                if let Some(allow) = spec.allow_blank {
                    dv = dv.ignore_blank(allow);
                }
                if let Some(t) = &spec.prompt_title {
                    dv = dv
                        .set_input_title(t)
                        .map_err(|e| PyErr::new::<PyValueError, _>(format!("input_title: {e}")))?;
                }
                if let Some(m) = &spec.prompt {
                    dv = dv.set_input_message(m).map_err(|e| {
                        PyErr::new::<PyValueError, _>(format!("input_message: {e}"))
                    })?;
                }
                if let Some(t) = &spec.error_title {
                    dv = dv
                        .set_error_title(t)
                        .map_err(|e| PyErr::new::<PyValueError, _>(format!("error_title: {e}")))?;
                }
                if let Some(m) = &spec.error {
                    dv = dv.set_error_message(m).map_err(|e| {
                        PyErr::new::<PyValueError, _>(format!("error_message: {e}"))
                    })?;
                }

                ws.add_data_validation(
                    spec.first_row,
                    spec.first_col,
                    spec.last_row,
                    spec.last_col,
                    &dv,
                )
                .map_err(|e| {
                    PyErr::new::<PyIOError, _>(format!("add_data_validation failed: {e}"))
                })?;
            }

            // Notes/comments.
            for note in &state.notes {
                // Excel defaults to prefixing the author name into the note text.
                // Openpyxl's comment reader treats that prefix as part of the text, so
                // disable it to keep verifier semantics stable.
                let mut n = Note::new(&note.text).add_author_prefix(false);
                if let Some(author) = &note.author {
                    n = n.set_author(author);
                }
                ws.insert_note(note.row, note.col, &n)
                    .map_err(|e| PyErr::new::<PyIOError, _>(format!("insert_note failed: {e}")))?;
            }

            // Images.
            for img in &state.images {
                let image = Image::new(&img.path).map_err(|e| {
                    PyErr::new::<PyIOError, _>(format!("Failed to open image: {e}"))
                })?;
                ws.insert_image_with_offset(img.row, img.col, &image, img.x_offset, img.y_offset)
                    .map_err(|e| PyErr::new::<PyIOError, _>(format!("insert_image failed: {e}")))?;
            }

            // Hyperlinks.
            for link in &state.hyperlinks {
                let mut url = if link.internal {
                    Url::new(format!("internal:{}", link.target.trim_start_matches('#')))
                } else {
                    Url::new(&link.target)
                };
                if let Some(text) = &link.display {
                    url = url.set_text(text);
                }
                if let Some(tip) = &link.tooltip {
                    url = url.set_tip(tip);
                }

                ws.write(link.row, link.col, url).map_err(|e| {
                    PyErr::new::<PyIOError, _>(format!("write hyperlink failed: {e}"))
                })?;
            }

            let mut coords: HashSet<(u32, u16)> = HashSet::new();
            coords.extend(state.cells.keys().copied());
            coords.extend(state.formats.keys().copied());
            coords.extend(state.borders.keys().copied());

            let mut coords_vec: Vec<(u32, u16)> = coords.into_iter().collect();
            coords_vec.sort_by(|a, b| (a.0, a.1).cmp(&(b.0, b.1)));

            for (row0, col) in coords_vec {
                if merged_cells.contains(&(row0, col)) {
                    continue;
                }
                let cell = state.cells.get(&(row0, col)).unwrap_or(&CellPayload::Blank);

                let fmt = build_format(
                    cell,
                    state.formats.get(&(row0, col)),
                    state.borders.get(&(row0, col)),
                );

                match cell {
                    CellPayload::Blank => {
                        if let Some(f) = &fmt {
                            ws.write_blank(row0, col, f).map_err(|e| {
                                PyErr::new::<PyIOError, _>(format!("write_blank failed: {e}"))
                            })?;
                        }
                    }
                    CellPayload::String(s) => {
                        if let Some(f) = &fmt {
                            ws.write_string_with_format(row0, col, s, f).map_err(|e| {
                                PyErr::new::<PyIOError, _>(format!("write_string failed: {e}"))
                            })?;
                        } else {
                            ws.write_string(row0, col, s).map_err(|e| {
                                PyErr::new::<PyIOError, _>(format!("write_string failed: {e}"))
                            })?;
                        }
                    }
                    CellPayload::Number(n) => {
                        if let Some(f) = &fmt {
                            ws.write_number_with_format(row0, col, *n, f).map_err(|e| {
                                PyErr::new::<PyIOError, _>(format!("write_number failed: {e}"))
                            })?;
                        } else {
                            ws.write_number(row0, col, *n).map_err(|e| {
                                PyErr::new::<PyIOError, _>(format!("write_number failed: {e}"))
                            })?;
                        }
                    }
                    CellPayload::Boolean(b) => {
                        if let Some(f) = &fmt {
                            ws.write_boolean_with_format(row0, col, *b, f)
                                .map_err(|e| {
                                    PyErr::new::<PyIOError, _>(format!("write_boolean failed: {e}"))
                                })?;
                        } else {
                            ws.write_boolean(row0, col, *b).map_err(|e| {
                                PyErr::new::<PyIOError, _>(format!("write_boolean failed: {e}"))
                            })?;
                        }
                    }
                    CellPayload::Formula(formula) => {
                        if let Some(f) = &fmt {
                            ws.write_formula_with_format(row0, col, formula.as_str(), f)
                                .map_err(|e| {
                                    PyErr::new::<PyIOError, _>(format!("write_formula failed: {e}"))
                                })?;
                        } else {
                            ws.write_formula(row0, col, formula.as_str()).map_err(|e| {
                                PyErr::new::<PyIOError, _>(format!("write_formula failed: {e}"))
                            })?;
                        }
                    }
                    CellPayload::Error(token) => {
                        // Prefer error formulas that OpenpyxlAdapter can recognize.
                        let formula = match token.as_str() {
                            "#DIV/0!" => Some("=1/0"),
                            "#N/A" => Some("=NA()"),
                            "#VALUE!" => Some("=\"text\"+1"),
                            _ => None,
                        };
                        if let Some(formula) = formula {
                            if let Some(f) = &fmt {
                                ws.write_formula_with_format(row0, col, formula, f)
                                    .map_err(|e| {
                                        PyErr::new::<PyIOError, _>(format!(
                                            "write_formula failed: {e}"
                                        ))
                                    })?;
                            } else {
                                ws.write_formula(row0, col, formula).map_err(|e| {
                                    PyErr::new::<PyIOError, _>(format!("write_formula failed: {e}"))
                                })?;
                            }
                        } else if let Some(f) = &fmt {
                            ws.write_string_with_format(row0, col, token, f)
                                .map_err(|e| {
                                    PyErr::new::<PyIOError, _>(format!("write_string failed: {e}"))
                                })?;
                        } else {
                            ws.write_string(row0, col, token).map_err(|e| {
                                PyErr::new::<PyIOError, _>(format!("write_string failed: {e}"))
                            })?;
                        }
                    }
                    CellPayload::Date(d) => {
                        let f = fmt.as_ref().ok_or_else(|| {
                            PyErr::new::<PyValueError, _>("internal: date missing format")
                        })?;
                        ws.write_datetime_with_format(row0, col, *d, f)
                            .map_err(|e| {
                                PyErr::new::<PyIOError, _>(format!("write_datetime failed: {e}"))
                            })?;
                    }
                    CellPayload::DateTime(dt) => {
                        let f = fmt.as_ref().ok_or_else(|| {
                            PyErr::new::<PyValueError, _>("internal: datetime missing format")
                        })?;
                        ws.write_datetime_with_format(row0, col, *dt, f)
                            .map_err(|e| {
                                PyErr::new::<PyIOError, _>(format!("write_datetime failed: {e}"))
                            })?;
                    }
                }
            }

            wb.push_worksheet(ws);
        }

        wb.save(path)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to save workbook: {e}")))
    }
}
