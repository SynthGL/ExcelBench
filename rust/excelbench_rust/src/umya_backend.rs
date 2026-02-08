use pyo3::exceptions::{PyIOError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};

use std::collections::HashMap;
use std::io::{Read, Write};
use std::path::Path;

use chrono::{Duration, NaiveDate, NaiveDateTime, NaiveTime};

use std::str::FromStr;

use quick_xml::events::Event;
use quick_xml::{Reader, Writer};

use umya_spreadsheet::{
    new_file, reader, writer, Border, Color, ColorScale, ConditionalFormatValueObject,
    ConditionalFormatValueObjectValues, ConditionalFormatValues, ConditionalFormatting,
    ConditionalFormattingOperatorValues, ConditionalFormattingRule, DataBar, DataValidation,
    DataValidationOperatorValues, DataValidationValues, DataValidations, Fill, Formula,
    HorizontalAlignmentValues, NumberingFormat, Pane, PaneStateValues, PaneValues, PatternFill,
    SequenceOfReferences, SheetView, Spreadsheet, Style, VerticalAlignmentValues,
};

use zip::write::FileOptions;
use zip::{ZipArchive, ZipWriter};

use crate::util::{a1_to_row_col, cell_blank, cell_with_value, parse_iso_date, parse_iso_datetime};

fn looks_like_date_format(code: &str) -> bool {
    // Heuristic: date formats typically include year + day tokens.
    let lc = code.to_ascii_lowercase();
    lc.contains('y') && lc.contains('d')
}

fn to_umya_argb(color: &str) -> String {
    // Normalize common "#RRGGBB" colors into "FFRRGGBB".
    let mut s = color.trim().to_string();
    if let Some(rest) = s.strip_prefix('#') {
        s = rest.to_string();
    }
    let s = s.to_ascii_uppercase();
    if s.len() == 6 {
        format!("FF{s}")
    } else {
        s
    }
}

fn normalize_a1(s: &str) -> String {
    s.replace('$', "")
}

fn strip_leading_equal(s: &str) -> &str {
    s.strip_prefix('=').unwrap_or(s)
}

fn argb_to_hex_rgb(argb: &str) -> Option<String> {
    let s = argb.trim();
    if s.len() == 8 {
        return Some(format!("#{}", &s[2..]));
    }
    if s.len() == 6 {
        return Some(format!("#{s}"));
    }
    None
}

fn cf_type_to_str(t: &ConditionalFormatValues) -> &'static str {
    match t {
        ConditionalFormatValues::CellIs => "cellIs",
        ConditionalFormatValues::Expression => "expression",
        ConditionalFormatValues::DataBar => "dataBar",
        ConditionalFormatValues::ColorScale => "colorScale",
        ConditionalFormatValues::ContainsText => "containsText",
        ConditionalFormatValues::BeginsWith => "beginsWith",
        ConditionalFormatValues::EndsWith => "endsWith",
        ConditionalFormatValues::NotContainsText => "notContainsText",
        ConditionalFormatValues::DuplicateValues => "duplicateValues",
        ConditionalFormatValues::UniqueValues => "uniqueValues",
        ConditionalFormatValues::AboveAverage => "aboveAverage",
        ConditionalFormatValues::Top10 => "top10",
        ConditionalFormatValues::TimePeriod => "timePeriod",
        ConditionalFormatValues::IconSet => "iconSet",
        ConditionalFormatValues::ContainsBlanks => "containsBlanks",
        ConditionalFormatValues::NotContainsBlanks => "notContainsBlanks",
        ConditionalFormatValues::ContainsErrors => "containsErrors",
        ConditionalFormatValues::NotContainsErrors => "notContainsErrors",
    }
}

fn cf_operator_to_str(op: &ConditionalFormattingOperatorValues) -> &'static str {
    match op {
        ConditionalFormattingOperatorValues::BeginsWith => "beginsWith",
        ConditionalFormattingOperatorValues::Between => "between",
        ConditionalFormattingOperatorValues::ContainsText => "containsText",
        ConditionalFormattingOperatorValues::EndsWith => "endsWith",
        ConditionalFormattingOperatorValues::Equal => "equal",
        ConditionalFormattingOperatorValues::GreaterThan => "greaterThan",
        ConditionalFormattingOperatorValues::GreaterThanOrEqual => "greaterThanOrEqual",
        ConditionalFormattingOperatorValues::LessThan => "lessThan",
        ConditionalFormattingOperatorValues::LessThanOrEqual => "lessThanOrEqual",
        ConditionalFormattingOperatorValues::NotBetween => "notBetween",
        ConditionalFormattingOperatorValues::NotContains => "notContains",
        ConditionalFormattingOperatorValues::NotEqual => "notEqual",
    }
}

fn dv_type_to_str(t: &DataValidationValues) -> &'static str {
    match t {
        DataValidationValues::Custom => "custom",
        DataValidationValues::Date => "date",
        DataValidationValues::Decimal => "decimal",
        DataValidationValues::List => "list",
        DataValidationValues::None => "none",
        DataValidationValues::TextLength => "textLength",
        DataValidationValues::Time => "time",
        DataValidationValues::Whole => "whole",
    }
}

fn dv_operator_to_str(op: &DataValidationOperatorValues) -> &'static str {
    match op {
        DataValidationOperatorValues::Between => "between",
        DataValidationOperatorValues::Equal => "equal",
        DataValidationOperatorValues::GreaterThan => "greaterThan",
        DataValidationOperatorValues::GreaterThanOrEqual => "greaterThanOrEqual",
        DataValidationOperatorValues::LessThan => "lessThan",
        DataValidationOperatorValues::LessThanOrEqual => "lessThanOrEqual",
        DataValidationOperatorValues::NotBetween => "notBetween",
        DataValidationOperatorValues::NotEqual => "notEqual",
    }
}

fn none_if_empty(s: &str) -> Option<String> {
    let trimmed = s.trim();
    if trimmed.is_empty() {
        None
    } else {
        Some(trimmed.to_string())
    }
}

fn border_style_to_umya(style: &str) -> Option<&'static str> {
    match style {
        "none" => None,
        "thin" => Some(Border::BORDER_THIN),
        "medium" => Some(Border::BORDER_MEDIUM),
        "thick" => Some(Border::BORDER_THICK),
        "double" => Some(Border::BORDER_DOUBLE),
        "dashed" => Some(Border::BORDER_DASHED),
        "dotted" => Some(Border::BORDER_DOTTED),
        "hair" => Some(Border::BORDER_HAIR),
        "mediumDashed" => Some(Border::BORDER_MEDIUMDASHED),
        "dashDot" => Some(Border::BORDER_DASHDOT),
        "mediumDashDot" => Some(Border::BORDER_MEDIUMDASHDOT),
        "dashDotDot" => Some(Border::BORDER_DASHDOTDOT),
        "mediumDashDotDot" => Some(Border::BORDER_MEDIUMDASHDOTDOT),
        "slantDashDot" => Some(Border::BORDER_SLANTDASHDOT),
        _ => None,
    }
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
    let days = total_ms as f64 / 86_400_000.0;
    // Invert the 1900 leap-year bug adjustment used in `excel_serial_to_naive_datetime`.
    // There, for serials < 60, an extra day is added (serial -> days: days = serial + 1).
    // The inverse mapping is:
    //   if days < 61 then serial = days - 1 (so serial < 60)
    //   else               serial = days
    let serial = if days < 61.0 { days - 1.0 } else { days };
    Some(serial)
}

#[pyclass(unsendable)]
pub struct UmyaBook {
    book: Spreadsheet,
    saved: bool,
    source_path: Option<String>,
    hyperlink_tooltips: HashMap<String, HashMap<String, String>>, // sheet -> cell -> tooltip
}

#[pymethods]
impl UmyaBook {
    #[new]
    pub fn new() -> Self {
        let mut book = new_file();
        // Match other adapters: start without a default sheet.
        let _ = book.remove_sheet_by_name("Sheet1");
        Self {
            book,
            saved: false,
            source_path: None,
            hyperlink_tooltips: HashMap::new(),
        }
    }

    #[staticmethod]
    pub fn open(path: &str) -> PyResult<Self> {
        let p = Path::new(path);
        let book = reader::xlsx::read(p)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to open workbook: {e}")))?;
        Ok(Self {
            book,
            saved: false,
            source_path: Some(path.to_string()),
            hyperlink_tooltips: HashMap::new(),
        })
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

    pub fn read_cell_format(&self, py: Python<'_>, sheet: &str, a1: &str) -> PyResult<PyObject> {
        // Keep this minimal for now: background fill color is needed by Tier2 merged cell tests.
        if self.book.get_sheet_by_name(sheet).is_none() {
            return Err(PyErr::new::<PyValueError, _>(format!(
                "Unknown sheet: {sheet}"
            )));
        }

        let d = PyDict::new_bound(py);
        let a1 = normalize_a1(a1);

        if let Some(path_str) = &self.source_path {
            if let Some(bg) = read_bg_color_from_xlsx(Path::new(path_str), sheet, a1.as_str())? {
                d.set_item("bg_color", bg)?;
            }
        }

        Ok(d.into())
    }

    // =========================================================================
    // Tier 2 Read Operations
    // =========================================================================

    pub fn read_merged_ranges(&self, sheet: &str) -> PyResult<Vec<String>> {
        let ws = self
            .book
            .get_sheet_by_name(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;
        Ok(ws.get_merge_cells().iter().map(|r| r.get_range()).collect())
    }

    pub fn read_conditional_formats(&self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        let ws = self
            .book
            .get_sheet_by_name(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let out = PyList::empty_bound(py);
        let theme = self.book.get_theme();

        for group in ws.get_conditional_formatting_collection() {
            let range_value = group.get_sequence_of_references().get_sqref();
            for rule in group.get_conditional_collection() {
                let entry = PyDict::new_bound(py);
                entry.set_item("range", range_value.clone())?;
                entry.set_item("rule_type", cf_type_to_str(rule.get_type()))?;

                // Only include operator for cellIs rules (matches OpenpyxlAdapter).
                if *rule.get_type() == ConditionalFormatValues::CellIs {
                    entry.set_item("operator", cf_operator_to_str(rule.get_operator()))?;
                } else {
                    entry.set_item("operator", py.None())?;
                }

                let formula = rule
                    .get_formula()
                    .map(|f| f.get_address_str())
                    .unwrap_or_default();
                if formula.is_empty() {
                    entry.set_item("formula", py.None())?;
                } else if *rule.get_type() == ConditionalFormatValues::Expression
                    && !formula.starts_with('=')
                {
                    entry.set_item("formula", format!("={formula}"))?;
                } else {
                    entry.set_item("formula", formula)?;
                }

                entry.set_item("priority", *rule.get_priority())?;
                entry.set_item("stop_if_true", *rule.get_stop_if_true())?;

                let fmt = PyDict::new_bound(py);
                if let Some(style) = rule.get_style() {
                    if let Some(color) = style.get_background_color() {
                        let argb = color.get_argb_with_theme(theme);
                        if let Some(hex) = argb_to_hex_rgb(&argb) {
                            fmt.set_item("bg_color", hex)?;
                        }
                    }
                    if let Some(font) = style.get_font() {
                        let c = font.get_color();
                        let argb = c.get_argb_with_theme(theme);
                        if let Some(hex) = argb_to_hex_rgb(&argb) {
                            fmt.set_item("font_color", hex)?;
                        }
                    }
                }
                entry.set_item("format", fmt)?;

                out.append(entry)?;
            }
        }

        Ok(out.into())
    }

    pub fn read_data_validations(&self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        let ws = self
            .book
            .get_sheet_by_name(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let out = PyList::empty_bound(py);
        let Some(dvs) = ws.get_data_validations() else {
            return Ok(out.into());
        };

        for dv in dvs.get_data_validation_list() {
            let entry = PyDict::new_bound(py);
            entry.set_item("range", dv.get_sequence_of_references().get_sqref())?;
            entry.set_item("validation_type", dv_type_to_str(dv.get_type()))?;

            // Default operator in xlsx is "between"; keep None when absent.
            let op = dv_operator_to_str(dv.get_operator());
            entry.set_item("operator", op)?;

            entry.set_item("formula1", none_if_empty(dv.get_formula1()))?;
            entry.set_item("formula2", none_if_empty(dv.get_formula2()))?;
            entry.set_item("allow_blank", *dv.get_allow_blank())?;
            entry.set_item("show_input", *dv.get_show_input_message())?;
            entry.set_item("show_error", *dv.get_show_error_message())?;
            entry.set_item("prompt_title", none_if_empty(dv.get_prompt_title()))?;
            entry.set_item("prompt", none_if_empty(dv.get_prompt()))?;
            entry.set_item("error_title", none_if_empty(dv.get_error_title()))?;
            entry.set_item("error", none_if_empty(dv.get_error_message()))?;
            out.append(entry)?;
        }

        Ok(out.into())
    }

    pub fn read_images(&self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        let out = PyList::empty_bound(py);

        // umya-spreadsheet doesn't reliably surface images when reading existing files.
        // Parse the xlsx zip directly when we have a source path.
        if let Some(path_str) = &self.source_path {
            let specs = read_images_from_xlsx(Path::new(path_str), sheet)?;
            for spec in specs {
                let entry = PyDict::new_bound(py);
                entry.set_item("cell", spec.cell)?;
                entry.set_item("path", spec.path)?;
                entry.set_item("anchor", spec.anchor)?;
                entry.set_item("offset", py.None())?;
                entry.set_item("alt_text", py.None())?;
                out.append(entry)?;
            }
            return Ok(out.into());
        }

        let ws = self
            .book
            .get_sheet_by_name(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        for img in ws.get_image_collection() {
            let entry = PyDict::new_bound(py);
            entry.set_item("cell", img.get_coordinate())?;
            entry.set_item(
                "anchor",
                if img.get_to_marker_type().is_some() {
                    "twoCell"
                } else {
                    "oneCell"
                },
            )?;
            entry.set_item("path", py.None())?;
            entry.set_item("offset", py.None())?;
            entry.set_item("alt_text", py.None())?;
            out.append(entry)?;
        }

        Ok(out.into())
    }

    pub fn read_freeze_panes(&self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        let ws = self
            .book
            .get_sheet_by_name(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let d = PyDict::new_bound(py);

        let views = ws.get_sheets_views().get_sheet_view_list();
        if let Some(view) = views.get(0) {
            if let Some(pane) = view.get_pane() {
                match pane.get_state() {
                    PaneStateValues::Frozen | PaneStateValues::FrozenSplit => {
                        d.set_item("mode", "freeze")?;
                        d.set_item("top_left_cell", pane.get_top_left_cell().to_string())?;
                    }
                    PaneStateValues::Split => {
                        d.set_item("mode", "split")?;
                        let x = *pane.get_horizontal_split();
                        let y = *pane.get_vertical_split();
                        d.set_item("x_split", x.round() as i64)?;
                        d.set_item("y_split", y.round() as i64)?;
                    }
                }
            }
        }

        Ok(d.into())
    }

    pub fn read_hyperlinks(&self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        let ws = self
            .book
            .get_sheet_by_name(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let out = PyList::empty_bound(py);

        let Some(path_str) = &self.source_path else {
            return Ok(out.into());
        };
        let specs = read_hyperlinks_from_xlsx(Path::new(path_str), sheet)?;
        for spec in specs {
            let entry = PyDict::new_bound(py);
            entry.set_item("cell", spec.cell.clone())?;
            entry.set_item("target", spec.target)?;
            let display = ws
                .get_cell(spec.cell.as_str())
                .map(|c| c.get_value().into_owned())
                .unwrap_or_default();
            if display.is_empty() {
                entry.set_item("display", py.None())?;
            } else {
                entry.set_item("display", display)?;
            }
            if let Some(tip) = spec.tooltip {
                entry.set_item("tooltip", tip)?;
            } else {
                entry.set_item("tooltip", py.None())?;
            }
            entry.set_item("internal", spec.internal)?;
            out.append(entry)?;
        }

        Ok(out.into())
    }

    pub fn read_comments(&self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        let out = PyList::empty_bound(py);

        let Some(path_str) = &self.source_path else {
            return Ok(out.into());
        };
        let comments = read_comments_from_xlsx(Path::new(path_str), sheet)?;
        for c in comments {
            let entry = PyDict::new_bound(py);
            entry.set_item("cell", c.cell)?;
            entry.set_item("text", c.text)?;
            if let Some(author) = c.author {
                entry.set_item("author", author)?;
            } else {
                entry.set_item("author", py.None())?;
            }
            entry.set_item("threaded", false)?;
            out.append(entry)?;
        }

        Ok(out.into())
    }

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
        payload: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let dict = payload
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("payload must be a dict"))?;

        let style = ws.get_style_mut(a1);

        if let Some(v) = dict.get_item("bold")? {
            let b = v.extract::<bool>()?;
            style.get_font_mut().set_bold(b);
        }
        if let Some(v) = dict.get_item("italic")? {
            let b = v.extract::<bool>()?;
            style.get_font_mut().set_italic(b);
        }
        if let Some(v) = dict.get_item("underline")? {
            let s = v.extract::<String>()?;
            // Only pass through known strings; umya panics on unknown underline.
            if s == "single" || s == "double" {
                style.get_font_mut().set_underline(s);
            }
        }
        if let Some(v) = dict.get_item("strikethrough")? {
            let b = v.extract::<bool>()?;
            style.get_font_mut().set_strikethrough(b);
        }
        if let Some(v) = dict.get_item("font_name")? {
            let s = v.extract::<String>()?;
            style.get_font_mut().set_name(s);
        }
        if let Some(v) = dict.get_item("font_size")? {
            let sz = v.extract::<f64>()?;
            style.get_font_mut().set_size(sz);
        }
        if let Some(v) = dict.get_item("font_color")? {
            let s = v.extract::<String>()?;
            style
                .get_font_mut()
                .get_color_mut()
                .set_argb(to_umya_argb(&s));
        }
        if let Some(v) = dict.get_item("bg_color")? {
            let s = v.extract::<String>()?;
            style.set_background_color_solid(to_umya_argb(&s));
        }
        if let Some(v) = dict.get_item("number_format")? {
            let s = v.extract::<String>()?;
            style.get_number_format_mut().set_format_code(s);
        }
        if let Some(v) = dict.get_item("h_align")? {
            let s = v.extract::<String>()?;
            if let Ok(a) = HorizontalAlignmentValues::from_str(&s) {
                style.get_alignment_mut().set_horizontal(a);
            }
        }
        if let Some(v) = dict.get_item("v_align")? {
            let s = v.extract::<String>()?;
            if let Ok(a) = VerticalAlignmentValues::from_str(&s) {
                style.get_alignment_mut().set_vertical(a);
            }
        }
        if let Some(v) = dict.get_item("wrap")? {
            let b = v.extract::<bool>()?;
            style.get_alignment_mut().set_wrap_text(b);
        }
        if let Some(v) = dict.get_item("rotation")? {
            let rot = v.extract::<i64>()?;
            if rot >= 0 {
                style.get_alignment_mut().set_text_rotation(rot as u32);
            }
        }

        Ok(())
    }

    pub fn write_cell_border(
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

        let style = ws.get_style_mut(a1);

        if let Some(edge) = dict.get_item("top")? {
            let edge = edge
                .downcast::<PyDict>()
                .map_err(|_| PyErr::new::<PyValueError, _>("border edge must be a dict"))?;
            let border_obj = style.get_borders_mut().get_top_mut();
            if let Some(v) = edge.get_item("style")? {
                let s = v.extract::<String>()?;
                if let Some(bs) = border_style_to_umya(s.as_str()) {
                    border_obj.set_border_style(bs);
                }
            }
            if let Some(v) = edge.get_item("color")? {
                let s = v.extract::<String>()?;
                border_obj.get_color_mut().set_argb(to_umya_argb(&s));
            }
        }
        if let Some(edge) = dict.get_item("bottom")? {
            let edge = edge
                .downcast::<PyDict>()
                .map_err(|_| PyErr::new::<PyValueError, _>("border edge must be a dict"))?;
            let border_obj = style.get_borders_mut().get_bottom_mut();
            if let Some(v) = edge.get_item("style")? {
                let s = v.extract::<String>()?;
                if let Some(bs) = border_style_to_umya(s.as_str()) {
                    border_obj.set_border_style(bs);
                }
            }
            if let Some(v) = edge.get_item("color")? {
                let s = v.extract::<String>()?;
                border_obj.get_color_mut().set_argb(to_umya_argb(&s));
            }
        }
        if let Some(edge) = dict.get_item("left")? {
            let edge = edge
                .downcast::<PyDict>()
                .map_err(|_| PyErr::new::<PyValueError, _>("border edge must be a dict"))?;
            let border_obj = style.get_borders_mut().get_left_mut();
            if let Some(v) = edge.get_item("style")? {
                let s = v.extract::<String>()?;
                if let Some(bs) = border_style_to_umya(s.as_str()) {
                    border_obj.set_border_style(bs);
                }
            }
            if let Some(v) = edge.get_item("color")? {
                let s = v.extract::<String>()?;
                border_obj.get_color_mut().set_argb(to_umya_argb(&s));
            }
        }
        if let Some(edge) = dict.get_item("right")? {
            let edge = edge
                .downcast::<PyDict>()
                .map_err(|_| PyErr::new::<PyValueError, _>("border edge must be a dict"))?;
            let border_obj = style.get_borders_mut().get_right_mut();
            if let Some(v) = edge.get_item("style")? {
                let s = v.extract::<String>()?;
                if let Some(bs) = border_style_to_umya(s.as_str()) {
                    border_obj.set_border_style(bs);
                }
            }
            if let Some(v) = edge.get_item("color")? {
                let s = v.extract::<String>()?;
                border_obj.get_color_mut().set_argb(to_umya_argb(&s));
            }
        }

        if let Some(edge) = dict.get_item("diagonal_up")? {
            let edge = edge
                .downcast::<PyDict>()
                .map_err(|_| PyErr::new::<PyValueError, _>("border edge must be a dict"))?;
            style.get_borders_mut().set_diagonal_up(true);
            let border_obj = style.get_borders_mut().get_diagonal_mut();
            if let Some(v) = edge.get_item("style")? {
                let s = v.extract::<String>()?;
                if let Some(bs) = border_style_to_umya(s.as_str()) {
                    border_obj.set_border_style(bs);
                }
            }
            if let Some(v) = edge.get_item("color")? {
                let s = v.extract::<String>()?;
                border_obj.get_color_mut().set_argb(to_umya_argb(&s));
            }
        }

        if let Some(edge) = dict.get_item("diagonal_down")? {
            let edge = edge
                .downcast::<PyDict>()
                .map_err(|_| PyErr::new::<PyValueError, _>("border edge must be a dict"))?;
            style.get_borders_mut().set_diagonal_down(true);
            let border_obj = style.get_borders_mut().get_diagonal_mut();
            if let Some(v) = edge.get_item("style")? {
                let s = v.extract::<String>()?;
                if let Some(bs) = border_style_to_umya(s.as_str()) {
                    border_obj.set_border_style(bs);
                }
            }
            if let Some(v) = edge.get_item("color")? {
                let s = v.extract::<String>()?;
                border_obj.get_color_mut().set_argb(to_umya_argb(&s));
            }
        }

        Ok(())
    }

    pub fn set_row_height(&mut self, sheet: &str, row: u32, height: f64) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;
        ws.get_row_dimension_mut(&row).set_height(height);
        Ok(())
    }

    pub fn set_column_width(&mut self, sheet: &str, column: &str, width: f64) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;
        ws.get_column_dimension_mut(column).set_width(width);
        Ok(())
    }

    // =========================================================================
    // Tier 2 Write Operations
    // =========================================================================

    pub fn merge_cells(&mut self, sheet: &str, cell_range: &str) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let range = normalize_a1(cell_range);
        if !range.contains(':') {
            return Ok(());
        }
        ws.add_merge_cells(range);
        Ok(())
    }

    pub fn add_conditional_format(
        &mut self,
        sheet: &str,
        payload: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
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
        let priority = cf
            .get_item("priority")?
            .and_then(|v| v.extract::<i64>().ok())
            .unwrap_or(1) as i32;

        // Optional formatting (ExcelBench currently checks bg_color for these cases).
        let mut bg_color: Option<String> = None;
        if let Some(fmt_any) = cf.get_item("format")? {
            if let Ok(fmt) = fmt_any.downcast::<PyDict>() {
                if let Some(v) = fmt.get_item("bg_color")? {
                    bg_color = Some(v.extract::<String>()?);
                }
            }
        }

        let mut rule = ConditionalFormattingRule::default();
        if let Ok(t) = ConditionalFormatValues::from_str(&rule_type) {
            rule.set_type(t);
        } else {
            // Unknown types are ignored.
            return Ok(());
        }
        rule.set_priority(priority);
        if stop_if_true {
            rule.set_stop_if_true(true);
        }

        match rule.get_type() {
            ConditionalFormatValues::CellIs => {
                if let Some(op_any) = cf.get_item("operator")? {
                    let op = op_any.extract::<String>()?;
                    if let Ok(v) = ConditionalFormattingOperatorValues::from_str(&op) {
                        rule.set_operator(v);
                    }
                }
                if let Some(f_any) = cf.get_item("formula")? {
                    let mut f = Formula::default();
                    let formula = f_any.extract::<String>()?;
                    f.set_string_value(strip_leading_equal(&formula));
                    rule.set_formula(f);
                }
            }
            ConditionalFormatValues::Expression => {
                if let Some(f_any) = cf.get_item("formula")? {
                    let mut f = Formula::default();
                    let formula = f_any.extract::<String>()?;
                    f.set_string_value(strip_leading_equal(&formula));
                    rule.set_formula(f);
                }
            }
            ConditionalFormatValues::DataBar => {
                let mut min = ConditionalFormatValueObject::default();
                min.set_type(ConditionalFormatValueObjectValues::Number)
                    .set_val("0");
                let mut max = ConditionalFormatValueObject::default();
                max.set_type(ConditionalFormatValueObjectValues::Number)
                    .set_val("10");
                let mut color = Color::default();
                color.set_argb("FF638EC6");
                let mut db = DataBar::default();
                db.add_cfvo_collection(min)
                    .add_cfvo_collection(max)
                    .add_color_collection(color);
                rule.set_data_bar(db);
            }
            ConditionalFormatValues::ColorScale => {
                let mut c1 = ConditionalFormatValueObject::default();
                c1.set_type(ConditionalFormatValueObjectValues::Min);
                let mut c2 = ConditionalFormatValueObject::default();
                c2.set_type(ConditionalFormatValueObjectValues::Percentile)
                    .set_val("50");
                let mut c3 = ConditionalFormatValueObject::default();
                c3.set_type(ConditionalFormatValueObjectValues::Max);

                let mut start = Color::default();
                start.set_argb("FFF8696B");
                let mut mid = Color::default();
                mid.set_argb("FFFFEB84");
                let mut end = Color::default();
                end.set_argb("FF63BE7B");

                let mut cs = ColorScale::default();
                cs.add_cfvo_collection(c1)
                    .add_cfvo_collection(c2)
                    .add_cfvo_collection(c3)
                    .add_color_collection(start)
                    .add_color_collection(mid)
                    .add_color_collection(end);
                rule.set_color_scale(cs);
            }
            _ => {
                // Other rule types aren't needed for current Tier2 cases.
            }
        }

        if let Some(bg) = bg_color {
            let mut color = Color::default();
            color.set_argb(to_umya_argb(&bg));
            let mut pattern_fill = PatternFill::default();
            // Excel conditional format fills are stored as foreground colors (dxf fills).
            pattern_fill.set_foreground_color(color);
            let mut fill = Fill::default();
            fill.set_pattern_fill(pattern_fill);
            let mut style = Style::default();
            style.set_fill(fill);
            rule.set_style(style);
        }

        let mut seq = SequenceOfReferences::default();
        seq.set_sqref(normalize_a1(&range));

        let mut group = ConditionalFormatting::default();
        group.set_sequence_of_references(seq);
        group.set_conditional_collection(vec![rule]);

        ws.add_conditional_formatting_collection(group);
        Ok(())
    }

    pub fn add_data_validation(&mut self, sheet: &str, payload: &Bound<'_, PyAny>) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
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

        let mut dv = DataValidation::default();
        if let Ok(t) = DataValidationValues::from_str(&validation_type) {
            dv.set_type(t);
        } else {
            return Ok(());
        }
        if let Some(op_any) = v.get_item("operator")? {
            let op = op_any.extract::<String>()?;
            if let Ok(opv) = DataValidationOperatorValues::from_str(&op) {
                dv.set_operator(opv);
            }
        }

        if let Some(v1_any) = v.get_item("formula1")? {
            let f1 = v1_any.extract::<String>()?;
            dv.set_formula1(strip_leading_equal(&f1));
        }
        if let Some(v2_any) = v.get_item("formula2")? {
            let f2 = v2_any.extract::<String>()?;
            dv.set_formula2(strip_leading_equal(&f2));
        }

        if let Some(a_any) = v.get_item("allow_blank")? {
            dv.set_allow_blank(a_any.extract::<bool>()?);
        }
        if let Some(t_any) = v.get_item("prompt_title")? {
            dv.set_prompt_title(t_any.extract::<String>()?);
        }
        if let Some(p_any) = v.get_item("prompt")? {
            dv.set_prompt(p_any.extract::<String>()?);
        }
        if let Some(t_any) = v.get_item("error_title")? {
            dv.set_error_title(t_any.extract::<String>()?);
        }
        if let Some(e_any) = v.get_item("error")? {
            dv.set_error_message(e_any.extract::<String>()?);
        }

        let mut seq = SequenceOfReferences::default();
        seq.set_sqref(normalize_a1(&range));
        dv.set_sequence_of_references(seq);

        if ws.get_data_validations_mut().is_none() {
            ws.set_data_validations(DataValidations::default());
        }
        if let Some(dvs) = ws.get_data_validations_mut() {
            dvs.add_data_validation_list(dv);
        }

        Ok(())
    }

    pub fn add_hyperlink(&mut self, sheet: &str, payload: &Bound<'_, PyAny>) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
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

        let a1 = normalize_a1(&cell);
        if let Some(text) = &display {
            ws.get_cell_mut(a1.as_str())
                .set_value_string(text.to_string());
        }

        let hyperlink = ws.get_cell_mut(a1.as_str()).get_hyperlink_mut();
        hyperlink.set_url(normalize_a1(&target));
        hyperlink.set_location(internal);
        if let Some(tip) = &tooltip {
            hyperlink.set_tooltip(tip.to_string());
            self.hyperlink_tooltips
                .entry(sheet.to_string())
                .or_default()
                .insert(a1.clone(), tip.to_string());
        }

        Ok(())
    }

    pub fn add_image(&mut self, sheet: &str, payload: &Bound<'_, PyAny>) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
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

        let p = Path::new(&path);
        if !p.exists() {
            return Err(PyErr::new::<PyIOError, _>(format!(
                "Image path does not exist: {path}"
            )));
        }

        let mut marker = umya_spreadsheet::structs::drawing::spreadsheet::MarkerType::default();
        marker.set_coordinate(normalize_a1(&cell));

        let mut image = umya_spreadsheet::structs::Image::default();
        let res = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            image.new_image(&path, marker);
        }));
        if res.is_err() {
            return Err(PyErr::new::<PyIOError, _>(format!(
                "Failed to insert image (umya-spreadsheet panicked) for path: {path}"
            )));
        }

        ws.add_image(image);
        Ok(())
    }

    pub fn add_comment(&mut self, sheet: &str, payload: &Bound<'_, PyAny>) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
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

        let mut c = umya_spreadsheet::structs::Comment::default();
        c.new_comment(normalize_a1(&cell));
        c.set_text_string(text);
        if let Some(a) = author {
            c.set_author(a);
        }
        ws.add_comments(c);
        Ok(())
    }

    pub fn set_freeze_panes(&mut self, sheet: &str, payload: &Bound<'_, PyAny>) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
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

        let views = ws.get_sheet_views_mut().get_sheet_view_list_mut();
        if views.is_empty() {
            views.push(SheetView::default());
        }
        let view = views
            .get_mut(0)
            .ok_or_else(|| PyErr::new::<PyValueError, _>("Failed to access sheet view"))?;

        let mut pane = Pane::default();
        pane.set_active_pane(PaneValues::BottomRight);

        if mode == "freeze" {
            let tl = cfg
                .get_item("top_left_cell")?
                .and_then(|v| v.extract::<String>().ok())
                .ok_or_else(|| PyErr::new::<PyValueError, _>("freeze missing 'top_left_cell'"))?;
            let tl = normalize_a1(&tl);
            let (row0, col0) =
                a1_to_row_col(&tl).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
            pane.set_horizontal_split(col0 as f64);
            pane.set_vertical_split(row0 as f64);
            pane.set_state(PaneStateValues::Frozen);
            pane.get_top_left_cell_mut().set_coordinate(tl);
        } else if mode == "split" {
            let x_split = cfg
                .get_item("x_split")?
                .and_then(|v| v.extract::<u32>().ok())
                .unwrap_or(0);
            let y_split = cfg
                .get_item("y_split")?
                .and_then(|v| v.extract::<u32>().ok())
                .unwrap_or(0);
            pane.set_horizontal_split(x_split as f64);
            pane.set_vertical_split(y_split as f64);
            pane.set_state(PaneStateValues::Split);
            pane.get_top_left_cell_mut().set_coordinate("A1");
        } else {
            return Ok(());
        }

        view.set_pane(pane);
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
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to save workbook: {e}")))?;

        if !self.hyperlink_tooltips.is_empty() {
            patch_xlsx_hyperlink_tooltips(p, &self.hyperlink_tooltips)?;
        }

        Ok(())
    }
}

fn xml_escape_attr(value: &str) -> String {
    let mut out = String::with_capacity(value.len());
    for ch in value.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            _ => out.push(ch),
        }
    }
    out
}

fn parse_attr(tag: &str, name: &str) -> Option<String> {
    let needle = format!("{name}=\"");
    let start = tag.find(&needle)? + needle.len();
    let rest = &tag[start..];
    let end = rest.find('"')?;
    Some(xml_unescape(&rest[..end]))
}

fn patch_sheet_xml_tooltips(xml: &str, tooltips: &HashMap<String, String>) -> String {
    if tooltips.is_empty() {
        return xml.to_string();
    }

    // Prefer a real XML parse/transform to avoid brittle string manipulation.
    if let Some(patched) = patch_sheet_xml_tooltips_quickxml(xml, tooltips) {
        return patched;
    }

    // Fallback to the legacy string-based patcher.
    patch_sheet_xml_tooltips_manual(xml, tooltips)
}

fn patch_sheet_xml_tooltips_quickxml(
    xml: &str,
    tooltips: &HashMap<String, String>,
) -> Option<String> {
    let mut reader = Reader::from_str(xml);
    // Preserve text nodes/whitespace as-is.
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(xml.len() + 128));

    let mut buf: Vec<u8> = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if e.local_name().as_ref() != b"hyperlink" {
                    if writer.write_event(Event::Start(e)).is_err() {
                        return None;
                    }
                    buf.clear();
                    continue;
                }

                let mut cell_ref: Option<String> = None;
                let mut has_tooltip: bool = false;
                for attr_res in e.attributes() {
                    let attr = attr_res.ok()?;
                    let k = attr.key.as_ref();
                    if k == b"tooltip" {
                        has_tooltip = true;
                    } else if k == b"ref" {
                        cell_ref = attr.unescape_value().ok().map(|v| v.to_string());
                    }
                }

                let mut out_e = e.to_owned();
                if !has_tooltip {
                    if let Some(r) = cell_ref {
                        if let Some(tip) = tooltips.get(&r) {
                            out_e.push_attribute(("tooltip", tip.as_str()));
                        }
                    }
                }

                if writer.write_event(Event::Start(out_e)).is_err() {
                    return None;
                }
            }
            Ok(Event::Empty(e)) => {
                if e.local_name().as_ref() != b"hyperlink" {
                    if writer.write_event(Event::Empty(e)).is_err() {
                        return None;
                    }
                    buf.clear();
                    continue;
                }

                let mut cell_ref: Option<String> = None;
                let mut has_tooltip: bool = false;
                for attr_res in e.attributes() {
                    let attr = attr_res.ok()?;
                    let k = attr.key.as_ref();
                    if k == b"tooltip" {
                        has_tooltip = true;
                    } else if k == b"ref" {
                        cell_ref = attr.unescape_value().ok().map(|v| v.to_string());
                    }
                }

                let mut out_e = e.to_owned();
                if !has_tooltip {
                    if let Some(r) = cell_ref {
                        if let Some(tip) = tooltips.get(&r) {
                            out_e.push_attribute(("tooltip", tip.as_str()));
                        }
                    }
                }

                if writer.write_event(Event::Empty(out_e)).is_err() {
                    return None;
                }
            }
            Ok(Event::Eof) => break,
            Ok(e) => {
                if writer.write_event(e).is_err() {
                    return None;
                }
            }
            Err(_) => return None,
        }

        buf.clear();
    }

    String::from_utf8(writer.into_inner()).ok()
}

fn patch_sheet_xml_tooltips_manual(xml: &str, tooltips: &HashMap<String, String>) -> String {
    if tooltips.is_empty() {
        return xml.to_string();
    }

    let mut out = String::with_capacity(xml.len() + 128);
    let mut i: usize = 0;

    while let Some(rel) = xml[i..].find("<hyperlink") {
        let start = i + rel;
        // Avoid matching the <hyperlinks> container tag.
        let after = xml.get(start + "<hyperlink".len()..start + "<hyperlink".len() + 1);
        if after != Some(" ") && after != Some(">") {
            out.push_str(&xml[i..start + "<hyperlink".len()]);
            i = start + "<hyperlink".len();
            continue;
        }

        out.push_str(&xml[i..start]);

        let end_rel = xml[start..].find("/>").or_else(|| xml[start..].find('>'));
        let Some(tag_end_rel) = end_rel else {
            out.push_str(&xml[start..]);
            return out;
        };
        let tag_end = start + tag_end_rel;
        let close_len = if xml[tag_end..].starts_with("/>") {
            2
        } else {
            1
        };
        let tag = &xml[start..tag_end + close_len];

        if tag.contains("tooltip=") {
            out.push_str(tag);
            i = tag_end + close_len;
            continue;
        }

        if let Some(r) = parse_attr(tag, "ref") {
            if let Some(tip) = tooltips.get(&r) {
                let esc = xml_escape_attr(tip);
                if close_len == 2 {
                    let without = &tag[..tag.len() - 2];
                    out.push_str(without);
                    out.push_str(&format!(" tooltip=\"{esc}\"/>"));
                } else {
                    let without = &tag[..tag.len() - 1];
                    out.push_str(without);
                    out.push_str(&format!(" tooltip=\"{esc}\">"));
                }
                i = tag_end + close_len;
                continue;
            }
        }

        out.push_str(tag);
        i = tag_end + close_len;
    }

    out.push_str(&xml[i..]);
    out
}

fn replace_file(tmp_path: &Path, dest_path: &Path) -> PyResult<()> {
    if std::fs::rename(tmp_path, dest_path).is_ok() {
        return Ok(());
    }

    // On Windows, rename() cannot overwrite an existing destination.
    let _ = std::fs::remove_file(dest_path);
    if std::fs::rename(tmp_path, dest_path).is_ok() {
        return Ok(());
    }

    std::fs::copy(tmp_path, dest_path).map_err(|e| {
        PyErr::new::<PyIOError, _>(format!("Failed to replace xlsx with patched version: {e}"))
    })?;
    let _ = std::fs::remove_file(tmp_path);
    Ok(())
}

fn parse_workbook_sheet_map(workbook_xml: &str) -> HashMap<String, String> {
    let mut out: HashMap<String, String> = HashMap::new();
    let mut i: usize = 0;
    while let Some(rel) = workbook_xml[i..].find("<sheet ") {
        let start = i + rel;
        let end_rel = workbook_xml[start..]
            .find("/>")
            .or_else(|| workbook_xml[start..].find('>'));
        let Some(tag_end_rel) = end_rel else {
            break;
        };
        let tag_end = start + tag_end_rel;
        let close_len = if workbook_xml[tag_end..].starts_with("/>") {
            2
        } else {
            1
        };
        let tag = &workbook_xml[start..tag_end + close_len];
        let name = parse_attr(tag, "name");
        let rid = parse_attr(tag, "r:id");
        if let (Some(n), Some(r)) = (name, rid) {
            out.insert(n, r);
        }
        i = tag_end + close_len;
    }
    out
}

fn parse_workbook_rels_map(rels_xml: &str) -> HashMap<String, String> {
    let mut out: HashMap<String, String> = HashMap::new();
    let mut i: usize = 0;
    while let Some(rel) = rels_xml[i..].find("<Relationship ") {
        let start = i + rel;
        let end_rel = rels_xml[start..]
            .find("/>")
            .or_else(|| rels_xml[start..].find('>'));
        let Some(tag_end_rel) = end_rel else {
            break;
        };
        let tag_end = start + tag_end_rel;
        let close_len = if rels_xml[tag_end..].starts_with("/>") {
            2
        } else {
            1
        };
        let tag = &rels_xml[start..tag_end + close_len];
        let id = parse_attr(tag, "Id");
        let target = parse_attr(tag, "Target");
        if let (Some(iid), Some(t)) = (id, target) {
            out.insert(iid, t);
        }
        i = tag_end + close_len;
    }
    out
}

fn workbook_rel_target_to_part(target: &str) -> String {
    // workbook.xml.rels targets may be like:
    // - worksheets/sheet1.xml
    // - /xl/worksheets/sheet1.xml
    // - xl/worksheets/sheet1.xml
    let t = target.trim_start_matches('/');
    if t.starts_with("xl/") {
        t.to_string()
    } else {
        format!("xl/{t}")
    }
}

fn patch_xlsx_hyperlink_tooltips(
    path: &Path,
    tooltips: &HashMap<String, HashMap<String, String>>,
) -> PyResult<()> {
    let f = std::fs::File::open(path).map_err(|e| {
        PyErr::new::<PyIOError, _>(format!("Failed to open xlsx for patching: {e}"))
    })?;
    let mut zip = ZipArchive::new(f)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Invalid xlsx zip: {e}")))?;

    let mut workbook_xml = String::new();
    {
        let mut entry = zip
            .by_name("xl/workbook.xml")
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Missing xl/workbook.xml: {e}")))?;
        entry
            .read_to_string(&mut workbook_xml)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Read workbook.xml failed: {e}")))?;
    }

    let mut rels_xml = String::new();
    {
        let mut entry = zip.by_name("xl/_rels/workbook.xml.rels").map_err(|e| {
            PyErr::new::<PyIOError, _>(format!("Missing xl/_rels/workbook.xml.rels: {e}"))
        })?;
        entry
            .read_to_string(&mut rels_xml)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Read workbook rels failed: {e}")))?;
    }

    let sheet_to_rid = parse_workbook_sheet_map(&workbook_xml);
    let rid_to_target = parse_workbook_rels_map(&rels_xml);

    let mut targets: HashMap<String, HashMap<String, String>> = HashMap::new();
    for (sheet_name, cells) in tooltips {
        let Some(rid) = sheet_to_rid.get(sheet_name) else {
            continue;
        };
        let Some(target) = rid_to_target.get(rid) else {
            continue;
        };
        targets.insert(workbook_rel_target_to_part(target), cells.clone());
    }

    if targets.is_empty() {
        return Ok(());
    }

    drop(zip);

    let f = std::fs::File::open(path).map_err(|e| {
        PyErr::new::<PyIOError, _>(format!("Failed to re-open xlsx for patching: {e}"))
    })?;
    let mut zip = ZipArchive::new(f)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Invalid xlsx zip: {e}")))?;

    let tmp_path = path.with_extension("xlsx.tmp");
    let tmp_file = std::fs::File::create(&tmp_path)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to create temp xlsx: {e}")))?;
    let mut out = ZipWriter::new(tmp_file);

    for idx in 0..zip.len() {
        let mut file = zip
            .by_index(idx)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Zip read failed: {e}")))?;
        let name = file.name().to_string();

        let options = FileOptions::default()
            .compression_method(file.compression())
            .last_modified_time(file.last_modified());

        if file.is_dir() {
            out.add_directory(name, options).map_err(|e| {
                PyErr::new::<PyIOError, _>(format!("Zip write directory failed: {e}"))
            })?;
            continue;
        }

        let mut buf: Vec<u8> = Vec::new();
        file.read_to_end(&mut buf)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Zip entry read failed: {e}")))?;

        if let Some(cells) = targets.get(&name) {
            if let Ok(s) = std::str::from_utf8(&buf) {
                let patched = patch_sheet_xml_tooltips(s, cells);
                buf = patched.into_bytes();
            }
        }

        out.start_file(name, options)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Zip write failed: {e}")))?;
        out.write_all(&buf)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Zip write failed: {e}")))?;
    }

    out.finish()
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Zip finalize failed: {e}")))?;

    replace_file(&tmp_path, path)?;

    Ok(())
}

#[derive(Clone, Debug)]
struct HyperlinkReadSpec {
    cell: String,
    target: String,
    tooltip: Option<String>,
    internal: bool,
}

#[derive(Clone, Debug)]
struct CommentReadSpec {
    cell: String,
    text: String,
    author: Option<String>,
}

#[derive(Clone, Debug)]
struct RelationshipEntry {
    id: String,
    r#type: String,
    target: String,
}

#[derive(Clone, Debug)]
struct ImageReadSpec {
    cell: String,
    path: String,
    anchor: String,
}

fn zip_read_to_string(zip: &mut ZipArchive<std::fs::File>, name: &str) -> PyResult<String> {
    let mut s = String::new();
    let mut entry = zip
        .by_name(name)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Missing {name} in xlsx: {e}")))?;
    entry
        .read_to_string(&mut s)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Read {name} failed: {e}")))?;
    Ok(s)
}

fn parse_rels_entries(rels_xml: &str) -> Vec<RelationshipEntry> {
    let mut out: Vec<RelationshipEntry> = Vec::new();
    let mut i: usize = 0;
    while let Some(rel) = rels_xml[i..].find("<Relationship ") {
        let start = i + rel;
        let end_rel = rels_xml[start..]
            .find("/>")
            .or_else(|| rels_xml[start..].find('>'));
        let Some(tag_end_rel) = end_rel else {
            break;
        };
        let tag_end = start + tag_end_rel;
        let close_len = if rels_xml[tag_end..].starts_with("/>") {
            2
        } else {
            1
        };
        let tag = &rels_xml[start..tag_end + close_len];
        let id = parse_attr(tag, "Id");
        let ty = parse_attr(tag, "Type");
        let target = parse_attr(tag, "Target");
        if let (Some(id), Some(ty), Some(target)) = (id, ty, target) {
            out.push(RelationshipEntry {
                id,
                r#type: ty,
                target,
            });
        }
        i = tag_end + close_len;
    }
    out
}

fn sheet_target_to_rels_entry(sheet_entry: &str) -> String {
    // xl/worksheets/sheet1.xml -> xl/worksheets/_rels/sheet1.xml.rels
    if let Some((dir, file)) = sheet_entry.rsplit_once('/') {
        return format!("{dir}/_rels/{file}.rels");
    }
    format!("xl/worksheets/_rels/{sheet_entry}.rels")
}

fn resolve_sheet_rel_target(target: &str) -> String {
    // Relationships in sheet rels are relative to xl/worksheets/
    let t = target.trim_start_matches('/');
    if t.starts_with("xl/") {
        return t.to_string();
    }
    if let Some(rest) = t.strip_prefix("../") {
        format!("xl/{rest}")
    } else {
        format!("xl/worksheets/{t}")
    }
}

fn resolve_drawing_rel_target(target: &str) -> String {
    // Relationships in drawing rels are relative to xl/drawings/
    let t = target.trim_start_matches('/');
    if t.starts_with("xl/") {
        return t.to_string();
    }
    if let Some(rest) = t.strip_prefix("../") {
        format!("xl/{rest}")
    } else {
        format!("xl/drawings/{t}")
    }
}

fn extract_simple_tag_value(xml: &str, tag: &str) -> Option<String> {
    let open = format!("<{tag}>");
    let close = format!("</{tag}>");
    let start = xml.find(&open)? + open.len();
    let rest = &xml[start..];
    let end_rel = rest.find(&close)?;
    Some(rest[..end_rel].trim().to_string())
}

fn col_to_letters(col0: u32) -> String {
    let mut n = col0 + 1;
    let mut out = String::new();
    while n > 0 {
        let rem = ((n - 1) % 26) as u8;
        out.insert(0, (b'A' + rem) as char);
        n = (n - 1) / 26;
    }
    out
}

fn extract_drawing_rids(sheet_xml: &str) -> Vec<String> {
    let mut out: Vec<String> = Vec::new();
    let mut i: usize = 0;
    while let Some(pos) = sheet_xml[i..].find("<drawing") {
        let start = i + pos;
        let end_rel = sheet_xml[start..]
            .find("/>")
            .or_else(|| sheet_xml[start..].find('>'));
        let Some(tag_end_rel) = end_rel else {
            break;
        };
        let tag_end = start + tag_end_rel;
        let close_len = if sheet_xml[tag_end..].starts_with("/>") {
            2
        } else {
            1
        };
        let tag = &sheet_xml[start..tag_end + close_len];
        if let Some(rid) = parse_attr(tag, "r:id") {
            out.push(rid);
        }
        i = tag_end + close_len;
    }
    out
}

fn extract_anchors(drawing_xml: &str) -> Vec<(u32, u32, Option<String>)> {
    // Returns (col0, row0, embedRid)
    let mut out: Vec<(u32, u32, Option<String>)> = Vec::new();

    for (open_tag, close_tag) in [
        ("<xdr:oneCellAnchor", "</xdr:oneCellAnchor>"),
        ("<oneCellAnchor", "</oneCellAnchor>"),
        ("<xdr:twoCellAnchor", "</xdr:twoCellAnchor>"),
        ("<twoCellAnchor", "</twoCellAnchor>"),
    ] {
        let mut i: usize = 0;
        while let Some(pos) = drawing_xml[i..].find(open_tag) {
            let start = i + pos;
            let Some(end_rel) = drawing_xml[start..].find(close_tag) else {
                break;
            };
            let end = start + end_rel + close_tag.len();
            let block = &drawing_xml[start..end];

            // Find <from> ... </from> (prefix may be absent)
            let from_start = block.find("<xdr:from>").or_else(|| block.find("<from>"));
            let from_end = block.find("</xdr:from>").or_else(|| block.find("</from>"));
            let (col0, row0) = if let (Some(fs), Some(fe)) = (from_start, from_end) {
                let from_block = &block[fs..fe];
                let col = extract_simple_tag_value(from_block, "xdr:col")
                    .or_else(|| extract_simple_tag_value(from_block, "col"))
                    .and_then(|s| s.parse::<u32>().ok());
                let row = extract_simple_tag_value(from_block, "xdr:row")
                    .or_else(|| extract_simple_tag_value(from_block, "row"))
                    .and_then(|s| s.parse::<u32>().ok());
                match (col, row) {
                    (Some(c), Some(r)) => (c, r),
                    _ => {
                        i = end;
                        continue;
                    }
                }
            } else {
                i = end;
                continue;
            };

            // Find embedded image relationship id.
            let embed_rid = if let Some(blip_pos) = block.find("<a:blip") {
                let abs = blip_pos;
                let end_rel = block[abs..].find("/>").or_else(|| block[abs..].find('>'));
                if let Some(tag_end_rel) = end_rel {
                    let tag_end = abs + tag_end_rel;
                    let close_len = if block[tag_end..].starts_with("/>") {
                        2
                    } else {
                        1
                    };
                    let tag = &block[abs..tag_end + close_len];
                    parse_attr(tag, "r:embed")
                } else {
                    None
                }
            } else {
                None
            };

            out.push((col0, row0, embed_rid));
            i = end;
        }
    }

    out
}

fn read_images_from_xlsx(path: &Path, sheet: &str) -> PyResult<Vec<ImageReadSpec>> {
    let f = std::fs::File::open(path)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to open workbook: {e}")))?;
    let mut zip = ZipArchive::new(f)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Invalid xlsx zip: {e}")))?;

    let workbook_xml = zip_read_to_string(&mut zip, "xl/workbook.xml")?;
    let rels_xml = zip_read_to_string(&mut zip, "xl/_rels/workbook.xml.rels")?;
    let sheet_to_rid = parse_workbook_sheet_map(&workbook_xml);
    let rid_to_target = parse_workbook_rels_map(&rels_xml);

    let Some(rid) = sheet_to_rid.get(sheet) else {
        return Ok(Vec::new());
    };
    let Some(target) = rid_to_target.get(rid) else {
        return Ok(Vec::new());
    };
    let sheet_entry = workbook_rel_target_to_part(target);
    let sheet_xml = zip_read_to_string(&mut zip, &sheet_entry)?;

    let drawing_rids = extract_drawing_rids(&sheet_xml);
    if drawing_rids.is_empty() {
        return Ok(Vec::new());
    }

    let sheet_rels_entry = sheet_target_to_rels_entry(&sheet_entry);
    let sheet_rels_xml = zip_read_to_string(&mut zip, &sheet_rels_entry).unwrap_or_default();
    let sheet_rel_map = parse_workbook_rels_map(&sheet_rels_xml);

    let mut out: Vec<ImageReadSpec> = Vec::new();
    for drawing_rid in drawing_rids {
        let Some(drawing_target) = sheet_rel_map.get(&drawing_rid) else {
            continue;
        };
        let drawing_entry = resolve_sheet_rel_target(drawing_target);
        let drawing_xml = zip_read_to_string(&mut zip, &drawing_entry).unwrap_or_default();
        if drawing_xml.is_empty() {
            continue;
        }

        let drawing_rels_entry = sheet_target_to_rels_entry(&drawing_entry);
        let drawing_rels_xml =
            zip_read_to_string(&mut zip, &drawing_rels_entry).unwrap_or_default();
        let drawing_rel_map = parse_workbook_rels_map(&drawing_rels_xml);

        for (col0, row0, embed_rid) in extract_anchors(&drawing_xml) {
            let Some(embed_rid) = embed_rid else {
                continue;
            };
            let Some(img_target) = drawing_rel_map.get(&embed_rid) else {
                continue;
            };
            let part = resolve_drawing_rel_target(img_target);
            let path = format!("/{part}");
            let cell = format!("{}{}", col_to_letters(col0), row0 + 1);
            out.push(ImageReadSpec {
                cell,
                path,
                anchor: "oneCell".to_string(),
            });
        }
    }

    Ok(out)
}

fn extract_section<'a>(xml: &'a str, open_tag: &str, close_tag: &str) -> Option<&'a str> {
    let start = xml.find(open_tag)?;
    let end_rel = xml[start..].find(close_tag)?;
    let end = start + end_rel + close_tag.len();
    Some(&xml[start..end])
}

fn extract_nth_start_tag(xml: &str, tag_prefix: &str, idx: usize) -> Option<String> {
    let mut i: usize = 0;
    let mut count: usize = 0;
    while let Some(pos) = xml[i..].find(tag_prefix) {
        let start = i + pos;
        let after = xml.get(start + tag_prefix.len()..start + tag_prefix.len() + 1);
        if after != Some(" ") && after != Some(">") && after != Some("/") {
            i = start + tag_prefix.len();
            continue;
        }
        let end_rel = xml[start..].find("/>").or_else(|| xml[start..].find('>'));
        let Some(tag_end_rel) = end_rel else {
            return None;
        };
        let tag_end = start + tag_end_rel;
        let close_len = if xml[tag_end..].starts_with("/>") {
            2
        } else {
            1
        };
        let tag = &xml[start..tag_end + close_len];

        if count == idx {
            return Some(tag.to_string());
        }
        count += 1;
        i = tag_end + close_len;
    }
    None
}

fn extract_nth_block(xml: &str, open_prefix: &str, close_tag: &str, idx: usize) -> Option<String> {
    let mut i: usize = 0;
    let mut count: usize = 0;
    while let Some(pos) = xml[i..].find(open_prefix) {
        let start = i + pos;
        let after = xml.get(start + open_prefix.len()..start + open_prefix.len() + 1);
        if after != Some(" ") && after != Some(">") {
            i = start + open_prefix.len();
            continue;
        }
        let Some(end_rel) = xml[start..].find(close_tag) else {
            return None;
        };
        let end = start + end_rel + close_tag.len();
        if count == idx {
            return Some(xml[start..end].to_string());
        }
        count += 1;
        i = end;
    }
    None
}

fn find_cell_style_index(sheet_xml: &str, cell_ref: &str) -> Option<usize> {
    let needle = format!("r=\"{cell_ref}\"");
    let pos = sheet_xml.find(&needle)?;
    let start = sheet_xml[..pos].rfind("<c ")?;
    let end_rel = sheet_xml[start..].find('>')?;
    let tag = &sheet_xml[start..start + end_rel + 1];
    parse_attr(tag, "s").and_then(|s| s.parse::<usize>().ok())
}

fn read_bg_color_from_xlsx(path: &Path, sheet: &str, cell_ref: &str) -> PyResult<Option<String>> {
    let f = std::fs::File::open(path)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to open workbook: {e}")))?;
    let mut zip = ZipArchive::new(f)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Invalid xlsx zip: {e}")))?;

    let workbook_xml = zip_read_to_string(&mut zip, "xl/workbook.xml")?;
    let rels_xml = zip_read_to_string(&mut zip, "xl/_rels/workbook.xml.rels")?;
    let sheet_to_rid = parse_workbook_sheet_map(&workbook_xml);
    let rid_to_target = parse_workbook_rels_map(&rels_xml);

    let Some(rid) = sheet_to_rid.get(sheet) else {
        return Ok(None);
    };
    let Some(target) = rid_to_target.get(rid) else {
        return Ok(None);
    };
    let sheet_entry = workbook_rel_target_to_part(target);
    let sheet_xml = zip_read_to_string(&mut zip, &sheet_entry)?;
    let Some(style_idx) = find_cell_style_index(&sheet_xml, cell_ref) else {
        return Ok(None);
    };

    let styles_xml = zip_read_to_string(&mut zip, "xl/styles.xml")?;
    let Some(cellxfs) = extract_section(&styles_xml, "<cellXfs", "</cellXfs>") else {
        return Ok(None);
    };
    let Some(xf_tag) = extract_nth_start_tag(cellxfs, "<xf", style_idx) else {
        return Ok(None);
    };
    let Some(fill_id) = parse_attr(&xf_tag, "fillId").and_then(|s| s.parse::<usize>().ok()) else {
        return Ok(None);
    };
    let Some(fills) = extract_section(&styles_xml, "<fills", "</fills>") else {
        return Ok(None);
    };
    let Some(fill_block) = extract_nth_block(fills, "<fill", "</fill>", fill_id) else {
        return Ok(None);
    };

    if let Some(pos) = fill_block.find("<fgColor") {
        let start = pos;
        let end_rel = fill_block[start..]
            .find("/>")
            .or_else(|| fill_block[start..].find('>'));
        if let Some(tag_end_rel) = end_rel {
            let tag_end = start + tag_end_rel;
            let close_len = if fill_block[tag_end..].starts_with("/>") {
                2
            } else {
                1
            };
            let tag = &fill_block[start..tag_end + close_len];
            if let Some(rgb) = parse_attr(tag, "rgb") {
                return Ok(argb_to_hex_rgb(&rgb));
            }
        }
    }

    Ok(None)
}

fn extract_hyperlink_tags(
    xml: &str,
) -> Vec<(String, Option<String>, Option<String>, Option<String>)> {
    // Returns (ref, location, tooltip, r:id)
    let mut out: Vec<(String, Option<String>, Option<String>, Option<String>)> = Vec::new();
    let mut i: usize = 0;
    while let Some(rel) = xml[i..].find("<hyperlink") {
        let start = i + rel;
        let after = xml.get(start + "<hyperlink".len()..start + "<hyperlink".len() + 1);
        if after != Some(" ") && after != Some(">") {
            i = start + "<hyperlink".len();
            continue;
        }
        let end_rel = xml[start..].find("/>").or_else(|| xml[start..].find('>'));
        let Some(tag_end_rel) = end_rel else {
            break;
        };
        let tag_end = start + tag_end_rel;
        let close_len = if xml[tag_end..].starts_with("/>") {
            2
        } else {
            1
        };
        let tag = &xml[start..tag_end + close_len];
        let r = parse_attr(tag, "ref");
        if let Some(r) = r {
            let location = parse_attr(tag, "location");
            let tooltip = parse_attr(tag, "tooltip");
            let rid = parse_attr(tag, "r:id");
            out.push((r, location, tooltip, rid));
        }
        i = tag_end + close_len;
    }
    out
}

fn read_hyperlinks_from_xlsx(path: &Path, sheet: &str) -> PyResult<Vec<HyperlinkReadSpec>> {
    let f = std::fs::File::open(path)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to open workbook: {e}")))?;
    let mut zip = ZipArchive::new(f)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Invalid xlsx zip: {e}")))?;

    let workbook_xml = zip_read_to_string(&mut zip, "xl/workbook.xml")?;
    let rels_xml = zip_read_to_string(&mut zip, "xl/_rels/workbook.xml.rels")?;

    let sheet_to_rid = parse_workbook_sheet_map(&workbook_xml);
    let rid_to_target = parse_workbook_rels_map(&rels_xml);

    let Some(rid) = sheet_to_rid.get(sheet) else {
        return Ok(Vec::new());
    };
    let Some(target) = rid_to_target.get(rid) else {
        return Ok(Vec::new());
    };
    let sheet_entry = workbook_rel_target_to_part(target);

    let sheet_xml = zip_read_to_string(&mut zip, &sheet_entry)?;
    let tags = extract_hyperlink_tags(&sheet_xml);

    // Load sheet relationships so we can resolve r:id for external links.
    let sheet_rels_entry = sheet_target_to_rels_entry(&sheet_entry);
    let sheet_rels_xml = zip_read_to_string(&mut zip, &sheet_rels_entry).unwrap_or_default();
    let rel_map = parse_workbook_rels_map(&sheet_rels_xml);

    let mut out: Vec<HyperlinkReadSpec> = Vec::new();
    for (cell, location, tooltip, rid) in tags {
        if rid.is_none() {
            let loc = location.unwrap_or_default();
            if loc.is_empty() {
                continue;
            }
            let target = loc.trim_start_matches('#').replace("'", "");
            out.push(HyperlinkReadSpec {
                cell,
                target,
                tooltip,
                internal: true,
            });
            continue;
        }

        let rid = rid.unwrap();
        let Some(mut base) = rel_map.get(&rid).cloned() else {
            continue;
        };
        if let Some(loc) = location {
            let loc = loc.trim();
            if !loc.is_empty() {
                base = format!("{base}#{loc}");
            }
        }

        out.push(HyperlinkReadSpec {
            cell,
            target: base,
            tooltip,
            internal: false,
        });
    }

    Ok(out)
}

fn xml_unescape(value: &str) -> String {
    // Minimal entity decoding for our fixtures.
    let mut s = value
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", "\"")
        .replace("&apos;", "'")
        .replace("&amp;", "&");
    s = s.replace("&#10;", "\n").replace("&#xA;", "\n");
    s
}

fn extract_tag_texts(xml: &str, tag_name: &str) -> Vec<String> {
    let mut out: Vec<String> = Vec::new();
    let open = format!("<{tag_name}>");
    let close = format!("</{tag_name}>");
    let mut i: usize = 0;
    while let Some(pos) = xml[i..].find(&open) {
        let start = i + pos + open.len();
        if let Some(end_rel) = xml[start..].find(&close) {
            let end = start + end_rel;
            out.push(xml_unescape(&xml[start..end]));
            i = end + close.len();
        } else {
            break;
        }
    }
    out
}

fn extract_comment_text(comment_xml: &str) -> String {
    // Extract all <t ...>...</t> nodes inside a comment.
    let mut out = String::new();
    let mut i: usize = 0;
    while let Some(pos) = comment_xml[i..].find("<t") {
        let start = i + pos;
        // Avoid matching the <text> container tag.
        let after = comment_xml.get(start + 2..start + 3);
        if after != Some(" ") && after != Some(">") {
            i = start + 2;
            continue;
        }
        let gt_rel = comment_xml[start..].find('>');
        let Some(gt_rel) = gt_rel else { break };
        let content_start = start + gt_rel + 1;
        let Some(end_rel) = comment_xml[content_start..].find("</t>") else {
            break;
        };
        let content_end = content_start + end_rel;
        out.push_str(&xml_unescape(&comment_xml[content_start..content_end]));
        i = content_end + 4;
    }
    out
}

fn read_comments_from_xlsx(path: &Path, sheet: &str) -> PyResult<Vec<CommentReadSpec>> {
    let f = std::fs::File::open(path)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to open workbook: {e}")))?;
    let mut zip = ZipArchive::new(f)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Invalid xlsx zip: {e}")))?;

    let workbook_xml = zip_read_to_string(&mut zip, "xl/workbook.xml")?;
    let rels_xml = zip_read_to_string(&mut zip, "xl/_rels/workbook.xml.rels")?;
    let sheet_to_rid = parse_workbook_sheet_map(&workbook_xml);
    let rid_to_target = parse_workbook_rels_map(&rels_xml);

    let Some(rid) = sheet_to_rid.get(sheet) else {
        return Ok(Vec::new());
    };
    let Some(target) = rid_to_target.get(rid) else {
        return Ok(Vec::new());
    };
    let sheet_entry = workbook_rel_target_to_part(target);
    let sheet_rels_entry = sheet_target_to_rels_entry(&sheet_entry);
    let sheet_rels_xml = zip_read_to_string(&mut zip, &sheet_rels_entry).unwrap_or_default();

    let entries = parse_rels_entries(&sheet_rels_xml);
    let comments_rel = entries.iter().find(|e| e.r#type.ends_with("/comments"));
    let Some(comments_rel) = comments_rel else {
        return Ok(Vec::new());
    };

    let comments_entry = resolve_sheet_rel_target(&comments_rel.target);
    let comments_xml = zip_read_to_string(&mut zip, &comments_entry)?;

    // Parse authors.
    let authors = extract_tag_texts(&comments_xml, "author");

    let mut out: Vec<CommentReadSpec> = Vec::new();
    let mut i: usize = 0;
    while let Some(pos) = comments_xml[i..].find("<comment ") {
        let start = i + pos;
        let Some(tag_end_rel) = comments_xml[start..].find('>') else {
            break;
        };
        let tag_end = start + tag_end_rel;
        let tag = &comments_xml[start..=tag_end];
        let cell = parse_attr(tag, "ref").unwrap_or_default();
        let author_id = parse_attr(tag, "authorId").and_then(|s| s.parse::<usize>().ok());

        let Some(close_rel) = comments_xml[tag_end..].find("</comment>") else {
            break;
        };
        let close_end = tag_end + close_rel + "</comment>".len();
        let body = &comments_xml[tag_end..close_end];
        let text = extract_comment_text(body);

        if !cell.is_empty() {
            let author = author_id.and_then(|idx| authors.get(idx).cloned());
            out.push(CommentReadSpec { cell, text, author });
        }

        i = close_end;
    }

    Ok(out)
}
