use pyo3::exceptions::{PyIOError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::PyDict;

use std::path::Path;

use chrono::{Duration, NaiveDate, NaiveDateTime, NaiveTime};

use umya_spreadsheet::{new_file, reader, writer, NumberingFormat, Spreadsheet};

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
