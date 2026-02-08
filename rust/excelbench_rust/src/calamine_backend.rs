use pyo3::exceptions::{PyIOError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::PyDict;

use std::fs::File;
use std::io::BufReader;

use calamine::{open_workbook_auto, Data, Reader, Sheets};

use chrono::NaiveTime;

type CalamineSheets = Sheets<BufReader<File>>;

use crate::util::{a1_to_row_col, cell_blank, cell_with_value, parse_iso_date, parse_iso_datetime};

fn map_error_value(err_str: &str) -> &'static str {
    // Best-effort normalization. If the underlying error representation changes,
    // callers still get a stable Excel-like token.
    let e = err_str.to_ascii_uppercase();
    match e.as_str() {
        "DIV0" | "DIV/0" | "#DIV/0!" => "#DIV/0!",
        "NA" | "#N/A" => "#N/A",
        "VALUE" | "#VALUE!" => "#VALUE!",
        "REF" | "#REF!" => "#REF!",
        "NAME" | "#NAME?" => "#NAME?",
        "NUM" | "#NUM!" => "#NUM!",
        "NULL" | "#NULL!" => "#NULL!",
        _ => "#ERROR!",
    }
}

#[pyclass(unsendable)]
pub struct CalamineBook {
    workbook: CalamineSheets,
    sheet_names: Vec<String>,
}

#[pymethods]
impl CalamineBook {
    #[staticmethod]
    pub fn open(path: &str) -> PyResult<Self> {
        let wb = open_workbook_auto(path)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to open workbook: {e}")))?;
        let names = wb.sheet_names().to_vec();
        Ok(Self {
            workbook: wb,
            sheet_names: names,
        })
    }

    pub fn sheet_names(&self) -> Vec<String> {
        self.sheet_names.clone()
    }

    pub fn read_cell_value(&mut self, py: Python<'_>, sheet: &str, a1: &str) -> PyResult<PyObject> {
        let (row, col) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;

        if !self.sheet_names.iter().any(|name| name == sheet) {
            return Err(PyErr::new::<PyValueError, _>(format!(
                "Unknown sheet: {sheet}"
            )));
        }

        let range = self.workbook.worksheet_range(sheet).map_err(|e| {
            PyErr::new::<PyIOError, _>(format!("Failed to read sheet {sheet}: {e}"))
        })?;

        let value = match range.get_value((row, col)) {
            None => return cell_blank(py),
            Some(v) => v,
        };

        let out = match value {
            Data::Empty => cell_blank(py)?,
            Data::String(s) => cell_with_value(py, "string", s.clone())?,
            Data::Float(f) => cell_with_value(py, "number", *f)?,
            Data::Int(i) => cell_with_value(py, "number", *i as f64)?,
            Data::Bool(b) => cell_with_value(py, "boolean", *b)?,

            // Date/datetime and durations: avoid debug-string garbage.
            // - DateTime(f64): Excel serial date/time
            // - DateTimeIso(String): ISO-8601-like string
            // - Duration(f64): numeric duration
            // - DurationIso(String): ISO duration string
            Data::DateTime(dt) => {
                // Preserve date vs datetime semantics for the harness.
                // If time component is midnight, surface as a DATE.
                if let Some(ndt) = dt.as_datetime() {
                    let midnight = NaiveTime::from_hms_opt(0, 0, 0).unwrap();
                    if ndt.time() == midnight {
                        let s = ndt.date().format("%Y-%m-%d").to_string();
                        cell_with_value(py, "date", s)?
                    } else {
                        let s = ndt.format("%Y-%m-%dT%H:%M:%S").to_string();
                        cell_with_value(py, "datetime", s)?
                    }
                } else {
                    // Fallback: report the raw Excel serial.
                    cell_with_value(py, "number", dt.as_f64())?
                }
            }
            Data::DateTimeIso(s) => {
                // Best-effort parse for midnight -> date.
                let raw = s.trim_end_matches('Z');
                if let Some(d) = parse_iso_date(raw) {
                    cell_with_value(py, "date", d.format("%Y-%m-%d").to_string())?
                } else if let Some(ndt) = parse_iso_datetime(raw) {
                    let midnight = NaiveTime::from_hms_opt(0, 0, 0).unwrap();
                    if ndt.time() == midnight {
                        cell_with_value(py, "date", ndt.date().format("%Y-%m-%d").to_string())?
                    } else {
                        cell_with_value(
                            py,
                            "datetime",
                            ndt.format("%Y-%m-%dT%H:%M:%S").to_string(),
                        )?
                    }
                } else {
                    // If parsing fails (timezone offsets, etc), keep the ISO string.
                    cell_with_value(py, "datetime", s.clone())?
                }
            }
            Data::DurationIso(s) => cell_with_value(py, "string", s.clone())?,

            Data::Error(e) => {
                let normalized = map_error_value(&format!("{e:?}"));
                let d = PyDict::new_bound(py);
                d.set_item("type", "error")?;
                d.set_item("value", normalized)?;
                d.into()
            }
        };

        Ok(out)
    }
}
