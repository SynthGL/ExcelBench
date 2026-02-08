use pyo3::exceptions::{PyIOError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::PyDict;

use indexmap::IndexMap;

use rust_xlsxwriter::{Format, Workbook, Worksheet};

use crate::util::{a1_to_row_col, parse_iso_date, parse_iso_datetime};

#[pyclass(unsendable)]
pub struct RustXlsxWriterBook {
    sheets: IndexMap<String, Worksheet>,
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

        let mut ws = Worksheet::new();
        ws.set_name(name)
            .map_err(|e| PyErr::new::<PyValueError, _>(format!("Invalid sheet name: {e}")))?;
        self.sheets.insert(name.to_string(), ws);
        Ok(())
    }

    pub fn write_cell_value(
        &mut self,
        sheet: &str,
        a1: &str,
        payload: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let ws = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let (row, col_0) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
        let col: u16 = col_0.try_into().map_err(|_| {
            PyErr::new::<PyValueError, _>(format!("Column out of range for Excel: {a1}"))
        })?;

        let dict = payload
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("payload must be a dict"))?;
        let type_obj = dict
            .get_item("type")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("payload missing 'type'"))?;
        let type_str: String = type_obj.extract()?;

        match type_str.as_str() {
            "blank" => {
                // rust_xlsxwriter doesn't require explicit blank writes.
                Ok(())
            }
            "string" => {
                let v = dict.get_item("value")?;
                let s = match v {
                    Some(v) => v.extract::<String>()?,
                    None => String::new(),
                };
                ws.write_string(row, col, s)
                    .map(|_| ())
                    .map_err(|e| PyErr::new::<PyIOError, _>(format!("write_string failed: {e}")))
            }
            "number" => {
                let v = dict.get_item("value")?.ok_or_else(|| {
                    PyErr::new::<PyValueError, _>("number payload missing 'value'")
                })?;
                let f = v.extract::<f64>()?;
                ws.write_number(row, col, f)
                    .map(|_| ())
                    .map_err(|e| PyErr::new::<PyIOError, _>(format!("write_number failed: {e}")))
            }
            "boolean" => {
                let v = dict.get_item("value")?.ok_or_else(|| {
                    PyErr::new::<PyValueError, _>("boolean payload missing 'value'")
                })?;
                let b = v.extract::<bool>()?;
                ws.write_boolean(row, col, b)
                    .map(|_| ())
                    .map_err(|e| PyErr::new::<PyIOError, _>(format!("write_boolean failed: {e}")))
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
                ws.write_formula(row, col, formula.as_str())
                    .map(|_| ())
                    .map_err(|e| PyErr::new::<PyIOError, _>(format!("write_formula failed: {e}")))
            }
            "error" => {
                let v = dict.get_item("value")?.ok_or_else(|| {
                    PyErr::new::<PyValueError, _>("error payload missing 'value'")
                })?;
                let token = v.extract::<String>()?;

                // Prefer error formulas that OpenpyxlAdapter can recognize.
                // For other errors, write the literal token as a string.
                let formula = match token.as_str() {
                    "#DIV/0!" => Some("=1/0"),
                    "#N/A" => Some("=NA()"),
                    "#VALUE!" => Some("=\"text\"+1"),
                    _ => None,
                };
                if let Some(f) = formula {
                    ws.write_formula(row, col, f).map(|_| ()).map_err(|e| {
                        PyErr::new::<PyIOError, _>(format!("write_formula failed: {e}"))
                    })
                } else {
                    ws.write_string(row, col, token).map(|_| ()).map_err(|e| {
                        PyErr::new::<PyIOError, _>(format!("write_string failed: {e}"))
                    })
                }
            }

            "date" => {
                let v = dict
                    .get_item("value")?
                    .ok_or_else(|| PyErr::new::<PyValueError, _>("date payload missing 'value'"))?;
                let s = v.extract::<String>()?;
                if let Some(d) = parse_iso_date(&s) {
                    let fmt = Format::new().set_num_format("yyyy-mm-dd");
                    ws.write_datetime_with_format(row, col, d, &fmt)
                        .map(|_| ())
                        .map_err(|e| {
                            PyErr::new::<PyIOError, _>(format!("write_datetime failed: {e}"))
                        })
                } else {
                    ws.write_string(row, col, s).map(|_| ()).map_err(|e| {
                        PyErr::new::<PyIOError, _>(format!("write_string failed: {e}"))
                    })
                }
            }

            "datetime" => {
                let v = dict.get_item("value")?.ok_or_else(|| {
                    PyErr::new::<PyValueError, _>("datetime payload missing 'value'")
                })?;
                let s = v.extract::<String>()?;
                if let Some(dt) = parse_iso_datetime(&s) {
                    let fmt = Format::new().set_num_format("yyyy-mm-dd hh:mm:ss");
                    ws.write_datetime_with_format(row, col, dt, &fmt)
                        .map(|_| ())
                        .map_err(|e| {
                            PyErr::new::<PyIOError, _>(format!("write_datetime failed: {e}"))
                        })
                } else {
                    ws.write_string(row, col, s).map(|_| ()).map_err(|e| {
                        PyErr::new::<PyIOError, _>(format!("write_string failed: {e}"))
                    })
                }
            }

            other => Err(PyErr::new::<PyValueError, _>(format!(
                "Unsupported cell type: {other}"
            ))),
        }
    }

    pub fn save(&mut self, path: &str) -> PyResult<()> {
        if self.saved {
            return Err(PyErr::new::<PyValueError, _>(
                "Workbook already saved (RustXlsxWriterBook is consumed-on-save)",
            ));
        }
        self.saved = true;

        let mut wb = Workbook::new();

        for (_name, ws) in self.sheets.drain(..) {
            wb.push_worksheet(ws);
        }

        wb.save(path)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to save workbook: {e}")))
    }
}
