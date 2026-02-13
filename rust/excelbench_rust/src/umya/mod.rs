use pyo3::exceptions::{PyIOError, PyValueError};
use pyo3::prelude::*;

use std::path::Path;

use umya_spreadsheet::{new_file, reader, writer, Spreadsheet};

mod borders;
mod cell_values;
mod comments;
mod conditional_fmt;
mod data_validation;
mod dimensions;
mod formatting;
mod freeze_panes;
mod hyperlinks;
mod images;
mod merged_cells;
mod named_ranges;
mod tables;
mod util;

#[pyclass(unsendable)]
pub struct UmyaBook {
    pub(super) book: Spreadsheet,
    pub(super) saved: bool,
}

#[pymethods]
impl UmyaBook {
    #[new]
    pub fn new() -> Self {
        let mut book = new_file();
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
