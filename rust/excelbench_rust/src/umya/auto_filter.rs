use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;

use super::UmyaBook;

#[pymethods]
impl UmyaBook {
    /// Read the auto filter range for a sheet, or None if not set.
    pub fn get_auto_filter(&self, sheet: &str) -> PyResult<Option<String>> {
        let ws = self
            .book
            .get_sheet_by_name(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        Ok(ws.get_auto_filter().map(|af| {
            af.get_range().get_range().replace('$', "")
        }))
    }

    /// Set an auto filter on a range (e.g. "A1:D10").
    pub fn set_auto_filter(&mut self, sheet: &str, range: &str) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        ws.set_auto_filter(range);
        Ok(())
    }

    /// Remove the auto filter from a sheet.
    pub fn remove_auto_filter(&mut self, sheet: &str) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        ws.remove_auto_filter();
        Ok(())
    }

    /// Check if a sheet has an auto filter.
    pub fn has_auto_filter(&self, sheet: &str) -> PyResult<bool> {
        let ws = self
            .book
            .get_sheet_by_name(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        Ok(ws.get_auto_filter().is_some())
    }
}
