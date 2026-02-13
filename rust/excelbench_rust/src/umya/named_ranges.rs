use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};

use super::UmyaBook;

/// Normalize a defined name address for benchmark comparison.
/// Strips leading `=`, `$` signs, and surrounding quotes from sheet names.
fn normalize_address(raw: &str) -> String {
    let raw = raw.trim().trim_start_matches('=');
    if raw.is_empty() {
        return String::new();
    }

    // Split on `!` to separate sheet name from cell address.
    if let Some((sheet_part, addr_part)) = raw.split_once('!') {
        let sheet = sheet_part
            .trim_start_matches('\'')
            .trim_end_matches('\'')
            .replace("''", "'");
        let addr = addr_part.replace('$', "");
        format!("{sheet}!{addr}")
    } else {
        raw.replace('$', "")
    }
}

#[pymethods]
impl UmyaBook {
    pub fn read_named_ranges(&self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        let result = PyList::empty_bound(py);

        // 1. Workbook-level defined names (no localSheetId).
        for dn in self.book.get_defined_names() {
            if dn.has_local_sheet_id() {
                continue;
            }
            let d = PyDict::new_bound(py);
            d.set_item("name", dn.get_name())?;
            d.set_item("scope", "workbook")?;
            d.set_item("refers_to", normalize_address(&dn.get_address()))?;
            result.append(d)?;
        }

        // 2. Worksheet-level defined names.
        //    umya puts workbook-scoped names that reference a sheet onto the
        //    worksheet, so we include both those and truly sheet-scoped ones.
        //    We need the sheet index to detect localSheetId.
        if let Some(ws) = self.book.get_sheet_by_name(sheet) {
            for dn in ws.get_defined_names() {
                let d = PyDict::new_bound(py);
                d.set_item("name", dn.get_name())?;
                // If localSheetId is set, it's sheet-scoped;
                // otherwise it's a workbook-scoped name stored on the sheet.
                let scope = if dn.has_local_sheet_id() {
                    "sheet"
                } else {
                    "workbook"
                };
                d.set_item("scope", scope)?;
                d.set_item("refers_to", normalize_address(&dn.get_address()))?;
                result.append(d)?;
            }
        }

        Ok(result.into())
    }

    pub fn add_named_range(
        &mut self,
        sheet: &str,
        nr_dict: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let dict = nr_dict
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("named_range must be a dict"))?;

        // Support optional wrapper key "named_range"
        let inner: Option<Bound<'_, PyAny>> = dict.get_item("named_range")?;
        let cfg: &Bound<'_, PyDict> = match &inner {
            Some(v) => v.downcast::<PyDict>().unwrap_or(dict),
            None => dict,
        };

        let name: String = cfg
            .get_item("name")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("named_range missing 'name'"))?
            .extract()?;
        let refers_to: String = cfg
            .get_item("refers_to")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("named_range missing 'refers_to'"))?
            .extract()?;
        let scope: String = cfg
            .get_item("scope")?
            .map(|v| v.extract::<String>())
            .transpose()?
            .unwrap_or_else(|| "workbook".to_string());

        // Ensure refers_to starts with `=` for the address parser if it
        // doesn't already — umya's `set_address` handles raw refs fine.
        let address = refers_to.trim_start_matches('=').to_string();

        // Use the worksheet-level convenience method which can call the
        // `pub(crate)` `set_name()`. This works for both scopes since the
        // write path serializes them correctly.
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        ws.add_defined_name(&name, &address)
            .map_err(|e| PyErr::new::<PyValueError, _>(format!("Failed to add named range: {e}")))?;

        // For sheet-scoped names, set the localSheetId on the just-added
        // DefinedName so it roundtrips correctly.
        if scope == "sheet" {
            // Find the sheet index to use as localSheetId.
            let sheet_names: Vec<String> = self
                .book
                .get_sheet_collection()
                .iter()
                .map(|s| s.get_name().to_string())
                .collect();
            if let Some(idx) = sheet_names.iter().position(|n| n == sheet) {
                let ws = self
                    .book
                    .get_sheet_by_name_mut(sheet)
                    .unwrap();
                if let Some(last) = ws.get_defined_names_mut().last_mut() {
                    last.set_local_sheet_id(idx as u32);
                }
            }
        }

        Ok(())
    }
}
