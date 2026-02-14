use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};

use umya_spreadsheet::structs::{Table, TableColumn, TableStyleInfo};

use super::UmyaBook;

/// Convert a Table's area coordinates to a range string like "A1:D10".
fn area_to_ref(table: &Table) -> String {
    let (start, end) = table.get_area();
    let start_str = start.to_string().replace('$', "");
    let end_str = end.to_string().replace('$', "");
    format!("{start_str}:{end_str}")
}

/// Parse a range string like "A1:D10" into two cell references.
fn parse_range_ref(range_ref: &str) -> Option<(&str, &str)> {
    let parts: Vec<&str> = range_ref.split(':').collect();
    if parts.len() == 2 {
        Some((parts[0].trim(), parts[1].trim()))
    } else {
        None
    }
}

#[pymethods]
impl UmyaBook {
    pub fn read_tables(&self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        let ws = self
            .book
            .get_sheet_by_name(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let result = PyList::empty_bound(py);

        for table in ws.get_tables() {
            if !table.is_ok() {
                continue;
            }

            let d = PyDict::new_bound(py);
            let name = table.get_display_name();
            let name = if name.is_empty() {
                table.get_name()
            } else {
                name
            };
            d.set_item("name", name)?;
            d.set_item("ref", area_to_ref(table))?;

            // Header row: tables always have a header row unless explicitly disabled.
            // umya doesn't store headerRowCount separately — if the table exists, it has headers.
            d.set_item("header_row", true)?;

            // Totals row
            let has_totals = *table.get_totals_row_count() > 0;
            d.set_item("totals_row", has_totals)?;

            // Style info
            if let Some(style) = table.get_style_info() {
                d.set_item("style", style.get_name())?;
            } else {
                d.set_item("style", py.None())?;
            }

            // Columns
            let cols = PyList::empty_bound(py);
            for col in table.get_columns() {
                cols.append(col.get_name())?;
            }
            d.set_item("columns", cols)?;

            // AutoFilter: umya Table doesn't expose table-level autoFilter,
            // so we check if the worksheet has an autoFilter whose range matches this table.
            let has_af = ws.get_auto_filter()
                .map(|af| {
                    let af_range = af.get_range().get_range().replace('$', "");
                    let table_ref = area_to_ref(table);
                    af_range == table_ref
                })
                .unwrap_or(false);
            d.set_item("autofilter", has_af)?;

            result.append(d)?;
        }

        Ok(result.into())
    }

    pub fn add_table(
        &mut self,
        sheet: &str,
        table_dict: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let dict = table_dict
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("table must be a dict"))?;

        // Support optional wrapper key "table"
        let inner: Option<Bound<'_, PyAny>> = dict.get_item("table")?;
        let cfg: &Bound<'_, PyDict> = match &inner {
            Some(v) => v.downcast::<PyDict>().unwrap_or(dict),
            None => dict,
        };

        let name: String = cfg
            .get_item("name")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("table missing 'name'"))?
            .extract()?;
        let ref_str: String = cfg
            .get_item("ref")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("table missing 'ref'"))?
            .extract()?;

        let (start, end) = parse_range_ref(&ref_str)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Invalid ref: {ref_str}")))?;

        let mut table = Table::new(&name, (start, end));

        // Display name
        let display_name: Option<String> = cfg
            .get_item("display_name")?
            .and_then(|v| v.extract::<String>().ok());
        if let Some(dn) = display_name {
            table.set_display_name(&dn);
        }

        // Columns
        let columns: Option<Vec<String>> = cfg
            .get_item("columns")?
            .and_then(|v| v.extract::<Vec<String>>().ok());
        if let Some(cols) = columns {
            for col_name in &cols {
                table.add_column(TableColumn::new(col_name));
            }
        }

        // Style info
        let style_name: Option<String> = cfg
            .get_item("style")?
            .and_then(|v| v.extract::<String>().ok());
        if let Some(sn) = style_name {
            let style = TableStyleInfo::new(&sn, false, false, true, false);
            table.set_style_info(Some(style));
        }

        // Totals row
        let totals_row: Option<bool> = cfg
            .get_item("totals_row")?
            .and_then(|v| v.extract::<bool>().ok());
        if totals_row == Some(true) {
            table.set_totals_row_count(1);
            table.set_totals_row_shown(true);
        }

        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        ws.add_table(table);

        // If autofilter is requested, set worksheet-level auto filter on the table range.
        let autofilter: Option<bool> = cfg
            .get_item("autofilter")?
            .and_then(|v| v.extract::<bool>().ok());
        if autofilter == Some(true) {
            ws.set_auto_filter(&ref_str);
        }

        Ok(())
    }
}
