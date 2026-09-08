use std::collections::BTreeMap;
use std::fs;
use std::io::{self, Read};
use std::path::Path;

use serde::Deserialize;
use serde_json::{json, Map, Value};
use zavora_xlsx::{
    Align, BorderStyle, CellValue, Chart, ChartType, Format, Table, TableColumn, TableStyle,
    Underline, Workbook,
};

#[derive(Debug, Deserialize)]
struct OracleRequest {
    fixture_id: String,
    operation: String,
    #[serde(default)]
    payload: Value,
    input_path: Option<String>,
    output_path: Option<String>,
}

fn main() {
    let response = run().unwrap_or_else(|message| failure(None, None, message));
    println!("{response}");
}

fn run() -> Result<Value, String> {
    let mut input = String::new();
    io::stdin()
        .read_to_string(&mut input)
        .map_err(|error| format!("read request: {error}"))?;
    let request: OracleRequest =
        serde_json::from_str(&input).map_err(|error| format!("decode request: {error}"))?;

    match request.operation.as_str() {
        "write_fixture" => write_fixture(&request),
        "mutate" => mutate(&request),
        "calculate" => calculate(&request),
        operation => Err(format!("unsupported operation {operation:?}")),
    }
}

fn failure(fixture_id: Option<&str>, operation: Option<&str>, message: String) -> Value {
    let mut response = Map::new();
    response.insert("ok".into(), Value::Bool(false));
    response.insert("error".into(), Value::String("zavora_oracle_failed".into()));
    response.insert("message".into(), Value::String(message));
    if let Some(fixture_id) = fixture_id {
        response.insert("fixture_id".into(), Value::String(fixture_id.into()));
    }
    if let Some(operation) = operation {
        response.insert("operation".into(), Value::String(operation.into()));
    }
    Value::Object(response)
}

fn write_fixture(request: &OracleRequest) -> Result<Value, String> {
    let output_path = required_path(request.output_path.as_deref(), "write_fixture requires output_path")?;
    ensure_supported_payload(&request.payload)?;

    let mut workbook = Workbook::new();
    configure_sheets(&mut workbook, &request.payload)?;
    apply_merges(&mut workbook, &request.payload)?;
    apply_cells(&mut workbook, &request.payload)?;
    apply_formats(&mut workbook, &request.payload)?;
    apply_borders(&mut workbook, &request.payload)?;
    apply_columns(&mut workbook, &request.payload)?;
    apply_row_heights(&mut workbook, &request.payload)?;
    apply_hyperlinks(&mut workbook, &request.payload)?;
    apply_comments(&mut workbook, &request.payload)?;
    apply_panes(&mut workbook, &request.payload)?;
    apply_named_ranges(&mut workbook, &request.payload)?;
    apply_tables(&mut workbook, &request.payload)?;
    apply_charts(&mut workbook, &request.payload)?;

    create_parent(output_path)?;
    workbook
        .save(output_path)
        .map_err(|error| format!("save workbook: {error}"))?;

    Ok(json!({
        "ok": true,
        "fixture_id": request.fixture_id,
        "operation": request.operation,
        "output_path": output_path,
        "tool": "zavora-xlsx",
        "counts": payload_counts(&request.payload),
    }))
}

fn mutate(request: &OracleRequest) -> Result<Value, String> {
    let input_path = required_path(request.input_path.as_deref(), "mutate requires input_path")?;
    let output_path = required_path(request.output_path.as_deref(), "mutate requires output_path")?;
    let payload = object(&request.payload, "mutate payload")?;
    let sheet = required_string(payload, "sheet")?;
    let cell = required_string(payload, "cell")?;
    let value = payload
        .get("value")
        .ok_or_else(|| "mutate payload requires value".to_string())?;
    let (row, col) = parse_cell(cell)?;
    let mut workbook = Workbook::open(input_path).map_err(|error| format!("open workbook: {error}"))?;
    write_json_value(workbook.worksheet_by_name(sheet).map_err(|error| error.to_string())?, row, col, value)?;
    create_parent(output_path)?;
    workbook
        .save(output_path)
        .map_err(|error| format!("save workbook: {error}"))?;
    Ok(json!({
        "ok": true,
        "fixture_id": request.fixture_id,
        "operation": request.operation,
        "output_path": output_path,
        "tool": "zavora-xlsx",
    }))
}

fn calculate(request: &OracleRequest) -> Result<Value, String> {
    let input_path = required_path(request.input_path.as_deref(), "calculate requires input_path")?;
    let output_path = required_path(request.output_path.as_deref(), "calculate requires output_path")?;
    let mut workbook = Workbook::open(input_path).map_err(|error| format!("open workbook: {error}"))?;
    let formula_cells = collect_formula_cells(&mut workbook)?;
    workbook
        .recalculate()
        .map_err(|error| format!("recalculate workbook: {error}"))?;
    let mut values = BTreeMap::new();
    for (sheet_name, row, col) in formula_cells {
        let value = workbook
            .worksheet_by_name(&sheet_name)
            .map_err(|error| error.to_string())?
            .read_cell(row, col);
        values.insert(
            format!("{sheet_name}!{}", cell_name(row, col)),
            calculated_value_to_json(&value),
        );
    }
    create_parent(output_path)?;
    workbook
        .save(output_path)
        .map_err(|error| format!("save workbook: {error}"))?;
    Ok(json!({
        "ok": true,
        "fixture_id": request.fixture_id,
        "operation": request.operation,
        "output_path": output_path,
        "tool": "zavora-xlsx",
        "values": values,
    }))
}

fn collect_formula_cells(workbook: &mut Workbook) -> Result<Vec<(String, u32, u16)>, String> {
    let sheet_names: Vec<String> = workbook.sheet_names().into_iter().map(str::to_owned).collect();
    let mut formulas = Vec::new();
    for sheet_name in sheet_names {
        let range = {
            let sheet = workbook
                .worksheet_by_name(&sheet_name)
                .map_err(|error| error.to_string())?;
            sheet.used_range()
        };
        if let Some((first_row, first_col, last_row, last_col)) = range {
            let sheet = workbook
                .worksheet_by_name(&sheet_name)
                .map_err(|error| error.to_string())?;
            for row in first_row..=last_row {
                for col in first_col..=last_col {
                    if matches!(sheet.read_cell(row, col), CellValue::Formula { .. }) {
                        formulas.push((sheet_name.clone(), row, col));
                    }
                }
            }
        }
    }
    Ok(formulas)
}

fn calculated_value_to_json(value: &CellValue) -> Value {
    match value {
        CellValue::Formula { cached_value, .. } => calculated_value_to_json(cached_value),
        CellValue::Number(number) => json!(number),
        CellValue::String(text) | CellValue::Error(text) => json!(text),
        CellValue::Bool(boolean) => json!(boolean),
        CellValue::DateTime(date_time) => json!(date_time.serial()),
        CellValue::Empty | CellValue::RichText(_) => Value::Null,
    }
}

fn configure_sheets(workbook: &mut Workbook, payload: &Value) -> Result<(), String> {
    let sheets = array(payload, "sheets")?;
    if sheets.is_empty() {
        return Ok(());
    }
    for (index, sheet) in sheets.iter().enumerate() {
        let name = required_string(object(sheet, "sheet entry")?, "name")?;
        if index == 0 {
            workbook
                .rename_worksheet(0, name)
                .map_err(|error| format!("rename first sheet: {error}"))?;
        } else {
            workbook
                .add_worksheet_with_name(name)
                .map_err(|error| format!("add sheet {name:?}: {error}"))?;
        }
    }
    Ok(())
}

fn apply_cells(workbook: &mut Workbook, payload: &Value) -> Result<(), String> {
    for entry in array(payload, "cells")? {
        let cell = object(entry, "cell entry")?;
        let sheet_name = required_string(cell, "sheet")?;
        let address = required_string(cell, "cell")?;
        let (row, col) = parse_cell(address)?;
        let sheet = workbook
            .worksheet_by_name(sheet_name)
            .map_err(|error| format!("cell {sheet_name}!{address}: {error}"))?;
        let cell_type = optional_string(cell, "type").unwrap_or("");
        if cell_type == "blank" {
            sheet
                .write_blank(row, col, &Format::new())
                .map_err(|error| error.to_string())?;
        } else if cell_type == "formula" || cell.get("formula").is_some_and(Value::is_string) {
            let formula = optional_string(cell, "formula")
                .or_else(|| cell.get("value").and_then(Value::as_str))
                .ok_or_else(|| format!("formula cell {sheet_name}!{address} requires formula or string value"))?;
            sheet
                .write_formula(row, col, formula)
                .map_err(|error| error.to_string())?;
        } else if matches!(cell_type, "date" | "datetime" | "error") {
            return Err(format!("unsupported cell type {cell_type:?} at {sheet_name}!{address}"));
        } else {
            let value = cell.get("value").unwrap_or(&Value::Null);
            write_json_value(sheet, row, col, value)?;
        }
    }
    Ok(())
}

fn write_json_value(
    sheet: &mut zavora_xlsx::Worksheet,
    row: u32,
    col: u16,
    value: &Value,
) -> Result<(), String> {
    match value {
        Value::Null => sheet.write_blank(row, col, &Format::new()),
        Value::Bool(boolean) => sheet.write(row, col, *boolean),
        Value::Number(number) => sheet.write(
            row,
            col,
            number
                .as_f64()
                .ok_or_else(|| "cell number cannot be represented as f64".to_string())?,
        ),
        Value::String(text) => sheet.write(row, col, text.as_str()),
        Value::Array(_) | Value::Object(_) => {
            return Err("cell value must be number, string, boolean, or null".into())
        }
    }
    .map(|_| ())
    .map_err(|error| format!("write cell: {error}"))
}

fn apply_formats(workbook: &mut Workbook, payload: &Value) -> Result<(), String> {
    for entry in array(payload, "formats")? {
        let spec = object(entry, "format entry")?;
        let sheet_name = required_string(spec, "sheet")?;
        let address = required_string(spec, "cell")?;
        let (row, col) = parse_cell(address)?;
        let format = build_format(spec)?;
        workbook
            .worksheet_by_name(sheet_name)
            .map_err(|error| error.to_string())?
            .set_cell_format(row, col, &format)
            .map_err(|error| format!("set format {sheet_name}!{address}: {error}"))?;
    }
    Ok(())
}

fn apply_borders(workbook: &mut Workbook, payload: &Value) -> Result<(), String> {
    for entry in array(payload, "borders")? {
        let spec = object(entry, "border entry")?;
        let sheet_name = required_string(spec, "sheet")?;
        let address = required_string(spec, "cell")?;
        let border = object(
            spec.get("border").ok_or_else(|| "border entry requires border".to_string())?,
            "border",
        )?;
        let mut format = Format::new();
        for (side, apply) in [
            ("top", Format::border_top as fn(Format, BorderStyle) -> Format),
            ("bottom", Format::border_bottom as fn(Format, BorderStyle) -> Format),
            ("left", Format::border_left as fn(Format, BorderStyle) -> Format),
            ("right", Format::border_right as fn(Format, BorderStyle) -> Format),
        ] {
            if let Some(side_value) = border.get(side) {
                let side_spec = object(side_value, "border side")?;
                let style = optional_string(side_spec, "style").unwrap_or("none");
                format = apply(format, border_style(style)?);
                if let Some(color) = optional_string(side_spec, "color") {
                    format = match side {
                        "top" => format.border_top_color(color),
                        "bottom" => format.border_bottom_color(color),
                        "left" => format.border_left_color(color),
                        "right" => format.border_right_color(color),
                        _ => format,
                    };
                }
            }
        }
        let (row, col) = parse_cell(address)?;
        workbook
            .worksheet_by_name(sheet_name)
            .map_err(|error| error.to_string())?
            .set_cell_format(row, col, &format)
            .map_err(|error| format!("set border {sheet_name}!{address}: {error}"))?;
    }
    Ok(())
}

fn build_format(spec: &Map<String, Value>) -> Result<Format, String> {
    let mut format = Format::new();
    if optional_bool(spec, "bold") == Some(true) { format = format.bold(); }
    if optional_bool(spec, "italic") == Some(true) { format = format.italic(); }
    if optional_bool(spec, "strikethrough") == Some(true) { format = format.strikethrough(); }
    if let Some(underline) = optional_string(spec, "underline") {
        format = format.underline(match underline {
            "none" | "" => Underline::None,
            "single" => Underline::Single,
            "double" => Underline::Double,
            _ => return Err(format!("unsupported underline {underline:?}")),
        });
    }
    if let Some(name) = optional_string(spec, "font_name") { format = format.font_name(name); }
    if let Some(size) = optional_number(spec, "font_size") { format = format.font_size(size); }
    if let Some(color) = optional_string(spec, "font_color") { format = format.font_color(color); }
    if let Some(color) = optional_string(spec, "bg_color") { format = format.background_color(color); }
    if let Some(number_format) = optional_string(spec, "number_format") { format = format.num_format(number_format); }
    if let Some(horizontal) = optional_string(spec, "h_align") { format = format.align(horizontal_alignment(horizontal)?); }
    if let Some(vertical) = optional_string(spec, "v_align") { format = format.align(vertical_alignment(vertical)?); }
    if optional_bool(spec, "wrap") == Some(true) { format = format.text_wrap(); }
    if let Some(rotation) = optional_number(spec, "rotation") { format = format.rotation(rotation as i16); }
    if let Some(indent) = optional_number(spec, "indent") { format = format.indent(indent as u8); }
    Ok(format)
}

fn apply_columns(workbook: &mut Workbook, payload: &Value) -> Result<(), String> {
    for entry in array(payload, "columns")? {
        let spec = object(entry, "column entry")?;
        let sheet_name = required_string(spec, "sheet")?;
        let start = required_string(spec, "start")?;
        let end = required_string(spec, "end")?;
        let width = required_number(spec, "width")?;
        let start_col = parse_column(start)?;
        let end_col = parse_column(end)?;
        if start_col > end_col { return Err(format!("invalid column range {start}:{end}")); }
        let sheet = workbook.worksheet_by_name(sheet_name).map_err(|error| error.to_string())?;
        for col in start_col..=end_col {
            sheet.set_column_width(col, width).map_err(|error| error.to_string())?;
        }
    }
    Ok(())
}

fn apply_row_heights(workbook: &mut Workbook, payload: &Value) -> Result<(), String> {
    for entry in array(payload, "row_heights")? {
        let spec = object(entry, "row height entry")?;
        let sheet_name = required_string(spec, "sheet")?;
        let row = required_u32(spec, "row")?.checked_sub(1).ok_or_else(|| "row must be positive".to_string())?;
        let height = required_number(spec, "height")?;
        workbook.worksheet_by_name(sheet_name).map_err(|error| error.to_string())?
            .set_row_height(row, height).map_err(|error| error.to_string())?;
    }
    Ok(())
}

fn apply_merges(workbook: &mut Workbook, payload: &Value) -> Result<(), String> {
    for entry in array(payload, "merges")? {
        let spec = object(entry, "merge entry")?;
        let sheet_name = required_string(spec, "sheet")?;
        let range = required_string(spec, "range")?;
        let (first_row, first_col, last_row, last_col) = parse_range(range)?;
        workbook.worksheet_by_name(sheet_name).map_err(|error| error.to_string())?
            .merge_range(first_row, first_col, last_row, last_col, "", &Format::new())
            .map_err(|error| format!("merge {sheet_name}!{range}: {error}"))?;
    }
    Ok(())
}

fn apply_hyperlinks(workbook: &mut Workbook, payload: &Value) -> Result<(), String> {
    for entry in array(payload, "hyperlinks")? {
        let spec = object(entry, "hyperlink entry")?;
        if spec.get("tooltip").is_some_and(|value| !value.is_null() && value != "") {
            return Err("unsupported hyperlink part: tooltip".into());
        }
        let sheet_name = required_string(spec, "sheet")?;
        let address = required_string(spec, "cell")?;
        let target = required_string(spec, "target")?;
        let display = optional_string(spec, "display").unwrap_or("");
        let (row, col) = parse_cell(address)?;
        let sheet = workbook.worksheet_by_name(sheet_name).map_err(|error| error.to_string())?;
        if optional_bool(spec, "internal") == Some(true) {
            sheet.write_internal_link(row, col, target, display).map_err(|error| error.to_string())?;
        } else {
            sheet.write_url(row, col, target, display).map_err(|error| error.to_string())?;
        }
    }
    Ok(())
}

fn apply_comments(workbook: &mut Workbook, payload: &Value) -> Result<(), String> {
    for entry in array(payload, "comments")? {
        let spec = object(entry, "comment entry")?;
        let sheet_name = required_string(spec, "sheet")?;
        let address = required_string(spec, "cell")?;
        let text = required_string(spec, "text")?;
        let author = optional_string(spec, "author").unwrap_or("Author");
        let (row, col) = parse_cell(address)?;
        workbook.worksheet_by_name(sheet_name).map_err(|error| error.to_string())?
            .add_comment_with_author(row, col, text, author);
    }
    Ok(())
}

fn apply_panes(workbook: &mut Workbook, payload: &Value) -> Result<(), String> {
    for entry in array(payload, "panes")? {
        let spec = object(entry, "pane entry")?;
        if let Some(mode) = optional_string(spec, "mode") {
            if mode != "" && mode != "freeze" { return Err(format!("unsupported pane mode {mode:?}")); }
        }
        if spec.get("top_left_cell").is_some_and(|value| !value.is_null() && value != "") {
            return Err("unsupported pane part: top_left_cell".into());
        }
        let sheet_name = required_string(spec, "sheet")?;
        let row = optional_u32(spec, "y_split").unwrap_or(0);
        let col = optional_u16(spec, "x_split").unwrap_or(0);
        workbook.worksheet_by_name(sheet_name).map_err(|error| error.to_string())?
            .set_freeze_panes(row, col).map_err(|error| error.to_string())?;
    }
    Ok(())
}

fn apply_named_ranges(workbook: &mut Workbook, payload: &Value) -> Result<(), String> {
    for entry in array(payload, "named_ranges")? {
        let spec = object(entry, "named range entry")?;
        let name = required_string(spec, "name")?;
        let formula = required_string(spec, "refers_to")?;
        if optional_string(spec, "scope") == Some("sheet") {
            let sheet_name = required_string(spec, "sheet")?;
            let sheet_index = workbook.sheet_names().iter().position(|candidate| *candidate == sheet_name)
                .ok_or_else(|| format!("named range sheet {sheet_name:?} does not exist"))?;
            workbook.define_name_scoped(name, formula, sheet_index);
        } else {
            workbook.define_name(name, formula);
        }
    }
    Ok(())
}

fn apply_tables(workbook: &mut Workbook, payload: &Value) -> Result<(), String> {
    for entry in array(payload, "tables")? {
        let spec = object(entry, "table entry")?;
        if optional_bool(spec, "show_header_row") == Some(false) {
            return Err("unsupported table part: show_header_row=false".into());
        }
        if optional_bool(spec, "show_row_stripes") == Some(false) {
            return Err("unsupported table part: show_row_stripes=false".into());
        }
        let sheet_name = required_string(spec, "sheet")?;
        let range = required_string(spec, "range")?;
        let (first_row, first_col, last_row, last_col) = parse_range(range)?;
        let mut table = Table::new();
        if let Some(name) = optional_string(spec, "name") { table.set_name(name); }
        if let Some(style) = optional_string(spec, "style") { table.set_style(parse_table_style(style)?); }
        table.set_total_row(optional_bool(spec, "totals_row").unwrap_or(false));
        table.set_autofilter(optional_bool(spec, "autofilter").unwrap_or(true));
        let columns: Result<Vec<TableColumn>, String> = object_array(spec, "columns")?.iter()
            .map(|value| value.as_str().map(TableColumn::new).ok_or_else(|| "table columns must be strings".to_string()))
            .collect();
        table.set_columns(&columns?);
        workbook.worksheet_by_name(sheet_name).map_err(|error| error.to_string())?
            .add_table(first_row, first_col, last_row, last_col, &table)
            .map_err(|error| format!("add table {sheet_name}!{range}: {error}"))?;
    }
    Ok(())
}

fn apply_charts(workbook: &mut Workbook, payload: &Value) -> Result<(), String> {
    for entry in array(payload, "charts")? {
        let spec = object(entry, "chart entry")?;
        if spec.get("vary_colors").is_some_and(|value| !value.is_null()) {
            return Err("unsupported chart part: vary_colors".into());
        }
        let sheet_name = required_string(spec, "sheet")?;
        let address = required_string(spec, "cell")?;
        let chart_type = parse_chart_type(required_string(spec, "type")?)?;
        let (row, col) = parse_cell(address)?;
        let mut chart = Chart::new(chart_type);
        if let Some(title) = optional_string(spec, "title") { chart.set_title(title); }
        if let Some(width) = optional_u32(spec, "width") { chart.set_width(width); }
        if let Some(height) = optional_u32(spec, "height") { chart.set_height(height); }
        if let Some(alt_text) = optional_string(spec, "alt_text") { chart.set_alt_text(optional_string(spec, "name").unwrap_or("Chart"), alt_text); }
        let series = object_array(spec, "series")?;
        if series.is_empty() {
            let values = required_string(spec, "values")?;
            let item = chart.add_series();
            item.set_values(values);
            if let Some(categories) = optional_string(spec, "categories") { item.set_categories(categories); }
            if let Some(name) = optional_string(spec, "name") { item.set_name(name); }
            item.set_data_labels(optional_bool(spec, "show_values").unwrap_or(false));
        } else {
            for series_spec in series {
                let series_object = object(series_spec, "chart series")?;
                let item = chart.add_series();
                item.set_values(required_string(series_object, "values")?);
                if let Some(categories) = optional_string(series_object, "categories") { item.set_categories(categories); }
                if let Some(name) = optional_string(series_object, "name") { item.set_name(name); }
                if let Some(color) = optional_string(series_object, "fill_color") { item.set_color(color); }
                item.set_data_labels(optional_bool(spec, "show_values").unwrap_or(false));
                for point in object_array(series_object, "data_points")? {
                    let point = object(point, "chart data point")?;
                    if let Some(color) = optional_string(point, "fill_color") {
                        item.set_point_color(required_u32(point, "index")? as usize, color);
                    }
                }
            }
        }
        workbook.worksheet_by_name(sheet_name).map_err(|error| error.to_string())?
            .insert_chart(row, col, &chart).map_err(|error| error.to_string())?;
    }
    Ok(())
}

fn ensure_supported_payload(payload: &Value) -> Result<(), String> {
    let root = object(payload, "write payload")?;
    for part in ["validations", "conditional_formats", "pivots", "slicers", "pictures"] {
        if !object_array(root, part)?.is_empty() {
            return Err(format!("unsupported payload part: {part}"));
        }
    }
    Ok(())
}

fn payload_counts(payload: &Value) -> Value {
    let mut counts = Map::new();
    for name in [
        "sheets", "cells", "formats", "borders", "merges", "row_heights", "validations",
        "hyperlinks", "comments", "panes", "named_ranges", "tables", "conditional_formats",
        "charts", "pivots", "slicers", "pictures",
    ] {
        let count = payload.get(name).and_then(Value::as_array).map_or(0, Vec::len);
        counts.insert(name.into(), json!(count));
    }
    Value::Object(counts)
}

fn object<'a>(value: &'a Value, context: &str) -> Result<&'a Map<String, Value>, String> {
    value.as_object().ok_or_else(|| format!("{context} must be an object"))
}

fn array<'a>(value: &'a Value, name: &str) -> Result<Vec<&'a Value>, String> {
    match value.get(name) {
        None | Some(Value::Null) => Ok(Vec::new()),
        Some(Value::Array(values)) => Ok(values.iter().collect()),
        Some(_) => Err(format!("{name} must be an array")),
    }
}

fn object_array<'a>(
    object: &'a Map<String, Value>,
    name: &str,
) -> Result<Vec<&'a Value>, String> {
    match object.get(name) {
        None | Some(Value::Null) => Ok(Vec::new()),
        Some(Value::Array(values)) => Ok(values.iter().collect()),
        Some(_) => Err(format!("{name} must be an array")),
    }
}

fn required_string<'a>(object: &'a Map<String, Value>, name: &str) -> Result<&'a str, String> {
    optional_string(object, name).filter(|value| !value.is_empty())
        .ok_or_else(|| format!("{name} is required"))
}

fn optional_string<'a>(object: &'a Map<String, Value>, name: &str) -> Option<&'a str> {
    object.get(name).and_then(Value::as_str)
}

fn required_number(object: &Map<String, Value>, name: &str) -> Result<f64, String> {
    optional_number(object, name).ok_or_else(|| format!("{name} must be a number"))
}

fn optional_number(object: &Map<String, Value>, name: &str) -> Option<f64> {
    object.get(name).and_then(Value::as_f64)
}

fn required_u32(object: &Map<String, Value>, name: &str) -> Result<u32, String> {
    optional_u32(object, name).ok_or_else(|| format!("{name} must be a positive integer"))
}

fn optional_u32(object: &Map<String, Value>, name: &str) -> Option<u32> {
    object.get(name).and_then(Value::as_u64).and_then(|value| u32::try_from(value).ok())
}

fn optional_u16(object: &Map<String, Value>, name: &str) -> Option<u16> {
    object.get(name).and_then(Value::as_u64).and_then(|value| u16::try_from(value).ok())
}

fn optional_bool(object: &Map<String, Value>, name: &str) -> Option<bool> {
    object.get(name).and_then(Value::as_bool)
}

fn required_path<'a>(path: Option<&'a str>, message: &str) -> Result<&'a str, String> {
    path.filter(|value| !value.is_empty()).ok_or_else(|| message.to_string())
}

fn create_parent(path: &str) -> Result<(), String> {
    let parent = Path::new(path).parent().unwrap_or_else(|| Path::new("."));
    fs::create_dir_all(parent).map_err(|error| format!("create output directory: {error}"))
}

fn parse_cell(value: &str) -> Result<(u32, u16), String> {
    let split = value.find(|character: char| character.is_ascii_digit())
        .ok_or_else(|| format!("invalid cell reference {value:?}"))?;
    let (column, row) = value.split_at(split);
    if column.is_empty() || row.is_empty() || !column.chars().all(|character| character.is_ascii_alphabetic()) {
        return Err(format!("invalid cell reference {value:?}"));
    }
    let row = row.parse::<u32>().map_err(|_| format!("invalid cell reference {value:?}"))?
        .checked_sub(1).ok_or_else(|| format!("invalid cell reference {value:?}"))?;
    Ok((row, parse_column(column)?))
}

fn parse_column(value: &str) -> Result<u16, String> {
    if value.is_empty() || !value.chars().all(|character| character.is_ascii_alphabetic()) {
        return Err(format!("invalid column {value:?}"));
    }
    let mut result = 0u32;
    for character in value.bytes() {
        result = result.checked_mul(26).and_then(|current| current.checked_add(u32::from(character.to_ascii_uppercase() - b'A' + 1)))
            .ok_or_else(|| format!("invalid column {value:?}"))?;
    }
    u16::try_from(result.checked_sub(1).ok_or_else(|| format!("invalid column {value:?}"))?)
        .map_err(|_| format!("column out of range {value:?}"))
}

fn parse_range(value: &str) -> Result<(u32, u16, u32, u16), String> {
    let (start, end) = value.split_once(':').ok_or_else(|| format!("invalid range {value:?}"))?;
    let (first_row, first_col) = parse_cell(start)?;
    let (last_row, last_col) = parse_cell(end)?;
    if first_row > last_row || first_col > last_col { return Err(format!("invalid range {value:?}")); }
    Ok((first_row, first_col, last_row, last_col))
}

fn cell_name(row: u32, col: u16) -> String {
    let mut column = u32::from(col) + 1;
    let mut letters = String::new();
    while column > 0 {
        let remainder = (column - 1) % 26;
        letters.insert(0, char::from(b'A' + u8::try_from(remainder).unwrap_or(0)));
        column = (column - 1) / 26;
    }
    format!("{}{}", letters, row + 1)
}

fn border_style(value: &str) -> Result<BorderStyle, String> {
    match value {
        "" | "none" => Ok(BorderStyle::None),
        "thin" => Ok(BorderStyle::Thin),
        "medium" => Ok(BorderStyle::Medium),
        "thick" => Ok(BorderStyle::Thick),
        "dashed" => Ok(BorderStyle::Dashed),
        "dotted" => Ok(BorderStyle::Dotted),
        "double" => Ok(BorderStyle::Double),
        _ => Err(format!("unsupported border style {value:?}")),
    }
}

fn horizontal_alignment(value: &str) -> Result<Align, String> {
    match value {
        "" | "general" | "left" => Ok(Align::Left),
        "center" => Ok(Align::Center),
        "right" => Ok(Align::Right),
        "fill" => Ok(Align::Fill),
        "justify" => Ok(Align::Justify),
        _ => Err(format!("unsupported horizontal alignment {value:?}")),
    }
}

fn vertical_alignment(value: &str) -> Result<Align, String> {
    match value {
        "" | "bottom" => Ok(Align::Bottom),
        "top" => Ok(Align::Top),
        "center" => Ok(Align::VerticalCenter),
        _ => Err(format!("unsupported vertical alignment {value:?}")),
    }
}

fn parse_table_style(value: &str) -> Result<TableStyle, String> {
    for (prefix, build) in [
        ("TableStyleLight", TableStyle::Light as fn(u8) -> TableStyle),
        ("TableStyleMedium", TableStyle::Medium as fn(u8) -> TableStyle),
        ("TableStyleDark", TableStyle::Dark as fn(u8) -> TableStyle),
    ] {
        if let Some(number) = value.strip_prefix(prefix) {
            return number.parse::<u8>().map(build).map_err(|_| format!("unsupported table style {value:?}"));
        }
    }
    Err(format!("unsupported table style {value:?}"))
}

fn parse_chart_type(value: &str) -> Result<ChartType, String> {
    match value {
        "area" => Ok(ChartType::Area),
        "bar" => Ok(ChartType::Bar),
        "col" | "column" => Ok(ChartType::Column),
        "line" => Ok(ChartType::Line),
        "pie" => Ok(ChartType::Pie),
        "doughnut" => Ok(ChartType::Doughnut),
        "scatter" => Ok(ChartType::Scatter),
        "bubble" => Ok(ChartType::Bubble),
        _ => Err(format!("unsupported chart type {value:?}")),
    }
}
