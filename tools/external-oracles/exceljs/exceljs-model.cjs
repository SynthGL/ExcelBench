// ExcelBench JSON-model operations for ExcelJS: describe, read_model,
// write_model, mutate. Every value reported here comes from the ExcelJS object
// model after `workbook.xlsx.readFile`; every write goes through the public
// ExcelJS API. Behavior ExcelJS loses or cannot express is left to fail the
// corresponding ExcelBench test case.
const fs = require("node:fs/promises");
const path = require("node:path");

const ExcelJS = require("exceljs");
const EXCELJS_VERSION = require("exceljs/package.json").version;

const { ValueType } = ExcelJS;

const UNSUPPORTED_READ = {
  pivot_tables: "ExcelJS 4.4.0 does not load pivot table parts; no pivot table read API",
  chart_anchors: "ExcelJS 4.4.0 does not load chart parts; no chart read API",
};

const UNSUPPORTED_WRITE = {
  pivot: "ExcelJS 4.4.0 has no pivot table creation API",
  chart: "ExcelJS 4.4.0 has no chart creation API",
};

// Error literal -> formula that produces it (contract write table).
const ERROR_FORMULAS = {
  "#DIV/0!": "1/0",
  "#N/A": "NA()",
  "#VALUE!": '"text"+1',
  "#REF!": "#REF!",
  "#NAME?": "_undefined_name_",
  "#NUM!": "SQRT(-1)",
  "#NULL!": "A1:A2 B1:B2",
};

function describe() {
  return {
    library: "exceljs",
    version: EXCELJS_VERSION,
    capabilities: ["read", "write", "modify"],
    unsupported_write: { ...UNSUPPORTED_WRITE },
  };
}

// ===========================================================================
// read_model
// ===========================================================================

async function readModel(request) {
  const inputPath = requirePath(request, "input_path", "read_model");
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(inputPath);
  return {
    model: {
      unsupported: { ...UNSUPPORTED_READ },
      named_ranges: readNamedRanges(workbook),
      sheets: workbook.worksheets.map((worksheet) => readSheet(workbook, worksheet)),
    },
  };
}

function readNamedRanges(workbook) {
  // ExcelJS keeps only {name, ranges}; localSheetId is discarded on load, so
  // every name it exposes is workbook-scoped.
  return workbook.definedNames.model.map((definedName) => ({
    name: definedName.name,
    scope: "workbook",
    sheet: null,
    refers_to: definedName.ranges.join(","),
  }));
}

function readSheet(workbook, worksheet) {
  const sheetModel = worksheet.model;
  const cells = {};
  const rowHeights = {};
  const hyperlinks = [];
  const comments = [];

  worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    if (typeof row.height === "number") {
      rowHeights[String(rowNumber)] = row.height;
    }
    row.eachCell({ includeEmpty: true }, (cell) => {
      const entry = {};
      const value = readCellValue(cell);
      if (value) entry.value = value;
      const format = readCellFormat(cell);
      if (Object.keys(format).length) entry.format = format;
      const border = readCellBorder(cell);
      if (Object.keys(border).length) entry.border = border;
      if (Object.keys(entry).length) cells[cell.address] = entry;

      if (cell.type === ValueType.Hyperlink && cell.hyperlink) {
        hyperlinks.push({
          cell: cell.address,
          target: cell.hyperlink,
          display: cell.text,
          tooltip: cell.value.tooltip ?? null,
          // ExcelJS only reconciles relationship-backed (external) hyperlinks;
          // location-only links are dropped on load.
          internal: false,
        });
      }
      if (cell.note) {
        comments.push({
          cell: cell.address,
          text: noteText(cell.note),
          author: null,
          threaded: false,
        });
      }
    });
  });

  const columnWidths = {};
  for (const column of worksheet.columns || []) {
    if (typeof column.width === "number") {
      columnWidths[column.letter] = column.width;
    }
  }

  return {
    name: worksheet.name,
    cells,
    row_heights: rowHeights,
    column_widths: columnWidths,
    merged_ranges: [...(sheetModel.merges || [])],
    conditional_formats: readConditionalFormats(worksheet),
    data_validations: readDataValidations(worksheet),
    hyperlinks,
    images: readImages(workbook, worksheet),
    pivot_tables: [],
    comments,
    freeze_panes: readFreezePanes(worksheet),
    tables: readTables(worksheet),
    sheet_protection: readSheetProtection(worksheet),
    page_setup: readPageSetup(worksheet),
    chart_anchors: [],
  };
}

function readCellValue(cell) {
  const raw = cell.value;
  switch (cell.type) {
    case ValueType.Null:
    case ValueType.Merge:
      // Merge cells are ExcelJS placeholders that proxy the master cell value.
      return null;
    case ValueType.Number:
      return { type: "number", value: raw, formula: null };
    case ValueType.String:
    case ValueType.SharedString:
    case ValueType.RichText:
    case ValueType.Hyperlink:
      return { type: "string", value: cell.text, formula: null };
    case ValueType.Boolean:
      return { type: "boolean", value: raw, formula: null };
    case ValueType.Date:
      return dateValue(raw);
    case ValueType.Error:
      return { type: "error", value: raw.error, formula: null };
    case ValueType.Formula: {
      const result = cell.result;
      if (result && typeof result === "object" && typeof result.error === "string") {
        return { type: "error", value: result.error, formula: null };
      }
      const text = `=${cell.formula}`;
      return { type: "formula", value: text, formula: text };
    }
    default:
      return null;
  }
}

function dateValue(date) {
  const iso = date.toISOString();
  const midnight =
    date.getUTCHours() === 0 &&
    date.getUTCMinutes() === 0 &&
    date.getUTCSeconds() === 0 &&
    date.getUTCMilliseconds() === 0;
  return midnight
    ? { type: "date", value: iso.slice(0, 10), formula: null }
    : { type: "datetime", value: iso.slice(0, 19), formula: null };
}

function readCellFormat(cell) {
  const format = {};
  const font = cell.font || {};
  if (font.bold) format.bold = true;
  if (font.italic) format.italic = true;
  if (font.underline) {
    format.underline = font.underline === true ? "single" : String(font.underline);
  }
  if (font.strike) format.strikethrough = true;
  if (font.name) format.font_name = font.name;
  if (typeof font.size === "number") format.font_size = font.size;
  const fontColor = argbToHex(font.color);
  if (fontColor) format.font_color = fontColor;

  const fill = cell.fill;
  if (fill && fill.type === "pattern" && fill.pattern === "solid") {
    const bg = argbToHex(fill.fgColor);
    if (bg) format.bg_color = bg;
  }
  if (cell.numFmt) format.number_format = cell.numFmt;

  const alignment = cell.alignment || {};
  if (alignment.horizontal) format.h_align = alignment.horizontal;
  if (alignment.vertical) {
    // ExcelJS names OOXML vertical="center" as "middle".
    format.v_align = alignment.vertical === "middle" ? "center" : alignment.vertical;
  }
  if (alignment.wrapText) format.wrap = true;
  if (alignment.textRotation) {
    // ExcelJS names OOXML textRotation="255" as "vertical".
    format.rotation = alignment.textRotation === "vertical" ? 255 : alignment.textRotation;
  }
  if (alignment.indent) format.indent = alignment.indent;
  return format;
}

function readCellBorder(cell) {
  const border = cell.border || {};
  const out = {};
  for (const edge of ["top", "bottom", "left", "right"]) {
    const item = borderEdge(border[edge]);
    if (item) out[edge] = item;
  }
  const diagonal = border.diagonal;
  if (diagonal) {
    const item = borderEdge(diagonal);
    if (item && diagonal.up) out.diagonal_up = item;
    if (item && diagonal.down) out.diagonal_down = item;
  }
  return out;
}

function borderEdge(side) {
  if (!side || !side.style) return null;
  const item = { style: side.style };
  const color = argbToHex(side.color);
  if (color) item.color = color;
  return item;
}

function readConditionalFormats(worksheet) {
  const out = [];
  for (const cf of worksheet.conditionalFormattings || []) {
    for (const rule of cf.rules || []) {
      const style = rule.style || {};
      const format = {};
      const fill = style.fill;
      if (fill) {
        // Differential fills carry the solid colour in bgColor (fgColor mirrors it).
        const bg = argbToHex(fill.bgColor) || argbToHex(fill.fgColor);
        if (bg) format.bg_color = bg;
      }
      const fontColor = style.font ? argbToHex(style.font.color) : null;
      if (fontColor) format.font_color = fontColor;
      const formulae = rule.formulae || [];
      out.push({
        range: cf.ref,
        rule_type: rule.type ?? null,
        operator: rule.operator ?? null,
        formula: formulae.length ? String(formulae[0]) : null,
        priority: typeof rule.priority === "number" ? rule.priority : null,
        stop_if_true: typeof rule.stopIfTrue === "boolean" ? rule.stopIfTrue : null,
        format,
      });
    }
  }
  return out;
}

function readDataValidations(worksheet) {
  const out = [];
  for (const [address, dv] of Object.entries(worksheet.dataValidations.model || {})) {
    if (!dv) continue;
    const formulae = dv.formulae || [];
    const formula1 = formulae.length > 0 ? formulaText(formulae[0]) : null;
    const formula2 = formulae.length > 1 ? formulaText(formulae[1]) : null;
    let operator = dv.operator ?? null;
    if (operator === null && formula2 !== null) operator = "between";
    out.push({
      range: address.replace(/^range:/, ""),
      validation_type: dv.type ?? null,
      operator,
      formula1,
      formula2,
      allow_blank: typeof dv.allowBlank === "boolean" ? dv.allowBlank : null,
      show_input: typeof dv.showInputMessage === "boolean" ? dv.showInputMessage : null,
      show_error: typeof dv.showErrorMessage === "boolean" ? dv.showErrorMessage : null,
      prompt_title: dv.promptTitle ?? null,
      prompt: dv.prompt ?? null,
      error_title: dv.errorTitle ?? null,
      error: dv.error ?? null,
    });
  }
  return out;
}

function formulaText(value) {
  if (value === null || value === undefined) return null;
  if (value instanceof Date) return value.toISOString();
  return String(value);
}

function readImages(workbook, worksheet) {
  return worksheet.getImages().map((image) => {
    const range = image.range || {};
    const tl = range.tl;
    const medium = workbook.getImage(image.imageId);
    return {
      cell: tl ? `${columnLetter(tl.nativeCol + 1)}${tl.nativeRow + 1}` : null,
      path: medium && medium.name ? `/xl/media/${medium.name}.${medium.extension}` : null,
      anchor: tl ? (range.br ? "twoCell" : "oneCell") : null,
      offset: tl ? [tl.nativeColOff, tl.nativeRowOff] : null,
      alt_text: null,
    };
  });
}

function readFreezePanes(worksheet) {
  const view = (worksheet.views || [])[0];
  if (!view) return {};
  if (view.state === "frozen") {
    const out = { mode: "freeze" };
    if (view.topLeftCell) out.top_left_cell = view.topLeftCell;
    return out;
  }
  if (view.state === "split") {
    const out = {
      mode: "split",
      x_split: typeof view.xSplit === "number" ? view.xSplit : null,
      y_split: typeof view.ySplit === "number" ? view.ySplit : null,
    };
    if (view.topLeftCell) out.top_left_cell = view.topLeftCell;
    if (view.activePane) out.active_pane = view.activePane;
    return out;
  }
  return {};
}

function readTables(worksheet) {
  return worksheet.getTables().map((table) => {
    const model = table.model || {};
    const style = model.style || {};
    return {
      name: model.name ?? null,
      ref: model.tableRef ?? model.ref ?? null,
      header_row: Boolean(model.headerRow),
      totals_row: Boolean(model.totalsRow),
      style: style.theme ?? null,
      columns: (model.columns || []).map((column) => column.name),
      autofilter: Boolean(model.autoFilterRef),
    };
  });
}

function readSheetProtection(worksheet) {
  const sp = worksheet.sheetProtection;
  if (!sp) {
    return {
      protected: false,
      password_hash_present: false,
      format_cells: null,
      insert_rows: null,
      select_locked_cells: null,
      select_unlocked_cells: null,
      sort: null,
      auto_filter: null,
    };
  }
  // ExcelJS models protection options as "allowed" flags whose defaults
  // (absent attribute) match OOXML: formatCells/insertRows/sort/autoFilter are
  // disallowed unless true, select*Cells are allowed unless false. The
  // contract uses raw OOXML semantics (true == blocked).
  return {
    protected: sp.sheet === true,
    password_hash_present: Boolean(sp.hashValue),
    format_cells: sp.formatCells !== true,
    insert_rows: sp.insertRows !== true,
    select_locked_cells: sp.selectLockedCells === false,
    select_unlocked_cells: sp.selectUnlockedCells === false,
    sort: sp.sort !== true,
    auto_filter: sp.autoFilter !== true,
  };
}

function readPageSetup(worksheet) {
  const ps = worksheet.pageSetup || {};
  const hf = worksheet.headerFooter || {};
  return {
    orientation: ps.orientation ?? null,
    fit_to_width: typeof ps.fitToWidth === "number" ? ps.fitToWidth : null,
    fit_to_height: typeof ps.fitToHeight === "number" ? ps.fitToHeight : null,
    scale: typeof ps.scale === "number" ? ps.scale : null,
    // ExcelJS stores print title rows as "1:2" (no "$").
    print_title_rows: ps.printTitlesRow ? absoluteRowRange(ps.printTitlesRow) : null,
    header_center: centerSection(hf.oddHeader),
    footer_center: centerSection(hf.oddFooter),
  };
}

function absoluteRowRange(value) {
  return String(value)
    .split(":")
    .map((part) => `$${part}`)
    .join(":");
}

// Split an OOXML header/footer string into its &L/&C/&R sections.
function centerSection(text) {
  if (!text) return null;
  const sections = {};
  let current = "C";
  let buffer = "";
  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    if (ch === "&" && i + 1 < text.length) {
      const next = text[i + 1];
      if (next === "L" || next === "C" || next === "R") {
        sections[current] = (sections[current] || "") + buffer;
        buffer = "";
        current = next;
        i += 1;
        continue;
      }
      buffer += ch + next;
      i += 1;
      continue;
    }
    buffer += ch;
  }
  sections[current] = (sections[current] || "") + buffer;
  return sections.C ? sections.C : null;
}

// ===========================================================================
// write_model
// ===========================================================================

async function writeModel(request) {
  const outputPath = requirePath(request, "output_path", "write_model");
  const ops = (request.payload && request.payload.ops) || [];
  const workbook = new ExcelJS.Workbook();

  for (const op of ops) {
    if (op.op === "add_sheet") workbook.addWorksheet(op.name);
  }
  for (const op of ops) {
    await applyOp(workbook, op);
  }

  await fs.mkdir(path.dirname(path.resolve(outputPath)), { recursive: true });
  await workbook.xlsx.writeFile(outputPath);
  return { written: outputPath };
}

async function applyOp(workbook, op) {
  switch (op.op) {
    case "add_sheet":
      return;
    case "cell_value":
      writeCellValue(sheetOf(workbook, op.sheet).getCell(op.cell), op);
      return;
    case "cell_format":
      writeCellFormat(sheetOf(workbook, op.sheet).getCell(op.cell), op.format || {});
      return;
    case "cell_border":
      writeCellBorder(sheetOf(workbook, op.sheet).getCell(op.cell), op.border || {});
      return;
    case "row_height":
      sheetOf(workbook, op.sheet).getRow(Number(op.row)).height = Number(op.height);
      return;
    case "column_width":
      sheetOf(workbook, op.sheet).getColumn(op.column).width = Number(op.width);
      return;
    case "merge":
      sheetOf(workbook, op.sheet).mergeCells(op.range);
      return;
    case "conditional_format":
      writeConditionalFormat(sheetOf(workbook, op.sheet), op.rule || {});
      return;
    case "data_validation":
      writeDataValidation(sheetOf(workbook, op.sheet), op.validation || {});
      return;
    case "hyperlink":
      writeHyperlink(sheetOf(workbook, op.sheet), op.link || {});
      return;
    case "image":
      await writeImage(workbook, sheetOf(workbook, op.sheet), op.image || {});
      return;
    case "comment": {
      const comment = op.comment || {};
      // ExcelJS notes have no author field; only the text is expressible.
      sheetOf(workbook, op.sheet).getCell(comment.cell).note = String(comment.text ?? "");
      return;
    }
    case "freeze":
      writeFreeze(sheetOf(workbook, op.sheet), op.settings || {});
      return;
    case "named_range":
      writeNamedRange(workbook, op.named_range || {});
      return;
    case "table":
      writeTable(sheetOf(workbook, op.sheet), op.table || {});
      return;
    case "protection":
      await writeProtection(sheetOf(workbook, op.sheet), op.settings || {});
      return;
    case "page_setup":
      writePageSetup(sheetOf(workbook, op.sheet), op.settings || {});
      return;
    case "pivot":
    case "chart":
      throw new Error(`${op.op}: ${UNSUPPORTED_WRITE[op.op]}`);
    default:
      throw new Error(`Unknown write op '${op.op}'.`);
  }
}

function writeCellValue(cell, op) {
  const { type, value } = op;
  switch (type) {
    case "blank":
      cell.value = null;
      return;
    case "string":
      cell.value = value === null || value === undefined ? "" : String(value);
      return;
    case "number":
      cell.value = Number(value);
      return;
    case "boolean":
      cell.value = Boolean(value);
      return;
    case "date":
      cell.value = new Date(`${String(value).slice(0, 10)}T00:00:00Z`);
      return;
    case "datetime":
      cell.value = new Date(`${String(value).slice(0, 19)}Z`);
      return;
    case "formula": {
      const text = String(op.formula ?? value ?? "");
      cell.value = { formula: text.replace(/^=/, "") };
      return;
    }
    case "error": {
      const formula = ERROR_FORMULAS[String(value).toUpperCase()];
      if (!formula) throw new Error(`No formula produces error literal '${value}'.`);
      cell.value = { formula };
      return;
    }
    default:
      throw new Error(`Unsupported cell value type '${type}'.`);
  }
}

function writeCellFormat(cell, format) {
  const font = {};
  if (format.bold !== undefined) font.bold = Boolean(format.bold);
  if (format.italic !== undefined) font.italic = Boolean(format.italic);
  if (format.underline !== undefined) font.underline = format.underline;
  if (format.strikethrough !== undefined) font.strike = Boolean(format.strikethrough);
  if (format.font_name !== undefined) font.name = format.font_name;
  if (format.font_size !== undefined) font.size = Number(format.font_size);
  if (format.font_color !== undefined) font.color = { argb: hexToArgb(format.font_color) };
  if (Object.keys(font).length) cell.font = { ...(cell.font || {}), ...font };

  if (format.bg_color !== undefined) {
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: hexToArgb(format.bg_color) },
    };
  }
  if (format.number_format !== undefined) cell.numFmt = format.number_format;

  const alignment = {};
  if (format.h_align !== undefined) alignment.horizontal = format.h_align;
  if (format.v_align !== undefined) {
    alignment.vertical = format.v_align === "center" ? "middle" : format.v_align;
  }
  if (format.wrap !== undefined) alignment.wrapText = Boolean(format.wrap);
  if (format.rotation !== undefined) {
    alignment.textRotation = Number(format.rotation) === 255 ? "vertical" : Number(format.rotation);
  }
  if (format.indent !== undefined) alignment.indent = Number(format.indent);
  if (Object.keys(alignment).length) {
    cell.alignment = { ...(cell.alignment || {}), ...alignment };
  }
}

function writeCellBorder(cell, border) {
  const out = {};
  for (const edge of ["top", "bottom", "left", "right"]) {
    if (border[edge]) out[edge] = borderSpec(border[edge]);
  }
  const up = border.diagonal_up;
  const down = border.diagonal_down;
  if (up || down) {
    out.diagonal = { ...borderSpec(up || down), up: Boolean(up), down: Boolean(down) };
  }
  cell.border = out;
}

function borderSpec(edge) {
  const spec = { style: edge.style };
  if (edge.color) spec.color = { argb: hexToArgb(edge.color) };
  return spec;
}

function writeConditionalFormat(worksheet, rule) {
  const cfRule = { type: rule.rule_type };
  if (typeof rule.priority === "number") cfRule.priority = rule.priority;
  switch (rule.rule_type) {
    case "cellIs":
      cfRule.operator = rule.operator;
      cfRule.formulae = rule.formula ? [stripEquals(rule.formula)] : [];
      break;
    case "expression":
      cfRule.formulae = rule.formula ? [stripEquals(rule.formula)] : [];
      break;
    case "dataBar":
      cfRule.cfvo = [{ type: "min" }, { type: "max" }];
      cfRule.color = { argb: "FF638EC6" };
      break;
    case "colorScale":
      cfRule.cfvo = [{ type: "min" }, { type: "percentile", value: 50 }, { type: "max" }];
      cfRule.color = [{ argb: "FFAA0000" }, { argb: "FFFFFF00" }, { argb: "FF00AA00" }];
      break;
    default:
      if (rule.operator) cfRule.operator = rule.operator;
      if (rule.formula) cfRule.formulae = [stripEquals(rule.formula)];
  }
  // ExcelJS has no stopIfTrue rule attribute; the flag cannot be expressed.
  const format = rule.format || {};
  const style = {};
  if (format.bg_color) {
    const argb = hexToArgb(format.bg_color);
    style.fill = { type: "pattern", pattern: "solid", fgColor: { argb }, bgColor: { argb } };
  }
  if (format.font_color) style.font = { color: { argb: hexToArgb(format.font_color) } };
  if (Object.keys(style).length) cfRule.style = style;
  worksheet.addConditionalFormatting({ ref: rule.range, rules: [cfRule] });
}

function writeDataValidation(worksheet, validation) {
  const dv = { type: validation.validation_type };
  if (validation.operator) dv.operator = validation.operator;
  const formulae = [];
  if (validation.formula1 !== null && validation.formula1 !== undefined) {
    formulae.push(stripEquals(validation.formula1));
  }
  if (validation.formula2 !== null && validation.formula2 !== undefined) {
    formulae.push(stripEquals(validation.formula2));
  }
  dv.formulae = formulae;
  if (validation.allow_blank !== null && validation.allow_blank !== undefined) {
    dv.allowBlank = Boolean(validation.allow_blank);
  }
  if (validation.show_input !== null && validation.show_input !== undefined) {
    dv.showInputMessage = Boolean(validation.show_input);
  }
  if (validation.show_error !== null && validation.show_error !== undefined) {
    dv.showErrorMessage = Boolean(validation.show_error);
  }
  if (validation.prompt_title) dv.promptTitle = validation.prompt_title;
  if (validation.prompt) dv.prompt = validation.prompt;
  if (validation.error_title) dv.errorTitle = validation.error_title;
  if (validation.error) dv.error = validation.error;
  worksheet.dataValidations.add(validation.range, dv);
}

function writeHyperlink(worksheet, link) {
  const target = link.internal
    ? String(link.target).replace(/^#/, "")
    : String(link.target);
  const value = { text: String(link.display ?? target), hyperlink: target };
  if (link.tooltip) value.tooltip = link.tooltip;
  worksheet.getCell(link.cell).value = value;
}

async function writeImage(workbook, worksheet, image) {
  const filename = image.path;
  if (!filename) throw new Error("image op requires an absolute path.");
  const buffer = await fs.readFile(filename);
  let extension = path.extname(filename).slice(1).toLowerCase();
  if (extension === "jpg") extension = "jpeg";
  const imageId = workbook.addImage({ buffer, extension });
  const { col, row } = decodeCell(image.cell);
  if (image.anchor === "twoCell") {
    worksheet.addImage(imageId, `${image.cell}:${image.cell}`);
    return;
  }
  // ExcelJS writes a oneCellAnchor when the range has tl + ext (pixels) and no br.
  const { width, height } = imagePixelSize(buffer, extension);
  worksheet.addImage(imageId, { tl: { col: col - 1, row: row - 1 }, ext: { width, height } });
}

function writeFreeze(worksheet, settings) {
  if (settings.mode === "freeze") {
    const { col, row } = decodeCell(settings.top_left_cell);
    worksheet.views = [
      { state: "frozen", xSplit: col - 1, ySplit: row - 1, topLeftCell: settings.top_left_cell },
    ];
    return;
  }
  if (settings.mode === "split") {
    const view = {
      state: "split",
      xSplit: Number(settings.x_split || 0),
      ySplit: Number(settings.y_split || 0),
    };
    if (settings.top_left_cell) view.topLeftCell = settings.top_left_cell;
    if (settings.active_pane) view.activePane = settings.active_pane;
    worksheet.views = [view];
    return;
  }
  throw new Error(`Unknown freeze mode '${settings.mode}'.`);
}

function writeNamedRange(workbook, namedRange) {
  // ExcelJS DefinedNames.add(location, name) has no scope parameter; every
  // name it writes is workbook-scoped.
  workbook.definedNames.add(stripEquals(String(namedRange.refers_to)), namedRange.name);
}

function writeTable(worksheet, table) {
  const range = decodeRange(table.ref);
  const headerRow = table.header_row !== false;
  const totalsRow = Boolean(table.totals_row);
  const firstDataRow = range.top + (headerRow ? 1 : 0);
  const lastDataRow = range.bottom - (totalsRow ? 1 : 0);
  // ExcelJS sizes a table from its rows and rewrites them into the sheet, so
  // pass the existing cell values inside the ref as the table rows.
  const rows = [];
  for (let r = firstDataRow; r <= lastDataRow; r += 1) {
    const values = [];
    for (let c = range.left; c <= range.right; c += 1) {
      values.push(worksheet.getCell(r, c).value);
    }
    rows.push(values);
  }
  const columns = (table.columns || []).map((name) => ({
    name: String(name),
    filterButton: table.autofilter !== false,
  }));
  worksheet.addTable({
    name: table.name,
    ref: `${columnLetter(range.left)}${range.top}`,
    headerRow,
    totalsRow,
    style: { theme: table.style ?? null },
    columns,
    rows,
  });
}

async function writeProtection(worksheet, settings) {
  if (settings.protected === false) {
    worksheet.unprotect();
    return;
  }
  // Translate raw OOXML "blocked" flags into ExcelJS "allowed" options.
  const options = {};
  const allowed = (key, flag) => {
    if (settings[flag] !== null && settings[flag] !== undefined) {
      options[key] = !settings[flag];
    }
  };
  allowed("formatCells", "format_cells");
  allowed("insertRows", "insert_rows");
  allowed("selectLockedCells", "select_locked_cells");
  allowed("selectUnlockedCells", "select_unlocked_cells");
  allowed("sort", "sort");
  allowed("autoFilter", "auto_filter");
  await worksheet.protect(settings.password || "", options);
}

function writePageSetup(worksheet, settings) {
  if (settings.orientation) worksheet.pageSetup.orientation = settings.orientation;
  const fitWidth = settings.fit_to_width;
  const fitHeight = settings.fit_to_height;
  if (isNumber(fitWidth) || isNumber(fitHeight)) {
    worksheet.pageSetup.fitToPage = true;
    if (isNumber(fitWidth)) worksheet.pageSetup.fitToWidth = fitWidth;
    if (isNumber(fitHeight)) worksheet.pageSetup.fitToHeight = fitHeight;
  }
  if (isNumber(settings.scale)) worksheet.pageSetup.scale = settings.scale;
  if (settings.print_title_rows) {
    worksheet.pageSetup.printTitlesRow = String(settings.print_title_rows).replace(/\$/g, "");
  }
  if (settings.header_center) worksheet.headerFooter.oddHeader = `&C${settings.header_center}`;
  if (settings.footer_center) worksheet.headerFooter.oddFooter = `&C${settings.footer_center}`;
}

// ===========================================================================
// mutate
// ===========================================================================

async function mutate(request) {
  const inputPath = requirePath(request, "input_path", "mutate");
  const outputPath = requirePath(request, "output_path", "mutate");
  const mutations = (request.payload && request.payload.mutations) || [];
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(inputPath);
  for (const mutation of mutations) {
    sheetOf(workbook, mutation.sheet).getCell(mutation.cell).value = mutation.value ?? null;
  }
  await fs.mkdir(path.dirname(path.resolve(outputPath)), { recursive: true });
  await workbook.xlsx.writeFile(outputPath);
  return { written: outputPath };
}

// ===========================================================================
// helpers
// ===========================================================================

function requirePath(request, key, operation) {
  if (!request[key]) throw new Error(`${operation} requires ${key}.`);
  return request[key];
}

function sheetOf(workbook, name) {
  const worksheet = workbook.getWorksheet(name);
  if (!worksheet) throw new Error(`Sheet not found: ${name}`);
  return worksheet;
}

function noteText(note) {
  if (typeof note === "string") return note;
  return (note.texts || []).map((run) => run.text).join("");
}

function argbToHex(color) {
  if (!color || typeof color.argb !== "string") return null;
  const argb = color.argb.toUpperCase();
  if (argb.length === 8) return `#${argb.slice(2)}`;
  if (argb.length === 6) return `#${argb}`;
  return null;
}

function hexToArgb(hex) {
  const stripped = String(hex).replace(/^#/, "").toUpperCase();
  return stripped.length === 6 ? `FF${stripped}` : stripped;
}

function stripEquals(text) {
  return String(text).replace(/^=/, "");
}

function isNumber(value) {
  return typeof value === "number" && Number.isFinite(value);
}

function columnLetter(index) {
  let n = index;
  let letters = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
}

function decodeCell(address) {
  const match = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(String(address));
  if (!match) throw new Error(`Invalid cell address '${address}'.`);
  let col = 0;
  for (const ch of match[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { col, row: Number(match[2]) };
}

function decodeRange(ref) {
  const [start, end] = String(ref).split(":");
  const a = decodeCell(start);
  const b = decodeCell(end || start);
  return {
    left: Math.min(a.col, b.col),
    right: Math.max(a.col, b.col),
    top: Math.min(a.row, b.row),
    bottom: Math.max(a.row, b.row),
  };
}

// Pixel size of a PNG or JPEG buffer, needed for ExcelJS oneCell `ext`.
function imagePixelSize(buffer, extension) {
  if (extension === "png") {
    if (buffer.length < 24 || buffer.readUInt32BE(12) !== 0x49484452) {
      throw new Error("Malformed PNG: missing IHDR chunk.");
    }
    return { width: buffer.readUInt32BE(16), height: buffer.readUInt32BE(20) };
  }
  if (extension === "jpeg") {
    let offset = 2;
    while (offset + 9 < buffer.length) {
      if (buffer[offset] !== 0xff) {
        offset += 1;
        continue;
      }
      const marker = buffer[offset + 1];
      const length = buffer.readUInt16BE(offset + 2);
      const isFrame =
        marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc;
      if (isFrame) {
        return { width: buffer.readUInt16BE(offset + 7), height: buffer.readUInt16BE(offset + 5) };
      }
      offset += 2 + length;
    }
    throw new Error("Malformed JPEG: no frame header.");
  }
  throw new Error(`Cannot size '${extension}' image for a oneCell anchor.`);
}

module.exports = { describe, readModel, writeModel, mutate };
