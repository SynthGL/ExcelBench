#!/usr/bin/env node
"use strict";

/*
 * ExcelBench helper for SheetJS Community Edition (npm `xlsx`).
 *
 * Speaks the JSON-model helper protocol: one request on stdin, one JSON object
 * on stdout. Everything reported or written goes through the public SheetJS
 * API (`XLSX.readFile`, `XLSX.utils`, `XLSX.writeFile`). Features the CE build
 * cannot read or write are declared unsupported with the concrete reason.
 */

const fs = require("fs");
const XLSX = require("xlsx");

const LIBRARY = "sheetjs";

/* Read options: expose every cell-level detail SheetJS CE can parse. */
const READ_OPTIONS = {
  cellFormula: true, // `f` formula text
  cellNF: true, // `z` number format string
  cellStyles: true, // `s` fill, `!cols` widths, `!rows` heights
  cellDates: true, // date-formatted serials become `t: "d"` Date values
  UTC: true, // interpret Excel date serials as UTC instants
  sheetStubs: true, // keep formatted / commented / linked empty cells
};

/* Template-mutation lane: keep stored serials as numbers so untouched date
 * cells are not round-tripped through Date objects. */
const MUTATE_READ_OPTIONS = {
  cellFormula: true,
  cellNF: true,
  cellStyles: true,
};

const WRITE_OPTIONS = { bookType: "xlsx", compression: true };

const CE_STYLE_WRITE =
  "the SheetJS CE writer only emits number formats in cell styles";

const UNSUPPORTED_READ = {
  cell_border:
    "SheetJS CE does not expose cell borders: with cellStyles the XLSX reader attaches only the fill to cell.s",
  conditional_formats:
    "SheetJS CE has no conditional formatting API; the XLSX reader skips <conditionalFormatting>",
  data_validations:
    "SheetJS CE has no data validation API; the XLSX reader skips <dataValidations>",
  images:
    "SheetJS CE does not parse worksheet drawings, so embedded images are not exposed",
  pivot_tables: "SheetJS CE does not parse pivot table parts",
  freeze_panes:
    "SheetJS CE parses <sheetView> only for zoom and right-to-left; panes are not exposed",
  tables:
    "SheetJS CE does not parse table parts (xl/tables/*.xml); only the sheet autofilter is exposed",
  sheet_protection:
    "SheetJS CE's XLSX reader does not parse <sheetProtection>; ws['!protect'] is populated only for XLS/XLSB/XLML inputs",
  chart_anchors:
    "SheetJS CE does not parse charts or drawing anchors in worksheets",
};

const UNSUPPORTED_WRITE = {
  cell_border: `SheetJS CE has no border API; ${CE_STYLE_WRITE}`,
  conditional_format:
    "SheetJS CE has no conditional formatting API; the XLSX writer emits no <conditionalFormatting>",
  data_validation:
    "SheetJS CE has no data validation API; the XLSX writer emits no <dataValidations>",
  image: "SheetJS CE cannot embed images (drawing parts are not written)",
  pivot: "SheetJS CE cannot create pivot tables",
  freeze:
    "SheetJS CE's XLSX writer emits <sheetView> without a <pane> element, so panes cannot be set",
  table:
    "SheetJS CE cannot create table parts; only a sheet autofilter (ws['!autofilter']) is writable",
  chart: "SheetJS CE cannot create charts",
};

/* ------------------------------------------------------------------------ */
/* helpers                                                                  */
/* ------------------------------------------------------------------------ */

function fail(message) {
  process.stdout.write(
    JSON.stringify({ error: "sheetjs_failed", message: String(message) }) + "\n"
  );
  process.exit(1);
}

function readStdin() {
  return fs.readFileSync(0, "utf8");
}

function pad2(n) {
  return String(n).padStart(2, "0");
}

function dateParts(d) {
  /* SheetJS hands back Date objects in UTC (READ_OPTIONS.UTC). Round to the
   * nearest second to drop floating-point serial noise. */
  const t = new Date(Math.round(d.getTime() / 1000) * 1000);
  const day = `${String(t.getUTCFullYear()).padStart(4, "0")}-${pad2(
    t.getUTCMonth() + 1
  )}-${pad2(t.getUTCDate())}`;
  const hasTime =
    t.getUTCHours() !== 0 || t.getUTCMinutes() !== 0 || t.getUTCSeconds() !== 0;
  const time = `${pad2(t.getUTCHours())}:${pad2(t.getUTCMinutes())}:${pad2(
    t.getUTCSeconds()
  )}`;
  return { day, hasTime, time };
}

function hexColor(color) {
  if (!color || typeof color.rgb !== "string") return null;
  const rgb = color.rgb.replace(/^#/, "");
  if (rgb.length < 6) return null;
  return `#${rgb.slice(-6).toUpperCase()}`;
}

function sheetCellEntries(ws) {
  return Object.keys(ws).filter((key) => key[0] !== "!");
}

/* ------------------------------------------------------------------------ */
/* read_model                                                               */
/* ------------------------------------------------------------------------ */

function modelValue(cell) {
  switch (cell.t) {
    case "z":
      return null;
    case "e":
      return {
        type: "error",
        value: cell.w || XLSX.utils.format_cell(cell),
        formula: cell.f ? `=${cell.f}` : null,
      };
    default:
      break;
  }
  if (typeof cell.f === "string" && cell.f.length > 0) {
    const formula = `=${cell.f}`;
    return { type: "formula", value: formula, formula };
  }
  switch (cell.t) {
    case "s":
      return { type: "string", value: cell.v == null ? "" : String(cell.v), formula: null };
    case "n":
      return { type: "number", value: cell.v, formula: null };
    case "b":
      return { type: "boolean", value: Boolean(cell.v), formula: null };
    case "d": {
      const parts = dateParts(cell.v instanceof Date ? cell.v : new Date(cell.v));
      if (parts.hasTime) {
        return { type: "datetime", value: `${parts.day}T${parts.time}`, formula: null };
      }
      return { type: "date", value: parts.day, formula: null };
    }
    default:
      return null;
  }
}

function modelFormat(cell) {
  const fmt = {};
  if (typeof cell.z === "string") fmt.number_format = cell.z;
  else if (typeof cell.z === "number") fmt.number_format = XLSX.SSF.get_table()[cell.z];
  const fill = cell.s;
  if (fill && fill.patternType === "solid") {
    const color = hexColor(fill.fgColor);
    if (color) fmt.bg_color = color;
  }
  return fmt;
}

function readPrintTitleRows(names, sheetIndex) {
  for (const name of names) {
    if (name.Name !== "_xlnm.Print_Titles" || name.Sheet !== sheetIndex) continue;
    for (const part of String(name.Ref || "").split(",")) {
      const addr = part.slice(part.lastIndexOf("!") + 1);
      if (/^\$?\d+:\$?\d+$/.test(addr)) return addr;
    }
  }
  return null;
}

function readSheet(wb, sheetName, sheetIndex) {
  const ws = wb.Sheets[sheetName];
  const names = (wb.Workbook && wb.Workbook.Names) || [];
  const cells = {};
  const hyperlinks = [];
  const comments = [];

  for (const addr of sheetCellEntries(ws)) {
    const cell = ws[addr];
    if (!cell || typeof cell !== "object") continue;
    const entry = {};
    const value = modelValue(cell);
    if (value) entry.value = value;
    const fmt = modelFormat(cell);
    if (Object.keys(fmt).length > 0) entry.format = fmt;
    if (Object.keys(entry).length > 0) cells[addr] = entry;

    if (cell.l && typeof cell.l.Target === "string") {
      const internal = Boolean(cell.l.Rel && cell.l.Rel.TargetMode === "Internal");
      hyperlinks.push({
        cell: addr,
        target: internal ? cell.l.Target.replace(/^#/, "") : cell.l.Target,
        display: cell.t === "z" ? null : cell.w != null ? cell.w : String(cell.v),
        tooltip: cell.l.Tooltip != null ? String(cell.l.Tooltip) : null,
        internal,
      });
    }
    if (Array.isArray(cell.c) && cell.c.length > 0) {
      const first = cell.c[0];
      comments.push({
        cell: addr,
        text: first.t == null ? "" : String(first.t),
        author: first.a != null ? String(first.a) : null,
        threaded: Boolean(first.T),
      });
    }
  }

  const rowHeights = {};
  (ws["!rows"] || []).forEach((row, idx) => {
    if (row && typeof row.hpt === "number") rowHeights[String(idx + 1)] = row.hpt;
  });
  const columnWidths = {};
  (ws["!cols"] || []).forEach((col, idx) => {
    if (col && typeof col.width === "number" && !Number.isNaN(col.width)) {
      columnWidths[XLSX.utils.encode_col(idx)] = col.width;
    }
  });
  const merged = (ws["!merges"] || []).map((rng) => XLSX.utils.encode_range(rng));

  const pageSetup = {};
  const titles = readPrintTitleRows(names, sheetIndex);
  if (titles) pageSetup.print_title_rows = titles;

  return {
    name: sheetName,
    cells,
    row_heights: rowHeights,
    column_widths: columnWidths,
    merged_ranges: merged,
    hyperlinks,
    comments,
    page_setup: pageSetup,
  };
}

function readModel(inputPath) {
  const wb = XLSX.readFile(inputPath, READ_OPTIONS);
  const names = (wb.Workbook && wb.Workbook.Names) || [];
  const namedRanges = [];
  for (const name of names) {
    /* `_xlnm.*` are Excel built-in names (print area/titles, filter
     * database), not user-defined named ranges. */
    if (String(name.Name).startsWith("_xlnm.")) continue;
    const local = name.Sheet != null;
    namedRanges.push({
      name: String(name.Name),
      scope: local ? "sheet" : "workbook",
      sheet: local ? wb.SheetNames[name.Sheet] || null : null,
      refers_to: String(name.Ref || ""),
    });
  }
  return {
    unsupported: UNSUPPORTED_READ,
    named_ranges: namedRanges,
    sheets: wb.SheetNames.map((sheetName, idx) => readSheet(wb, sheetName, idx)),
  };
}

/* ------------------------------------------------------------------------ */
/* write_model                                                              */
/* ------------------------------------------------------------------------ */

const ERROR_FORMULAS = {
  "#DIV/0!": "1/0",
  "#N/A": "NA()",
  "#VALUE!": '"text"+1',
  "#REF!": "#REF!",
  "#NAME?": "_undefined_name_",
  "#NUM!": "SQRT(-1)",
  "#NULL!": "A1:A2 B1:B2",
};

function requireSheet(wb, name) {
  const ws = wb.Sheets[name];
  if (!ws) throw new Error(`worksheet not found: ${name}`);
  return ws;
}

function extendRef(ws, addr) {
  const pos = XLSX.utils.decode_cell(addr);
  if (!ws["!ref"]) {
    ws["!ref"] = XLSX.utils.encode_range(pos, pos);
    return;
  }
  const range = XLSX.utils.decode_range(ws["!ref"]);
  range.s.r = Math.min(range.s.r, pos.r);
  range.s.c = Math.min(range.s.c, pos.c);
  range.e.r = Math.max(range.e.r, pos.r);
  range.e.c = Math.max(range.e.c, pos.c);
  ws["!ref"] = XLSX.utils.encode_range(range);
}

function cellFor(ws, addr) {
  const ref = String(addr).toUpperCase();
  const cell = XLSX.utils.sheet_get_cell(ws, ref);
  extendRef(ws, ref);
  return cell;
}

function setValue(cell, t, v) {
  delete cell.f;
  delete cell.F;
  delete cell.w;
  cell.t = t;
  cell.v = v;
}

function setFormula(cell, formula) {
  delete cell.v;
  delete cell.w;
  delete cell.t;
  cell.f = String(formula).replace(/^=/, "");
}

function applyCellValue(ws, op) {
  const cell = cellFor(ws, op.cell);
  const value = op.value;
  switch (op.type) {
    case "string":
      setValue(cell, "s", value == null ? "" : String(value));
      break;
    case "number":
      setValue(cell, "n", Number(value));
      break;
    case "boolean":
      setValue(cell, "b", Boolean(value));
      break;
    case "date":
    case "datetime":
      /* SheetJS date cells: `t: "d"` with an ISO string; the writer converts
       * it to a serial and applies the default date format when none set. */
      setValue(cell, "d", String(value));
      break;
    case "formula":
      setFormula(cell, op.formula != null ? op.formula : value);
      break;
    case "error": {
      const formula = ERROR_FORMULAS[String(value)];
      if (!formula) throw new Error(`no formula known to produce error ${value}`);
      setFormula(cell, formula);
      break;
    }
    case "blank":
      setValue(cell, "z", undefined);
      delete cell.v;
      break;
    default:
      throw new Error(`unsupported cell value type: ${op.type}`);
  }
}

function applyCellFormat(ws, op) {
  const fmt = op.format || {};
  const cell = cellFor(ws, op.cell);
  if (fmt.number_format != null) {
    XLSX.utils.cell_set_number_format(cell, String(fmt.number_format));
  }
  const dropped = Object.keys(fmt).filter((key) => key !== "number_format");
  if (dropped.length > 0) {
    process.stderr.write(
      `sheetjs: ${op.sheet}!${op.cell} format keys not writable (${CE_STYLE_WRITE}): ${dropped.join(", ")}\n`
    );
  }
}

function applyHyperlink(ws, op) {
  const link = op.link || {};
  const cell = cellFor(ws, link.cell);
  if (link.display != null) setValue(cell, "s", String(link.display));
  const tooltip = link.tooltip != null ? String(link.tooltip) : undefined;
  if (link.internal) {
    XLSX.utils.cell_set_internal_link(cell, String(link.target).replace(/^#/, ""), tooltip);
  } else {
    XLSX.utils.cell_set_hyperlink(cell, String(link.target), tooltip);
  }
}

function applyComment(ws, op) {
  const comment = op.comment || {};
  if (!comment.cell || comment.text == null) return;
  const cell = cellFor(ws, comment.cell);
  XLSX.utils.cell_add_comment(cell, String(comment.text), comment.author || undefined);
}

function workbookNames(wb) {
  if (!wb.Workbook) wb.Workbook = {};
  if (!wb.Workbook.Names) wb.Workbook.Names = [];
  return wb.Workbook.Names;
}

function quoteSheet(name) {
  return `'${String(name).replace(/'/g, "''")}'`;
}

function applyNamedRange(wb, op) {
  const nr = op.named_range || {};
  if (!nr.name || !nr.refers_to) return;
  const entry = { Name: String(nr.name), Ref: String(nr.refers_to).replace(/^=/, "") };
  if (nr.scope === "sheet") {
    const idx = wb.SheetNames.indexOf(op.sheet);
    if (idx < 0) throw new Error(`worksheet not found: ${op.sheet}`);
    entry.Sheet = idx;
  }
  workbookNames(wb).push(entry);
}

function applyProtection(ws, op) {
  const cfg = op.settings || {};
  if (!cfg.protected) {
    delete ws["!protect"];
    return;
  }
  /* SheetJS `!protect` keys mirror the raw <sheetProtection> attributes:
   * true means the action is blocked, matching the op semantics. */
  const protect = {};
  const keys = {
    format_cells: "formatCells",
    insert_rows: "insertRows",
    select_locked_cells: "selectLockedCells",
    select_unlocked_cells: "selectUnlockedCells",
    sort: "sort",
    auto_filter: "autoFilter",
  };
  for (const [src, dst] of Object.entries(keys)) {
    if (cfg[src] != null) protect[dst] = Boolean(cfg[src]);
  }
  if (cfg.password) protect.password = String(cfg.password);
  ws["!protect"] = protect;
}

function applyPageSetup(wb, op) {
  const cfg = op.settings || {};
  if (cfg.print_title_rows != null) {
    const idx = wb.SheetNames.indexOf(op.sheet);
    if (idx < 0) throw new Error(`worksheet not found: ${op.sheet}`);
    workbookNames(wb).push({
      Name: "_xlnm.Print_Titles",
      Sheet: idx,
      Ref: `${quoteSheet(op.sheet)}!${cfg.print_title_rows}`,
    });
  }
  const dropped = Object.keys(cfg).filter(
    (key) => key !== "print_title_rows" && cfg[key] != null
  );
  if (dropped.length > 0) {
    process.stderr.write(
      `sheetjs: ${op.sheet} page setup keys not writable (SheetJS CE writes no <pageSetup>/<headerFooter>): ${dropped.join(", ")}\n`
    );
  }
}

function writeModel(outputPath, ops) {
  const wb = XLSX.utils.book_new();
  for (const op of ops) {
    const reason = UNSUPPORTED_WRITE[op.op];
    if (reason) throw new Error(`${op.op}: ${reason}`);
    switch (op.op) {
      case "add_sheet":
        XLSX.utils.book_append_sheet(wb, XLSX.utils.sheet_new(), String(op.name));
        break;
      case "cell_value":
        applyCellValue(requireSheet(wb, op.sheet), op);
        break;
      case "cell_format":
        applyCellFormat(requireSheet(wb, op.sheet), op);
        break;
      case "row_height": {
        const ws = requireSheet(wb, op.sheet);
        if (!ws["!rows"]) ws["!rows"] = [];
        const idx = Number(op.row) - 1;
        ws["!rows"][idx] = Object.assign(ws["!rows"][idx] || {}, { hpt: Number(op.height) });
        break;
      }
      case "column_width": {
        const ws = requireSheet(wb, op.sheet);
        if (!ws["!cols"]) ws["!cols"] = [];
        ws["!cols"][XLSX.utils.decode_col(String(op.column))] = { width: Number(op.width) };
        break;
      }
      case "merge": {
        const ws = requireSheet(wb, op.sheet);
        if (!ws["!merges"]) ws["!merges"] = [];
        ws["!merges"].push(XLSX.utils.decode_range(String(op.range)));
        break;
      }
      case "hyperlink":
        applyHyperlink(requireSheet(wb, op.sheet), op);
        break;
      case "comment":
        applyComment(requireSheet(wb, op.sheet), op);
        break;
      case "named_range":
        applyNamedRange(wb, op);
        break;
      case "protection":
        applyProtection(requireSheet(wb, op.sheet), op);
        break;
      case "page_setup":
        applyPageSetup(wb, op);
        break;
      default:
        throw new Error(`unknown write op: ${op.op}`);
    }
  }
  if (wb.SheetNames.length === 0) throw new Error("write_model received no add_sheet op");
  XLSX.writeFile(wb, outputPath, WRITE_OPTIONS);
}

/* ------------------------------------------------------------------------ */
/* mutate                                                                   */
/* ------------------------------------------------------------------------ */

function mutate(inputPath, outputPath, mutations) {
  const wb = XLSX.readFile(inputPath, MUTATE_READ_OPTIONS);
  for (const m of mutations) {
    const ws = requireSheet(wb, m.sheet);
    /* sheet_add_aoa is SheetJS's documented way to set values in an existing
     * sheet; it keeps the target cell's number format. */
    XLSX.utils.sheet_add_aoa(ws, [[m.value]], { origin: String(m.cell).toUpperCase() });
  }
  XLSX.writeFile(wb, outputPath, WRITE_OPTIONS);
}

/* ------------------------------------------------------------------------ */
/* dispatch                                                                 */
/* ------------------------------------------------------------------------ */

function main() {
  let request;
  try {
    request = JSON.parse(readStdin());
  } catch (err) {
    fail(`invalid JSON request on stdin: ${err.message}`);
  }
  const payload = request.payload || {};
  try {
    switch (request.operation) {
      case "describe":
        return {
          library: LIBRARY,
          version: XLSX.version,
          capabilities: ["read", "write", "modify"],
          unsupported_write: UNSUPPORTED_WRITE,
        };
      case "read_model":
        if (!request.input_path) throw new Error("read_model requires input_path");
        return { model: readModel(request.input_path) };
      case "write_model":
        if (!request.output_path) throw new Error("write_model requires output_path");
        writeModel(request.output_path, payload.ops || []);
        return { written: request.output_path };
      case "mutate":
        if (!request.input_path || !request.output_path) {
          throw new Error("mutate requires input_path and output_path");
        }
        mutate(request.input_path, request.output_path, payload.mutations || []);
        return { written: request.output_path };
      default:
        throw new Error(`unknown operation: ${request.operation}`);
    }
  } catch (err) {
    fail(err instanceof Error ? err.message : err);
  }
  return null;
}

const result = main();
process.stdout.write(JSON.stringify(result) + "\n");
