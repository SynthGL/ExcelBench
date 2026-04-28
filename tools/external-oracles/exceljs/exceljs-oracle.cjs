const fs = require("node:fs/promises");
const path = require("node:path");
let ExcelJS;
let JSZip;

const PIXEL_PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";

main().catch((error) => {
  process.stdout.write(
    `${JSON.stringify({ error: "exceljs_oracle_failed", message: error.message })}\n`,
  );
  process.exitCode = 1;
});

async function main() {
  loadDependencies();
  const request = JSON.parse(await readStdin());
  const payload =
    request.operation === "write_fixture"
      ? await writeFixture(request)
      : request.operation === "read_metadata"
        ? await readMetadata(request)
        : fail(`Unsupported operation '${request.operation}'.`);
  process.stdout.write(`${JSON.stringify(payload)}\n`);
}

function loadDependencies() {
  try {
    ExcelJS = require("exceljs");
    JSZip = require("jszip");
  } catch (error) {
    fail(`Missing ExcelJS oracle dependency: ${error.message}`);
  }
}

async function writeFixture(request) {
  if (!request.output_path) {
    fail("write_fixture requires output_path.");
  }
  const payload = request.payload || {};
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "ExcelBench ExcelJS Oracle";
  workbook.created = new Date("2026-04-28T00:00:00.000Z");
  workbook.modified = new Date("2026-04-28T00:00:00.000Z");

  configureSheets(workbook, payload.sheets || []);
  applyColumns(workbook, payload.columns || []);
  applyCells(workbook, payload.cells || []);
  applyRichText(workbook, payload.rich_text || []);
  applyHyperlinks(workbook, payload.hyperlinks || []);
  applyComments(workbook, payload.comments || []);
  applyMergedRanges(workbook, payload.merged_ranges || []);
  applyDataValidations(workbook, payload.data_validations || []);
  applyTables(workbook, payload.tables || []);
  applyImages(workbook, payload.images || []);
  await applyProtection(workbook, payload.protection || []);

  await fs.mkdir(path.dirname(path.resolve(request.output_path)), { recursive: true });
  await workbook.xlsx.writeFile(request.output_path);
  return {
    fixture_id: request.fixture_id,
    operation: request.operation,
    output_path: request.output_path,
    tool: "exceljs",
    counts: {
      sheets: workbook.worksheets.length,
      cells: (payload.cells || []).length,
      formulas: (payload.cells || []).filter((cell) => cell.type === "formula").length,
      rich_text: (payload.rich_text || []).length,
      comments: (payload.comments || []).length,
      hyperlinks: (payload.hyperlinks || []).length,
      tables: (payload.tables || []).length,
      data_validations: (payload.data_validations || []).length,
      merged_ranges: (payload.merged_ranges || []).length,
      images: (payload.images || []).length,
      protected_sheets: (payload.protection || []).length,
    },
  };
}

async function readMetadata(request) {
  const inputPath = request.input_path || request.output_path;
  if (!inputPath) {
    fail("read_metadata requires input_path or output_path.");
  }
  const buffer = await fs.readFile(inputPath);
  const zip = await JSZip.loadAsync(buffer);
  const partNames = Object.keys(zip.files);
  return {
    fixture_id: request.fixture_id,
    operation: request.operation,
    input_path: inputPath,
    tool: "exceljs",
    counts: {
      worksheets: countParts(partNames, "xl/worksheets/sheet"),
      tables: countParts(partNames, "xl/tables/table"),
      drawings: countParts(partNames, "xl/drawings/drawing"),
      media: countParts(partNames, "xl/media/"),
      comments: countParts(partNames, "xl/comments"),
      vml_drawings: countParts(partNames, "xl/drawings/vmlDrawing"),
      shared_strings: countParts(partNames, "xl/sharedStrings"),
      calc_chain: countParts(partNames, "xl/calcChain"),
    },
  };
}

function configureSheets(workbook, sheets) {
  const sheetSpecs = sheets.length ? sheets : [{ name: "Sheet1" }];
  for (const sheet of sheetSpecs) {
    const worksheet = workbook.addWorksheet(sheet.name);
    if (sheet.freeze_panes) {
      worksheet.views = [
        {
          state: "frozen",
          xSplit: Number(sheet.freeze_panes.x_split || 0),
          ySplit: Number(sheet.freeze_panes.y_split || 0),
        },
      ];
    }
  }
}

function applyColumns(workbook, columns) {
  for (const item of columns) {
    const worksheet = getWorksheet(workbook, item.sheet);
    const column = worksheet.getColumn(item.key || item.column);
    if (item.width !== undefined) {
      column.width = Number(item.width);
    }
  }
}

function applyCells(workbook, cells) {
  for (const item of cells) {
    const cell = getCell(workbook, item.sheet, item.cell);
    if (item.type === "formula") {
      cell.value = { formula: item.formula || "", result: item.result ?? null };
    } else {
      cell.value = item.value ?? null;
    }
    applyCellStyle(cell, item);
  }
}

function applyRichText(workbook, items) {
  for (const item of items) {
    const cell = getCell(workbook, item.sheet, item.cell);
    cell.value = {
      richText: item.runs.map((run) => ({
        text: run.text,
        font: {
          bold: Boolean(run.bold),
          italic: Boolean(run.italic),
          color: run.font_color ? { argb: normalizeColor(run.font_color) } : undefined,
        },
      })),
    };
  }
}

function applyHyperlinks(workbook, items) {
  for (const item of items) {
    const cell = getCell(workbook, item.sheet, item.cell);
    cell.value = {
      text: item.text || item.url,
      hyperlink: item.url,
      tooltip: item.tooltip,
    };
  }
}

function applyComments(workbook, comments) {
  for (const item of comments) {
    const cell = getCell(workbook, item.sheet, item.cell);
    cell.note = {
      texts: [
        {
          text: item.text,
          font: { size: 11, color: { argb: "FF000000" }, name: "Calibri" },
        },
      ],
      margins: { insetmode: "auto" },
      protection: { locked: true, lockText: false },
      editAs: "absolute",
    };
  }
}

function applyMergedRanges(workbook, mergedRanges) {
  for (const item of mergedRanges) {
    getWorksheet(workbook, item.sheet).mergeCells(item.range);
  }
}

function applyDataValidations(workbook, dataValidations) {
  for (const item of dataValidations) {
    const cell = getCell(workbook, item.sheet, item.cell);
    cell.dataValidation = {
      type: item.type || "list",
      allowBlank: item.allow_blank ?? true,
      formulae: item.formulae || item.values || [],
      showErrorMessage: true,
      errorTitle: item.error_title || "Invalid value",
      error: item.error || "Choose a listed value.",
    };
  }
}

function applyTables(workbook, tables) {
  for (const item of tables) {
    const worksheet = getWorksheet(workbook, item.sheet);
    worksheet.addTable({
      name: item.name,
      ref: item.ref,
      headerRow: true,
      totalsRow: Boolean(item.totals_row),
      style: { theme: item.theme || "TableStyleMedium2", showRowStripes: true },
      columns: item.columns.map((column) => ({
        name: column.name,
        totalsRowFunction: column.totals_row_function,
      })),
      rows: item.rows,
    });
  }
}

function applyImages(workbook, images) {
  for (const item of images) {
    const worksheet = getWorksheet(workbook, item.sheet);
    const imageId = workbook.addImage({
      base64: item.base64 || PIXEL_PNG_BASE64,
      extension: item.extension || "png",
    });
    worksheet.addImage(imageId, item.range || `${item.cell || "A1"}:${item.cell || "A1"}`);
  }
}

async function applyProtection(workbook, protections) {
  for (const item of protections) {
    const worksheet = getWorksheet(workbook, item.sheet);
    await worksheet.protect(item.password || "", {
      selectLockedCells: true,
      selectUnlockedCells: true,
      formatCells: false,
      insertRows: false,
      deleteRows: false,
    });
  }
}

function applyCellStyle(cell, item) {
  if (item.num_fmt) {
    cell.numFmt = item.num_fmt;
  }
  if (item.font) {
    cell.font = {
      bold: Boolean(item.font.bold),
      italic: Boolean(item.font.italic),
      color: item.font.color ? { argb: normalizeColor(item.font.color) } : undefined,
    };
  }
  if (item.fill_color) {
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: normalizeColor(item.fill_color) },
    };
  }
}

function getWorksheet(workbook, name) {
  const worksheet = workbook.getWorksheet(name);
  if (!worksheet) {
    fail(`Sheet not found: ${name}`);
  }
  return worksheet;
}

function getCell(workbook, sheetName, address) {
  return getWorksheet(workbook, sheetName).getCell(address);
}

function normalizeColor(color) {
  const stripped = String(color).replace(/^#/, "").toUpperCase();
  return stripped.length === 6 ? `FF${stripped}` : stripped;
}

function countParts(partNames, fragment) {
  return partNames.filter((partName) => partName.includes(fragment)).length;
}

function fail(message) {
  throw new Error(message);
}

async function readStdin() {
  const chunks = [];
  for await (const chunk of process.stdin) {
    chunks.push(Buffer.from(chunk));
  }
  return Buffer.concat(chunks).toString("utf8");
}
