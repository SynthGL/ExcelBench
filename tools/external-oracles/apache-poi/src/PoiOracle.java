import java.io.File;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.charset.StandardCharsets;
import java.sql.Timestamp;
import java.sql.Date;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Base64;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.PatternFormatting;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.model.ThemesTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide;
import org.apache.poi.xssf.usermodel.XSSFPatternFormatting;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheetConditionalFormatting;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumn;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumns;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableStyleInfo;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTConditionalFormatting;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCfRule;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPane;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STPaneState;

public final class PoiOracle {
    private static final String PIXEL_PNG_BASE64 =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";

    private PoiOracle() {}

    public static void main(String[] args) throws Exception {
        if (args.length == 1 && "--self-test".equals(args[0])) {
            System.out.println("OK");
            return;
        }
        if (args.length == 4 && "write_adapter_workbook".equals(args[0])) {
            writeAdapterWorkbook(args[1], args[2], args[3]);
            return;
        }
        if (args.length != 3 || !"write_fixture".equals(args[0])) {
            throw new IllegalArgumentException("Usage: PoiOracle write_fixture <fixture_id> <output_path>");
        }
        writeFixture(args[1], args[2]);
    }

    private static void writeAdapterWorkbook(String fixtureId, String outputPath, String specPath) throws Exception {
        File output = new File(outputPath);
        File parent = output.getParentFile();
        if (parent != null) {
            parent.mkdirs();
        }

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            workbook.getProperties().getCoreProperties().setCreator("ExcelBench Apache POI Adapter");
            List<String> lines = Files.readAllLines(Path.of(specPath), StandardCharsets.UTF_8);
            Map<String, XSSFSheet> sheets = new HashMap<>();

            if (lines.isEmpty()) {
                XSSFSheet defaultSheet = workbook.createSheet("Sheet1");
                sheets.put("Sheet1", defaultSheet);
            }

            for (String line : lines) {
                if (line.isBlank()) {
                    continue;
                }
                String[] parts = line.split("\t", -1);
                String op = parts[0];
                if ("SHEET".equals(op)) {
                    String sheetName = decode(parts[1]);
                    if (!sheets.containsKey(sheetName)) {
                        sheets.put(sheetName, workbook.createSheet(sheetName));
                    }
                }
            }
            if (sheets.isEmpty()) {
                XSSFSheet defaultSheet = workbook.createSheet("Sheet1");
                sheets.put("Sheet1", defaultSheet);
            }

            for (String line : lines) {
                if (line.isBlank()) {
                    continue;
                }
                String[] parts = line.split("\t", -1);
                String op = parts[0];
                if ("SHEET".equals(op)) {
                    continue;
                }
                applySpecLine(workbook, sheets, op, parts);
            }

            try (FileOutputStream stream = new FileOutputStream(output)) {
                workbook.write(stream);
            }
        }

        System.out.println("{"
            + "\"fixture_id\":\"" + json(fixtureId) + "\","
            + "\"operation\":\"write_adapter_workbook\","
            + "\"output_path\":\"" + json(outputPath) + "\","
            + "\"tool\":\"apache-poi\""
            + "}");
    }

    private static void writeFixture(String fixtureId, String outputPath) throws Exception {
        File output = new File(outputPath);
        File parent = output.getParentFile();
        if (parent != null) {
            parent.mkdirs();
        }

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            workbook.getProperties().getCoreProperties().setCreator("ExcelBench Apache POI Oracle");
            workbook.lockStructure();
            XSSFSheet sheet = workbook.createSheet("POI");
            sheet.createFreezePane(1, 1);
            sheet.setColumnWidth(0, 18 * 256);
            sheet.setColumnWidth(1, 14 * 256);
            sheet.setColumnWidth(3, 24 * 256);

            XSSFCellStyle currency = workbook.createCellStyle();
            currency.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("$#,##0"));

            setText(sheet, "A1", "Metric", true);
            setText(sheet, "B1", "Value", true);
            setText(sheet, "A2", "Revenue", false);
            setNumber(sheet, "B2", 1200, currency);
            setText(sheet, "A3", "COGS", false);
            setNumber(sheet, "B3", -450, currency);
            setText(sheet, "A4", "Gross profit", false);
            Cell formula = getCell(sheet, "B4");
            formula.setCellFormula("SUM(B2:B3)");
            formula.setCellStyle(currency);

            setText(sheet, "D1", "Merged review header", false);
            sheet.addMergedRegion(CellRangeAddress.valueOf("D1:F1"));
            setRichText(workbook, sheet, "D2");
            setHyperlink(workbook, sheet, "D4");
            addComment(workbook, sheet, "B4");
            addDataValidation(sheet, "C2");
            addTable(sheet);
            addImage(workbook, sheet);
            sheet.protectSheet("audit");

            try (FileOutputStream stream = new FileOutputStream(output)) {
                workbook.write(stream);
            }
        }

        System.out.println("{"
            + "\"fixture_id\":\"" + json(fixtureId) + "\","
            + "\"operation\":\"write_fixture\","
            + "\"output_path\":\"" + json(outputPath) + "\","
            + "\"tool\":\"apache-poi\","
            + "\"counts\":{"
            + "\"sheets\":1,"
            + "\"cells\":9,"
            + "\"formulas\":1,"
            + "\"rich_text\":1,"
            + "\"comments\":1,"
            + "\"hyperlinks\":1,"
            + "\"tables\":1,"
            + "\"data_validations\":1,"
            + "\"merged_ranges\":1,"
            + "\"images\":1,"
            + "\"protected_sheets\":1,"
            + "\"protected_workbook\":1"
            + "}}");
    }

    private static void setText(XSSFSheet sheet, String address, String value, boolean bold) {
        Cell cell = getCell(sheet, address);
        cell.setCellValue(value);
        if (bold) {
            Workbook workbook = sheet.getWorkbook();
            XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();
            Font font = workbook.createFont();
            font.setBold(true);
            style.setFont(font);
            cell.setCellStyle(style);
        }
    }

    private static void setNumber(XSSFSheet sheet, String address, double value, XSSFCellStyle style) {
        Cell cell = getCell(sheet, address);
        cell.setCellValue(value);
        cell.setCellStyle(style);
    }

    private static void setRichText(XSSFWorkbook workbook, XSSFSheet sheet, String address) {
        XSSFFont boldFont = workbook.createFont();
        boldFont.setBold(true);
        XSSFFont italicFont = workbook.createFont();
        italicFont.setItalic(true);
        XSSFRichTextString text = new XSSFRichTextString("Apache POI rich text");
        text.applyFont(0, 10, boldFont);
        text.applyFont(10, text.length(), italicFont);
        getCell(sheet, address).setCellValue(text);
    }

    private static void setHyperlink(XSSFWorkbook workbook, XSSFSheet sheet, String address) {
        CreationHelper helper = workbook.getCreationHelper();
        Hyperlink link = helper.createHyperlink(HyperlinkType.URL);
        link.setAddress("https://poi.apache.org/");
        link.setLabel("Apache POI");
        Cell cell = getCell(sheet, address);
        cell.setCellValue("Apache POI");
        cell.setHyperlink(link);
    }

    private static void addComment(XSSFWorkbook workbook, XSSFSheet sheet, String address) {
        CreationHelper helper = workbook.getCreationHelper();
        Drawing<?> drawing = sheet.createDrawingPatriarch();
        ClientAnchor anchor = helper.createClientAnchor();
        Cell cell = getCell(sheet, address);
        anchor.setCol1(cell.getColumnIndex() + 1);
        anchor.setCol2(cell.getColumnIndex() + 4);
        anchor.setRow1(cell.getRowIndex());
        anchor.setRow2(cell.getRowIndex() + 3);
        Comment comment = drawing.createCellComment(anchor);
        RichTextString text = helper.createRichTextString("POI formula comment.");
        comment.setString(text);
        comment.setAuthor("Apache POI Oracle");
        cell.setCellComment(comment);
    }

    private static void addDataValidation(XSSFSheet sheet, String address) {
        CellReference ref = new CellReference(address);
        DataValidationHelper helper = sheet.getDataValidationHelper();
        DataValidationConstraint constraint =
            helper.createExplicitListConstraint(new String[] {"Open", "Closed", "Review"});
        CellRangeAddressList range = new CellRangeAddressList(ref.getRow(), ref.getRow(), ref.getCol(), ref.getCol());
        DataValidation validation = helper.createValidation(constraint, range);
        validation.setShowErrorBox(true);
        validation.createErrorBox("Invalid value", "Choose a listed value.");
        sheet.addValidationData(validation);
    }

    private static void addTable(XSSFSheet sheet) {
        setText(sheet, "F1", "Item", false);
        setText(sheet, "G1", "Status", false);
        setText(sheet, "F2", "Revenue", false);
        setText(sheet, "G2", "Open", false);
        setText(sheet, "F3", "COGS", false);
        setText(sheet, "G3", "Closed", false);
        setText(sheet, "F4", "Gross profit", false);
        setText(sheet, "G4", "Review", false);

        AreaReference area = new AreaReference(
            new CellReference("F1"),
            new CellReference("G4"),
            SpreadsheetVersion.EXCEL2007
        );
        XSSFTable table = sheet.createTable(area);
        CTTable ctTable = table.getCTTable();
        ctTable.setId(1);
        ctTable.setName("PoiReviewTable");
        ctTable.setDisplayName("PoiReviewTable");
        ctTable.setRef("F1:G4");
        ctTable.setTotalsRowShown(false);
        if (ctTable.getTableColumns() == null) {
            CTTableColumns columns = ctTable.addNewTableColumns();
            columns.setCount(2);
            CTTableColumn first = columns.addNewTableColumn();
            first.setId(1);
            first.setName("Item");
            CTTableColumn second = columns.addNewTableColumn();
            second.setId(2);
            second.setName("Status");
        }
        CTTableStyleInfo style = ctTable.isSetTableStyleInfo()
            ? ctTable.getTableStyleInfo()
            : ctTable.addNewTableStyleInfo();
        style.setName("TableStyleMedium2");
        style.setShowRowStripes(true);
    }

    private static void addImage(XSSFWorkbook workbook, XSSFSheet sheet) {
        byte[] image = Base64.getDecoder().decode(PIXEL_PNG_BASE64.getBytes(StandardCharsets.US_ASCII));
        int picture = workbook.addPicture(image, Workbook.PICTURE_TYPE_PNG);
        CreationHelper helper = workbook.getCreationHelper();
        Drawing<?> drawing = sheet.createDrawingPatriarch();
        ClientAnchor anchor = helper.createClientAnchor();
        anchor.setCol1(3);
        anchor.setCol2(5);
        anchor.setRow1(5);
        anchor.setRow2(8);
        drawing.createPicture(anchor, picture);
    }

    private static void applySpecLine(
        XSSFWorkbook workbook,
        Map<String, XSSFSheet> sheets,
        String op,
        String[] parts
    ) {
        switch (op) {
            case "VALUE":
                applyValue(workbook, sheets.get(decode(parts[1])), decode(parts[2]), decode(parts[3]), decode(parts[4]), decode(parts[5]));
                return;
            case "FORMAT":
                applyFormat(workbook, sheets.get(decode(parts[1])), decode(parts[2]), parts);
                return;
            case "BORDER":
                applyBorder(workbook, sheets.get(decode(parts[1])), decode(parts[2]), decode(parts[3]));
                return;
            case "CF":
                applyConditionalFormat(sheets.get(decode(parts[1])), decode(parts[2]), decode(parts[3]), decode(parts[4]), decode(parts[5]), "1".equals(parts[6]), decode(parts[7]));
                return;
            case "ROW_HEIGHT":
                applyRowHeight(sheets.get(decode(parts[1])), Integer.parseInt(parts[2]), Float.parseFloat(parts[3]));
                return;
            case "COL_WIDTH":
                applyColumnWidth(sheets.get(decode(parts[1])), decode(parts[2]), Float.parseFloat(parts[3]));
                return;
            case "MERGE":
                sheets.get(decode(parts[1])).addMergedRegion(CellRangeAddress.valueOf(decode(parts[2])));
                return;
            case "VALIDATION":
                applyValidation(sheets.get(decode(parts[1])), decode(parts[2]), decode(parts[3]), decode(parts[4]), decode(parts[5]), decode(parts[6]), "1".equals(parts[7]), decode(parts[8]), decode(parts[9]));
                return;
            case "HYPERLINK":
                applyHyperlink(workbook, sheets.get(decode(parts[1])), decode(parts[2]), decode(parts[3]), decode(parts[4]), decode(parts[5]), "1".equals(parts[6]));
                return;
            case "COMMENT":
                applyComment(workbook, sheets.get(decode(parts[1])), decode(parts[2]), decode(parts[3]), decode(parts[4]));
                return;
            case "IMAGE":
                applyImage(workbook, sheets.get(decode(parts[1])), decode(parts[2]), decode(parts[3]));
                return;
            case "NAME":
                applyNamedRange(workbook, decode(parts[1]), decode(parts[2]), decode(parts[3]), decode(parts[4]));
                return;
            case "FREEZE":
                applyPane(sheets.get(decode(parts[1])), decode(parts[2]), Integer.parseInt(parts[3]), Integer.parseInt(parts[4]), decode(parts[5]));
                return;
            case "TABLE":
                applyTable(sheets.get(decode(parts[1])), decode(parts[2]), decode(parts[3]), decode(parts[4]), "1".equals(parts[5]), "1".equals(parts[6]));
                return;
            default:
                throw new IllegalArgumentException("Unsupported spec op: " + op);
        }
    }

    private static void applyValue(XSSFWorkbook workbook, XSSFSheet sheet, String address, String type, String rawValue, String rawFormula) {
        Cell cell = getCell(sheet, address);
        if ("blank".equals(type)) {
            cell.setBlank();
            return;
        }
        if ("string".equals(type)) {
            cell.setCellValue(rawValue);
            return;
        }
        if ("number".equals(type)) {
            cell.setCellValue(Double.parseDouble(rawValue));
            return;
        }
        if ("boolean".equals(type)) {
            cell.setCellValue(Boolean.parseBoolean(rawValue));
            return;
        }
        if ("formula".equals(type)) {
            String formula = rawFormula.isEmpty() ? rawValue : rawFormula;
            if (formula.startsWith("=")) {
                formula = formula.substring(1);
            }
            cell.setCellFormula(formula);
            return;
        }
        if ("error".equals(type)) {
            FormulaError error = FormulaError.forString(rawValue);
            cell.setCellErrorValue(error.getCode());
            return;
        }
        if ("date".equals(type)) {
            cell.setCellValue(Date.valueOf(LocalDate.parse(rawValue)));
            XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();
            style.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("yyyy-mm-dd"));
            cell.setCellStyle(style);
            return;
        }
        if ("datetime".equals(type)) {
            String normalized = rawValue.contains("T") ? rawValue : rawValue.replace(" ", "T");
            normalized = normalized.replace("Z", "");
            cell.setCellValue(Timestamp.valueOf(LocalDateTime.parse(normalized)));
            XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();
            style.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("yyyy-mm-dd hh:mm:ss"));
            cell.setCellStyle(style);
            return;
        }
        throw new IllegalArgumentException("Unsupported cell type: " + type);
    }

    private static void applyFormat(XSSFWorkbook workbook, XSSFSheet sheet, String address, String[] parts) {
        Cell cell = getCell(sheet, address);
        XSSFCellStyle style = cloneStyle(workbook, cell);
        XSSFFont font = cloneFont(workbook, style);
        if ("1".equals(parts[3])) {
            font.setBold(true);
        }
        if ("1".equals(parts[4])) {
            font.setItalic(true);
        }
        String underline = decode(parts[5]);
        if ("single".equals(underline)) {
            font.setUnderline(Font.U_SINGLE);
        } else if ("double".equals(underline)) {
            font.setUnderline(Font.U_DOUBLE);
        }
        if ("1".equals(parts[6])) {
            font.setStrikeout(true);
        }
        String fontName = decode(parts[7]);
        if (!fontName.isEmpty()) {
            font.setFontName(fontName);
        }
        String fontSize = decode(parts[8]);
        if (!fontSize.isEmpty()) {
            font.setFontHeight(Double.parseDouble(fontSize));
        }
        applyFontColor(font, decode(parts[9]));
        applyFillColor(style, decode(parts[10]));
        String numFmt = decode(parts[11]);
        if (!numFmt.isEmpty()) {
            style.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat(numFmt));
        }
        applyHorizontalAlignment(style, decode(parts[12]));
        applyVerticalAlignment(style, decode(parts[13]));
        if ("1".equals(parts[14])) {
            style.setWrapText(true);
        }
        String rotation = decode(parts[15]);
        if (!rotation.isEmpty()) {
            style.setRotation(Short.parseShort(rotation));
        }
        String indent = decode(parts[16]);
        if (!indent.isEmpty()) {
            style.setIndention(Short.parseShort(indent));
        }
        style.setFont(font);
        cell.setCellStyle(style);
    }

    private static void applyBorder(XSSFWorkbook workbook, XSSFSheet sheet, String address, String borderJson) {
        Cell cell = getCell(sheet, address);
        XSSFCellStyle style = cloneStyle(workbook, cell);
        applyBorderEdge(style, borderJson, "top");
        applyBorderEdge(style, borderJson, "bottom");
        applyBorderEdge(style, borderJson, "left");
        applyBorderEdge(style, borderJson, "right");
        applyDiagonalBorders(workbook, style, borderJson);
        cell.setCellStyle(style);
    }

    private static void applyConditionalFormat(XSSFSheet sheet, String rangeRef, String ruleType, String operator, String formula, boolean stopIfTrue, String bgColor) {
        SheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();
        CellRangeAddress[] regions = new CellRangeAddress[] {CellRangeAddress.valueOf(rangeRef)};
        ConditionalFormattingRule rule;
        if ("cellIs".equals(ruleType)) {
            rule = scf.createConditionalFormattingRule(mapComparisonOperator(operator), formula);
        } else if ("expression".equals(ruleType)) {
            rule = scf.createConditionalFormattingRule(formula);
        } else if ("dataBar".equals(ruleType) && scf instanceof XSSFSheetConditionalFormatting) {
            XSSFColor color = colorFromHex(bgColor == null || bgColor.isEmpty() ? "#638EC6" : bgColor);
            rule = ((XSSFSheetConditionalFormatting) scf).createConditionalFormattingRule(color);
        } else if ("colorScale".equals(ruleType) && scf instanceof XSSFSheetConditionalFormatting) {
            rule = ((XSSFSheetConditionalFormatting) scf).createConditionalFormattingColorScaleRule();
        } else {
            throw new IllegalArgumentException("Unsupported conditional formatting rule type: " + ruleType);
        }
        if (bgColor != null && !bgColor.isEmpty()) {
            PatternFormatting pattern = rule.createPatternFormatting();
            XSSFColor fillColor = colorFromHex(bgColor);
            if (pattern instanceof XSSFPatternFormatting) {
                ((XSSFPatternFormatting) pattern).setFillBackgroundColor(fillColor);
                ((XSSFPatternFormatting) pattern).setFillForegroundColor(fillColor);
            } else {
                short colorIndex = rgbToClosestIndexedColor(bgColor);
                pattern.setFillBackgroundColor(colorIndex);
                pattern.setFillForegroundColor(colorIndex);
            }
            pattern.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
        }
        scf.addConditionalFormatting(regions, rule);
        if (stopIfTrue) {
            CTConditionalFormatting[] blocks = sheet.getCTWorksheet().getConditionalFormattingArray();
            if (blocks.length > 0) {
                CTConditionalFormatting block = blocks[blocks.length - 1];
                CTCfRule[] cfRules = block.getCfRuleArray();
                if (cfRules.length > 0) {
                    cfRules[cfRules.length - 1].setStopIfTrue(true);
                }
            }
        }
    }

    private static void applyRowHeight(XSSFSheet sheet, int rowIndex, float height) {
        Row row = sheet.getRow(rowIndex - 1);
        if (row == null) {
            row = sheet.createRow(rowIndex - 1);
        }
        row.setHeightInPoints(height);
    }

    private static void applyColumnWidth(XSSFSheet sheet, String column, float width) {
        sheet.setColumnWidth(columnToIndex(column), Math.round(width * 256));
    }

    private static void applyValidation(XSSFSheet sheet, String address, String validationType, String operator, String formula1, String formula2, boolean allowBlank, String errorTitle, String errorMessage) {
        CellReference ref = new CellReference(address);
        DataValidationHelper helper = sheet.getDataValidationHelper();
        DataValidationConstraint constraint;
        if ("list".equals(validationType)) {
            if (formula1.startsWith("\"") && formula1.endsWith("\"")) {
                String body = formula1.substring(1, formula1.length() - 1);
                constraint = helper.createExplicitListConstraint(body.split(","));
            } else {
                constraint = helper.createFormulaListConstraint(formula1);
            }
        } else if ("custom".equals(validationType)) {
            constraint = helper.createCustomConstraint(formula1);
        } else if ("whole".equals(validationType)) {
            int poiOperator = "between".equals(operator) ? DataValidationConstraint.OperatorType.BETWEEN : DataValidationConstraint.OperatorType.IGNORED;
            constraint = helper.createIntegerConstraint(poiOperator, formula1, formula2);
        } else {
            return;
        }
        CellRangeAddressList range = new CellRangeAddressList(ref.getRow(), ref.getRow(), ref.getCol(), ref.getCol());
        DataValidation validation = helper.createValidation(constraint, range);
        validation.setEmptyCellAllowed(allowBlank);
        validation.setShowErrorBox(errorTitle != null && !errorTitle.isEmpty());
        if (errorTitle != null && !errorTitle.isEmpty()) {
            validation.createErrorBox(errorTitle, errorMessage == null ? "" : errorMessage);
        }
        sheet.addValidationData(validation);
    }

    private static void applyHyperlink(XSSFWorkbook workbook, XSSFSheet sheet, String address, String target, String label, String tooltip, boolean internal) {
        CreationHelper helper = workbook.getCreationHelper();
        Hyperlink link = helper.createHyperlink(internal ? HyperlinkType.DOCUMENT : HyperlinkType.URL);
        link.setAddress(target);
        if (tooltip != null && !tooltip.isEmpty() && link instanceof XSSFHyperlink) {
            ((XSSFHyperlink) link).setTooltip(tooltip);
        }
        Cell cell = getCell(sheet, address);
        cell.setCellValue(label.isEmpty() ? target : label);
        cell.setHyperlink(link);
    }

    private static void applyComment(XSSFWorkbook workbook, XSSFSheet sheet, String address, String text, String author) {
        CreationHelper helper = workbook.getCreationHelper();
        Drawing<?> drawing = sheet.createDrawingPatriarch();
        ClientAnchor anchor = helper.createClientAnchor();
        Cell cell = getCell(sheet, address);
        anchor.setCol1(cell.getColumnIndex() + 1);
        anchor.setCol2(cell.getColumnIndex() + 4);
        anchor.setRow1(cell.getRowIndex());
        anchor.setRow2(cell.getRowIndex() + 3);
        Comment comment = drawing.createCellComment(anchor);
        comment.setString(helper.createRichTextString(text));
        comment.setAuthor(author);
        cell.setCellComment(comment);
    }

    private static void applyImage(XSSFWorkbook workbook, XSSFSheet sheet, String address, String path) {
        if (path == null || path.isEmpty()) {
            return;
        }
        byte[] imageBytes;
        try {
            imageBytes = Files.readAllBytes(Path.of(path));
        } catch (Exception e) {
            throw new IllegalArgumentException("Failed to read image path: " + path, e);
        }
        int pictureType = pictureTypeFromPath(path);
        int pictureIdx = workbook.addPicture(imageBytes, pictureType);
        CreationHelper helper = workbook.getCreationHelper();
        Drawing<?> drawing = sheet.createDrawingPatriarch();
        ClientAnchor anchor = helper.createClientAnchor();
        CellReference ref = new CellReference(address);
        anchor.setCol1(ref.getCol());
        anchor.setRow1(ref.getRow());
        anchor.setCol2(ref.getCol() + 2);
        anchor.setRow2(ref.getRow() + 3);
        drawing.createPicture(anchor, pictureIdx);
    }

    private static void applyPane(XSSFSheet sheet, String mode, int xSplit, int ySplit, String topLeftCell) {
        if ("split".equals(mode)) {
            if (sheet.getCTWorksheet().getSheetViews() == null || sheet.getCTWorksheet().getSheetViews().sizeOfSheetViewArray() == 0) {
                sheet.getCTWorksheet().addNewSheetViews().addNewSheetView();
            }
            CTPane pane = sheet.getCTWorksheet().getSheetViews().getSheetViewArray(0).isSetPane()
                ? sheet.getCTWorksheet().getSheetViews().getSheetViewArray(0).getPane()
                : sheet.getCTWorksheet().getSheetViews().getSheetViewArray(0).addNewPane();
            pane.setState(STPaneState.SPLIT);
            pane.setXSplit(xSplit);
            pane.setYSplit(ySplit);
            if (topLeftCell != null && !topLeftCell.isEmpty()) {
                pane.setTopLeftCell(topLeftCell);
            }
            return;
        }
        sheet.createFreezePane(xSplit, ySplit);
    }

    private static void applyNamedRange(XSSFWorkbook workbook, String sheetName, String name, String scope, String refersTo) {
        if (name == null || name.isEmpty() || refersTo == null || refersTo.isEmpty()) {
            return;
        }
        org.apache.poi.ss.usermodel.Name definedName = workbook.createName();
        definedName.setNameName(name);
        definedName.setRefersToFormula(refersTo);
        if ("sheet".equals(scope)) {
            int sheetIndex = workbook.getSheetIndex(sheetName);
            if (sheetIndex >= 0) {
                definedName.setSheetIndex(sheetIndex);
            }
        }
    }

    private static void applyTable(XSSFSheet sheet, String ref, String name, String styleName, boolean totalsRow, boolean autoFilter) {
        AreaReference area = new AreaReference(ref, SpreadsheetVersion.EXCEL2007);
        XSSFTable table = sheet.createTable(area);
        CTTable ctTable = table.getCTTable();
        ctTable.setId(1);
        ctTable.setName(name);
        ctTable.setDisplayName(name);
        ctTable.setRef(ref);
        ctTable.setTotalsRowShown(totalsRow);
        if (totalsRow) {
            ctTable.setTotalsRowCount(1);
        }
        if (autoFilter) {
            ctTable.addNewAutoFilter().setRef(ref);
        }
        if (styleName != null && !styleName.isEmpty()) {
            CTTableStyleInfo style = ctTable.isSetTableStyleInfo() ? ctTable.getTableStyleInfo() : ctTable.addNewTableStyleInfo();
            style.setName(styleName);
            style.setShowRowStripes(true);
        }
    }

    private static XSSFCellStyle cloneStyle(XSSFWorkbook workbook, Cell cell) {
        XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();
        if (cell.getCellStyle() != null) {
            style.cloneStyleFrom(cell.getCellStyle());
        }
        return style;
    }

    private static XSSFFont cloneFont(XSSFWorkbook workbook, XSSFCellStyle style) {
        XSSFFont font = workbook.createFont();
        if (style.getFontIndexAsInt() > 0) {
            XSSFFont existing = workbook.getFontAt(style.getFontIndexAsInt());
            font.setBold(existing.getBold());
            font.setItalic(existing.getItalic());
            font.setFontHeight(existing.getFontHeight());
            font.setFontName(existing.getFontName());
            font.setUnderline(existing.getUnderline());
            if (existing.getXSSFColor() != null) {
                font.setColor(existing.getXSSFColor());
            }
        }
        return font;
    }

    private static void applyFontColor(XSSFFont font, String hex) {
        if (hex == null || hex.isEmpty()) {
            return;
        }
        font.setColor(colorFromHex(hex));
    }

    private static void applyFillColor(XSSFCellStyle style, String hex) {
        if (hex == null || hex.isEmpty()) {
            return;
        }
        style.setFillForegroundColor(colorFromHex(hex));
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }

    private static void applyHorizontalAlignment(XSSFCellStyle style, String value) {
        if (value == null || value.isEmpty()) {
            return;
        }
        if ("left".equals(value)) {
            style.setAlignment(HorizontalAlignment.LEFT);
        } else if ("center".equals(value)) {
            style.setAlignment(HorizontalAlignment.CENTER);
        } else if ("right".equals(value)) {
            style.setAlignment(HorizontalAlignment.RIGHT);
        } else if ("justify".equals(value)) {
            style.setAlignment(HorizontalAlignment.JUSTIFY);
        }
    }

    private static void applyVerticalAlignment(XSSFCellStyle style, String value) {
        if (value == null || value.isEmpty()) {
            return;
        }
        if ("top".equals(value)) {
            style.setVerticalAlignment(VerticalAlignment.TOP);
        } else if ("center".equals(value)) {
            style.setVerticalAlignment(VerticalAlignment.CENTER);
        } else if ("bottom".equals(value)) {
            style.setVerticalAlignment(VerticalAlignment.BOTTOM);
        }
    }

    private static void applyBorderEdge(XSSFCellStyle style, String borderJson, String edgeName) {
        String marker = "\"" + edgeName + "\":{";
        if (!borderJson.contains(marker)) {
            return;
        }
        String styleValue = extractBorderStyle(borderJson, edgeName);
        BorderStyle poiStyle = mapBorderStyle(styleValue);
        String colorValue = extractBorderColor(borderJson, edgeName);
        if ("top".equals(edgeName)) {
            style.setBorderTop(poiStyle);
            if (colorValue != null) {
                style.setTopBorderColor(colorFromHex(colorValue));
            }
        } else if ("bottom".equals(edgeName)) {
            style.setBorderBottom(poiStyle);
            if (colorValue != null) {
                style.setBottomBorderColor(colorFromHex(colorValue));
            }
        } else if ("left".equals(edgeName)) {
            style.setBorderLeft(poiStyle);
            if (colorValue != null) {
                style.setLeftBorderColor(colorFromHex(colorValue));
            }
        } else if ("right".equals(edgeName)) {
            style.setBorderRight(poiStyle);
            if (colorValue != null) {
                style.setRightBorderColor(colorFromHex(colorValue));
            }
        }
    }

    private static void applyDiagonalBorders(XSSFWorkbook workbook, XSSFCellStyle style, String borderJson) {
        String upStyle = extractBorderStyle(borderJson, "diagonal_up");
        String downStyle = extractBorderStyle(borderJson, "diagonal_down");
        String upColor = extractBorderColor(borderJson, "diagonal_up");
        String downColor = extractBorderColor(borderJson, "diagonal_down");
        boolean hasUp = upStyle != null && !"none".equals(upStyle);
        boolean hasDown = downStyle != null && !"none".equals(downStyle);
        if (!hasUp && !hasDown) {
            return;
        }

        StylesTable styles = workbook.getStylesSource();
        XSSFCellBorder border = new XSSFCellBorder();
        border.setBorderStyle(BorderSide.TOP, style.getBorderTop());
        border.setBorderStyle(BorderSide.BOTTOM, style.getBorderBottom());
        border.setBorderStyle(BorderSide.LEFT, style.getBorderLeft());
        border.setBorderStyle(BorderSide.RIGHT, style.getBorderRight());
        XSSFColor topColor = style.getTopBorderXSSFColor();
        XSSFColor bottomColor = style.getBottomBorderXSSFColor();
        XSSFColor leftColor = style.getLeftBorderXSSFColor();
        XSSFColor rightColor = style.getRightBorderXSSFColor();
        if (topColor != null) border.setBorderColor(BorderSide.TOP, topColor);
        if (bottomColor != null) border.setBorderColor(BorderSide.BOTTOM, bottomColor);
        if (leftColor != null) border.setBorderColor(BorderSide.LEFT, leftColor);
        if (rightColor != null) border.setBorderColor(BorderSide.RIGHT, rightColor);

        String diagonalStyle = hasUp ? upStyle : downStyle;
        border.setBorderStyle(BorderSide.DIAGONAL, mapBorderStyle(diagonalStyle));
        String diagonalColor = upColor != null ? upColor : downColor;
        if (diagonalColor != null) {
            border.setBorderColor(BorderSide.DIAGONAL, colorFromHex(diagonalColor));
        }
        border.getCTBorder().setDiagonalUp(hasUp);
        border.getCTBorder().setDiagonalDown(hasDown);

        long borderId = styles.putBorder(border);
        style.getCoreXf().setBorderId(borderId);
        style.getCoreXf().setApplyBorder(true);
    }

    private static String extractBorderStyle(String borderJson, String edgeName) {
        String marker = "\"" + edgeName + "\":{";
        int start = borderJson.indexOf(marker);
        if (start < 0) {
            return "none";
        }
        int styleIndex = borderJson.indexOf("\"style\":\"", start);
        if (styleIndex < 0) {
            return "none";
        }
        int valueStart = styleIndex + 9;
        int valueEnd = borderJson.indexOf('"', valueStart);
        if (valueEnd < 0) {
            return "none";
        }
        return borderJson.substring(valueStart, valueEnd);
    }

    private static String extractBorderColor(String borderJson, String edgeName) {
        String marker = "\"" + edgeName + "\":{";
        int start = borderJson.indexOf(marker);
        if (start < 0) {
            return null;
        }
        int colorIndex = borderJson.indexOf("\"color\":\"", start);
        if (colorIndex < 0) {
            return null;
        }
        int valueStart = colorIndex + 9;
        int valueEnd = borderJson.indexOf('"', valueStart);
        if (valueEnd < 0) {
            return null;
        }
        return borderJson.substring(valueStart, valueEnd);
    }

    private static BorderStyle mapBorderStyle(String style) {
        if ("thin".equals(style)) return BorderStyle.THIN;
        if ("medium".equals(style)) return BorderStyle.MEDIUM;
        if ("thick".equals(style)) return BorderStyle.THICK;
        if ("double".equals(style)) return BorderStyle.DOUBLE;
        if ("dashed".equals(style)) return BorderStyle.DASHED;
        if ("dotted".equals(style)) return BorderStyle.DOTTED;
        if ("hair".equals(style)) return BorderStyle.HAIR;
        if ("mediumDashed".equals(style)) return BorderStyle.MEDIUM_DASHED;
        if ("dashDot".equals(style)) return BorderStyle.DASH_DOT;
        if ("mediumDashDot".equals(style)) return BorderStyle.MEDIUM_DASH_DOT;
        if ("dashDotDot".equals(style)) return BorderStyle.DASH_DOT_DOT;
        if ("mediumDashDotDot".equals(style)) return BorderStyle.MEDIUM_DASH_DOT_DOT;
        if ("slantDashDot".equals(style)) return BorderStyle.SLANTED_DASH_DOT;
        return BorderStyle.NONE;
    }

    private static int columnToIndex(String column) {
        int col = 0;
        for (int i = 0; i < column.length(); i++) {
            col = col * 26 + (column.charAt(i) - 'A' + 1);
        }
        return col - 1;
    }

    private static XSSFColor colorFromHex(String hex) {
        String normalized = hex.startsWith("#") ? hex.substring(1) : hex;
        int r = Integer.parseInt(normalized.substring(0, 2), 16);
        int g = Integer.parseInt(normalized.substring(2, 4), 16);
        int b = Integer.parseInt(normalized.substring(4, 6), 16);
        return new XSSFColor(new byte[] {(byte) r, (byte) g, (byte) b}, new DefaultIndexedColorMap());
    }

    private static String decode(String value) {
        if (value == null || value.isEmpty()) {
            return "";
        }
        return new String(Base64.getDecoder().decode(value), StandardCharsets.UTF_8);
    }

    private static int pictureTypeFromPath(String path) {
        String lower = path.toLowerCase();
        if (lower.endsWith(".jpg") || lower.endsWith(".jpeg")) {
            return Workbook.PICTURE_TYPE_JPEG;
        }
        return Workbook.PICTURE_TYPE_PNG;
    }

    private static byte mapComparisonOperator(String operator) {
        if ("greaterThan".equals(operator)) return ComparisonOperator.GT;
        if ("greaterThanOrEqual".equals(operator)) return ComparisonOperator.GE;
        if ("lessThan".equals(operator)) return ComparisonOperator.LT;
        if ("lessThanOrEqual".equals(operator)) return ComparisonOperator.LE;
        if ("between".equals(operator)) return ComparisonOperator.BETWEEN;
        if ("equal".equals(operator)) return ComparisonOperator.EQUAL;
        return ComparisonOperator.NO_COMPARISON;
    }

    private static short rgbToClosestIndexedColor(String hex) {
        String normalized = hex.startsWith("#") ? hex.substring(1) : hex;
        if (normalized.equalsIgnoreCase("FFFF00")) return org.apache.poi.ss.usermodel.IndexedColors.YELLOW.getIndex();
        if (normalized.equalsIgnoreCase("FF00FF")) return org.apache.poi.ss.usermodel.IndexedColors.VIOLET.getIndex();
        if (normalized.equalsIgnoreCase("FF0000")) return org.apache.poi.ss.usermodel.IndexedColors.RED.getIndex();
        if (normalized.equalsIgnoreCase("00FF00")) return org.apache.poi.ss.usermodel.IndexedColors.BRIGHT_GREEN.getIndex();
        return org.apache.poi.ss.usermodel.IndexedColors.YELLOW.getIndex();
    }

    private static Cell getCell(XSSFSheet sheet, String address) {
        CellReference ref = new CellReference(address);
        Row row = sheet.getRow(ref.getRow());
        if (row == null) {
            row = sheet.createRow(ref.getRow());
        }
        Cell cell = row.getCell(ref.getCol());
        if (cell == null) {
            cell = row.createCell(ref.getCol(), CellType.BLANK);
        }
        return cell;
    }

    private static String json(String value) {
        return value.replace("\\", "\\\\").replace("\"", "\\\"");
    }
}
