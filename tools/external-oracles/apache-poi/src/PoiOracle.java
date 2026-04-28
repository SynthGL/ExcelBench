import java.io.File;
import java.io.FileOutputStream;
import java.nio.charset.StandardCharsets;
import java.util.Base64;

import org.apache.poi.common.usermodel.HyperlinkType;
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
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumn;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumns;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableStyleInfo;

public final class PoiOracle {
    private static final String PIXEL_PNG_BASE64 =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";

    private PoiOracle() {}

    public static void main(String[] args) throws Exception {
        if (args.length == 1 && "--self-test".equals(args[0])) {
            System.out.println("OK");
            return;
        }
        if (args.length != 3 || !"write_fixture".equals(args[0])) {
            throw new IllegalArgumentException("Usage: PoiOracle write_fixture <fixture_id> <output_path>");
        }
        writeFixture(args[1], args[2]);
    }

    private static void writeFixture(String fixtureId, String outputPath) throws Exception {
        File output = new File(outputPath);
        File parent = output.getParentFile();
        if (parent != null) {
            parent.mkdirs();
        }

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            workbook.getProperties().getCoreProperties().setCreator("ExcelBench Apache POI Oracle");
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
            + "\"protected_sheets\":1"
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
