package main

import (
	"archive/zip"
	"bytes"
	"encoding/base64"
	"encoding/json"
	"errors"
	"fmt"
	_ "image/png"
	"io"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

type oracleRequest struct {
	FixtureID  string          `json:"fixture_id"`
	Operation  string          `json:"operation"`
	Payload    json.RawMessage `json:"payload"`
	InputPath  string          `json:"input_path,omitempty"`
	OutputPath string          `json:"output_path,omitempty"`
}

type writePayload struct {
	Sheets             []sheetSpec             `json:"sheets"`
	Cells              []cellSpec              `json:"cells"`
	Formats            []formatSpec            `json:"formats"`
	Borders            []borderSpec            `json:"borders"`
	Columns            []columnSpec            `json:"columns"`
	RowHeights         []rowHeightSpec         `json:"row_heights"`
	Merges             []mergeSpec             `json:"merges"`
	Validations        []validationSpec        `json:"validations"`
	Hyperlinks         []hyperlinkSpec         `json:"hyperlinks"`
	Comments           []commentSpec           `json:"comments"`
	Panes              []paneSpec              `json:"panes"`
	NamedRanges        []namedRangeSpec        `json:"named_ranges"`
	Tables             []tableSpec             `json:"tables"`
	ConditionalFormats []conditionalFormatSpec `json:"conditional_formats"`
	Charts             []chartSpec             `json:"charts"`
	Pivots             []pivotSpec             `json:"pivots"`
	Slicers            []slicerSpec            `json:"slicers"`
	Pictures           []pictureSpec           `json:"pictures"`
}

type sheetSpec struct {
	Name string `json:"name"`
}

type cellSpec struct {
	Sheet   string      `json:"sheet"`
	Cell    string      `json:"cell"`
	Type    string      `json:"type"`
	Value   interface{} `json:"value"`
	Formula string      `json:"formula"`
}

type formatSpec struct {
	Sheet         string   `json:"sheet"`
	Cell          string   `json:"cell"`
	Bold          *bool    `json:"bold"`
	Italic        *bool    `json:"italic"`
	Underline     string   `json:"underline"`
	Strikethrough *bool    `json:"strikethrough"`
	FontName      string   `json:"font_name"`
	FontSize      *float64 `json:"font_size"`
	FontColor     string   `json:"font_color"`
	BGColor       string   `json:"bg_color"`
	NumberFormat  string   `json:"number_format"`
	HAlign        string   `json:"h_align"`
	VAlign        string   `json:"v_align"`
	Wrap          *bool    `json:"wrap"`
	Rotation      *int     `json:"rotation"`
	Indent        *int     `json:"indent"`
}

type columnSpec struct {
	Sheet string  `json:"sheet"`
	Start string  `json:"start"`
	End   string  `json:"end"`
	Width float64 `json:"width"`
}

type borderSpec struct {
	Sheet  string                 `json:"sheet"`
	Cell   string                 `json:"cell"`
	Border map[string]interface{} `json:"border"`
}

type rowHeightSpec struct {
	Sheet  string  `json:"sheet"`
	Row    int     `json:"row"`
	Height float64 `json:"height"`
}

type mergeSpec struct {
	Sheet string `json:"sheet"`
	Range string `json:"range"`
}

type validationSpec struct {
	Sheet          string `json:"sheet"`
	Range          string `json:"range"`
	ValidationType string `json:"validation_type"`
	Operator       string `json:"operator"`
	Formula1       string `json:"formula1"`
	Formula2       string `json:"formula2"`
	AllowBlank     *bool  `json:"allow_blank"`
	ErrorTitle     string `json:"error_title"`
	Error          string `json:"error"`
}

type hyperlinkSpec struct {
	Sheet    string `json:"sheet"`
	Cell     string `json:"cell"`
	Target   string `json:"target"`
	Display  string `json:"display"`
	Tooltip  string `json:"tooltip"`
	Internal bool   `json:"internal"`
}

type commentSpec struct {
	Sheet  string `json:"sheet"`
	Cell   string `json:"cell"`
	Text   string `json:"text"`
	Author string `json:"author"`
}

type paneSpec struct {
	Sheet       string `json:"sheet"`
	Mode        string `json:"mode"`
	XSplit      int    `json:"x_split"`
	YSplit      int    `json:"y_split"`
	TopLeftCell string `json:"top_left_cell"`
}

type namedRangeSpec struct {
	Sheet    string `json:"sheet"`
	Name     string `json:"name"`
	Scope    string `json:"scope"`
	RefersTo string `json:"refers_to"`
}

type tableSpec struct {
	Sheet          string   `json:"sheet"`
	Range          string   `json:"range"`
	Name           string   `json:"name"`
	Style          string   `json:"style"`
	ShowHeaderRow  *bool    `json:"show_header_row"`
	ShowRowStripes *bool    `json:"show_row_stripes"`
	TotalsRow      bool     `json:"totals_row"`
	Columns        []string `json:"columns"`
	AutoFilter     bool     `json:"autofilter"`
}

type conditionalFormatSpec struct {
	Sheet          string `json:"sheet"`
	Range          string `json:"range"`
	Type           string `json:"type"`
	Criteria       string `json:"criteria"`
	Value          string `json:"value"`
	BGColor        string `json:"bg_color"`
	MinType        string `json:"min_type"`
	MidType        string `json:"mid_type"`
	MaxType        string `json:"max_type"`
	MinColor       string `json:"min_color"`
	MidColor       string `json:"mid_color"`
	MaxColor       string `json:"max_color"`
	BarColor       string `json:"bar_color"`
	BarBorderColor string `json:"bar_border_color"`
	IconStyle      string `json:"icon_style"`
	StopIfTrue     bool   `json:"stop_if_true"`
}

type chartSpec struct {
	Sheet      string          `json:"sheet"`
	Cell       string          `json:"cell"`
	Type       string          `json:"type"`
	Title      string          `json:"title"`
	Name       string          `json:"name"`
	AltText    string          `json:"alt_text"`
	Width      uint            `json:"width"`
	Height     uint            `json:"height"`
	Categories string          `json:"categories"`
	Values     string          `json:"values"`
	Series     []seriesSpec    `json:"series"`
	ShowValues bool            `json:"show_values"`
	DataPoints []dataPointSpec `json:"data_points"`
	VaryColors *bool           `json:"vary_colors"`
}

type seriesSpec struct {
	Name       string          `json:"name"`
	Categories string          `json:"categories"`
	Values     string          `json:"values"`
	FillColor  string          `json:"fill_color"`
	DataPoints []dataPointSpec `json:"data_points"`
}

type dataPointSpec struct {
	Index     int    `json:"index"`
	FillColor string `json:"fill_color"`
}

type pivotSpec struct {
	DataRange      string       `json:"data_range"`
	Range          string       `json:"range"`
	Name           string       `json:"name"`
	Rows           []pivotField `json:"rows"`
	Columns        []pivotField `json:"columns"`
	Data           []pivotField `json:"data"`
	Filters        []pivotField `json:"filters"`
	Style          string       `json:"style"`
	RowGrandTotals *bool        `json:"row_grand_totals"`
	ColGrandTotals *bool        `json:"col_grand_totals"`
	ShowRowStripes bool         `json:"show_row_stripes"`
	ShowColStripes bool         `json:"show_col_stripes"`
}

type pivotField struct {
	Name     string `json:"name"`
	Data     string `json:"data"`
	Subtotal string `json:"subtotal"`
	NumFmt   int    `json:"num_fmt"`
}

type slicerSpec struct {
	Sheet         string `json:"sheet"`
	Name          string `json:"name"`
	Cell          string `json:"cell"`
	TableSheet    string `json:"table_sheet"`
	TableName     string `json:"table_name"`
	Caption       string `json:"caption"`
	Width         uint   `json:"width"`
	Height        uint   `json:"height"`
	DisplayHeader *bool  `json:"display_header"`
}

type pictureSpec struct {
	Sheet       string  `json:"sheet"`
	Cell        string  `json:"cell"`
	Extension   string  `json:"extension"`
	Base64      string  `json:"base64"`
	Name        string  `json:"name"`
	AltText     string  `json:"alt_text"`
	ScaleX      float64 `json:"scale_x"`
	ScaleY      float64 `json:"scale_y"`
	Positioning string  `json:"positioning"`
}

func main() {
	if err := run(os.Stdin, os.Stdout); err != nil {
		_ = json.NewEncoder(os.Stdout).Encode(map[string]interface{}{
			"error":   "excelize_oracle_failed",
			"message": err.Error(),
		})
		os.Exit(1)
	}
}

func run(input io.Reader, output io.Writer) error {
	var request oracleRequest
	decoder := json.NewDecoder(input)
	decoder.UseNumber()
	if err := decoder.Decode(&request); err != nil {
		return fmt.Errorf("decode request: %w", err)
	}

	switch request.Operation {
	case "write_fixture":
		return writeFixture(output, request)
	case "read_metadata":
		return readMetadata(output, request)
	default:
		return fmt.Errorf("unsupported operation %q", request.Operation)
	}
}

func writeFixture(output io.Writer, request oracleRequest) error {
	if request.OutputPath == "" {
		return errors.New("write_fixture requires output_path")
	}
	var payload writePayload
	if err := json.Unmarshal(request.Payload, &payload); err != nil {
		return fmt.Errorf("decode write payload: %w", err)
	}

	workbook := excelize.NewFile()
	defer workbook.Close()
	if err := configureSheets(workbook, payload.Sheets); err != nil {
		return err
	}
	if err := applyCells(workbook, payload.Cells); err != nil {
		return err
	}
	if err := applyFormats(workbook, payload.Formats); err != nil {
		return err
	}
	if err := applyBorders(workbook, payload.Borders); err != nil {
		return err
	}
	if err := applyColumns(workbook, payload.Columns); err != nil {
		return err
	}
	if err := applyRowHeights(workbook, payload.RowHeights); err != nil {
		return err
	}
	if err := applyMerges(workbook, payload.Merges); err != nil {
		return err
	}
	if err := applyDataValidations(workbook, payload.Validations); err != nil {
		return err
	}
	if err := applyHyperlinks(workbook, payload.Hyperlinks); err != nil {
		return err
	}
	if err := applyComments(workbook, payload.Comments); err != nil {
		return err
	}
	if err := applyPanes(workbook, payload.Panes); err != nil {
		return err
	}
	if err := applyNamedRanges(workbook, payload.NamedRanges); err != nil {
		return err
	}
	if err := applyTables(workbook, payload.Tables); err != nil {
		return err
	}
	if err := applyConditionalFormats(workbook, payload.ConditionalFormats); err != nil {
		return err
	}
	if err := applyCharts(workbook, payload.Charts); err != nil {
		return err
	}
	if err := applyPivots(workbook, payload.Pivots); err != nil {
		return err
	}
	if err := applySlicers(workbook, payload.Slicers); err != nil {
		return err
	}
	if err := applyPictures(workbook, payload.Pictures); err != nil {
		return err
	}

	if err := os.MkdirAll(filepath.Dir(request.OutputPath), 0o755); err != nil {
		return fmt.Errorf("create output dir: %w", err)
	}
	if err := workbook.SaveAs(request.OutputPath); err != nil {
		return fmt.Errorf("save workbook: %w", err)
	}
	if err := normalizeExcelizeConditionalFormats(request.OutputPath, payload.ConditionalFormats); err != nil {
		return err
	}
	if err := normalizeExcelizeTables(request.OutputPath, payload.Tables); err != nil {
		return err
	}

	return json.NewEncoder(output).Encode(map[string]interface{}{
		"fixture_id":  request.FixtureID,
		"operation":   request.Operation,
		"output_path": request.OutputPath,
		"tool":        "excelize",
		"counts": map[string]int{
			"sheets":              len(workbook.GetSheetList()),
			"cells":               len(payload.Cells),
			"formats":             len(payload.Formats),
			"borders":             len(payload.Borders),
			"merges":              len(payload.Merges),
			"row_heights":         len(payload.RowHeights),
			"validations":         len(payload.Validations),
			"hyperlinks":          len(payload.Hyperlinks),
			"comments":            len(payload.Comments),
			"panes":               len(payload.Panes),
			"named_ranges":        len(payload.NamedRanges),
			"tables":              len(payload.Tables),
			"conditional_formats": len(payload.ConditionalFormats),
			"charts":              len(payload.Charts),
			"pivots":              len(payload.Pivots),
			"slicers":             len(payload.Slicers),
			"pictures":            len(payload.Pictures),
		},
	})
}

func readMetadata(output io.Writer, request oracleRequest) error {
	inputPath := request.InputPath
	if inputPath == "" {
		inputPath = request.OutputPath
	}
	if inputPath == "" {
		return errors.New("read_metadata requires input_path")
	}
	workbook, err := excelize.OpenFile(inputPath)
	if err != nil {
		return fmt.Errorf("open workbook: %w", err)
	}
	defer workbook.Close()

	sheetInfo := make([]map[string]interface{}, 0)
	for _, sheet := range workbook.GetSheetList() {
		tables, _ := workbook.GetTables(sheet)
		pivots, _ := workbook.GetPivotTables(sheet)
		slicers, _ := workbook.GetSlicers(sheet)
		cf, _ := workbook.GetConditionalFormats(sheet)
		sheetInfo = append(sheetInfo, map[string]interface{}{
			"name":                sheet,
			"tables":              len(tables),
			"pivots":              len(pivots),
			"slicers":             len(slicers),
			"conditional_formats": len(cf),
		})
	}
	return json.NewEncoder(output).Encode(map[string]interface{}{
		"fixture_id": request.FixtureID,
		"operation":  request.Operation,
		"tool":       "excelize",
		"sheets":     sheetInfo,
	})
}

func configureSheets(workbook *excelize.File, sheets []sheetSpec) error {
	if len(sheets) == 0 {
		return nil
	}
	for i, sheet := range sheets {
		if sheet.Name == "" {
			return errors.New("sheet name is required")
		}
		if i == 0 {
			if err := workbook.SetSheetName("Sheet1", sheet.Name); err != nil {
				return fmt.Errorf("rename Sheet1: %w", err)
			}
			continue
		}
		if _, err := workbook.NewSheet(sheet.Name); err != nil {
			return fmt.Errorf("add sheet %q: %w", sheet.Name, err)
		}
	}
	return nil
}

func applyCells(workbook *excelize.File, cells []cellSpec) error {
	for _, cell := range cells {
		if cell.Sheet == "" || cell.Cell == "" {
			return errors.New("cell entries require sheet and cell")
		}
		if cell.Type == "blank" {
			continue
		}
		if cell.Type == "formula" || cell.Formula != "" {
			formula := cell.Formula
			if formula == "" {
				formula = stringify(cell.Value)
			}
			formula = trimFormulaEquals(formula)
			if err := workbook.SetCellFormula(cell.Sheet, cell.Cell, formula); err != nil {
				return fmt.Errorf("set formula %s!%s: %w", cell.Sheet, cell.Cell, err)
			}
			continue
		}
		if cell.Type == "date" || cell.Type == "datetime" {
			parsed, err := parseTemporalCell(cell.Type, cell.Value)
			if err != nil {
				return fmt.Errorf("parse %s %s!%s: %w", cell.Type, cell.Sheet, cell.Cell, err)
			}
			if err := workbook.SetCellValue(cell.Sheet, cell.Cell, parsed); err != nil {
				return fmt.Errorf("set %s %s!%s: %w", cell.Type, cell.Sheet, cell.Cell, err)
			}
			continue
		}
		if cell.Type == "error" {
			if err := workbook.SetCellValue(cell.Sheet, cell.Cell, stringify(cell.Value)); err != nil {
				return fmt.Errorf("set error %s!%s: %w", cell.Sheet, cell.Cell, err)
			}
			continue
		}
		value := normalizeValue(cell.Value)
		if err := workbook.SetCellValue(cell.Sheet, cell.Cell, value); err != nil {
			return fmt.Errorf("set value %s!%s: %w", cell.Sheet, cell.Cell, err)
		}
	}
	return nil
}

func applyBorders(workbook *excelize.File, borders []borderSpec) error {
	for _, spec := range borders {
		if spec.Sheet == "" || spec.Cell == "" {
			return errors.New("border entries require sheet and cell")
		}
		style := &excelize.Style{Border: buildBorders(spec.Border)}
		styleID, err := workbook.NewStyle(style)
		if err != nil {
			return fmt.Errorf("new border style %s!%s: %w", spec.Sheet, spec.Cell, err)
		}
		if err := workbook.SetCellStyle(spec.Sheet, spec.Cell, spec.Cell, styleID); err != nil {
			return fmt.Errorf("set border style %s!%s: %w", spec.Sheet, spec.Cell, err)
		}
	}
	return nil
}

func applyFormats(workbook *excelize.File, formats []formatSpec) error {
	for _, spec := range formats {
		if spec.Sheet == "" || spec.Cell == "" {
			return errors.New("format entries require sheet and cell")
		}
		style := &excelize.Style{}
		var hasStyle bool
		if spec.Bold != nil || spec.Italic != nil || spec.Underline != "" || spec.Strikethrough != nil || spec.FontName != "" || spec.FontSize != nil || spec.FontColor != "" {
			font := &excelize.Font{}
			if spec.Bold != nil {
				font.Bold = *spec.Bold
			}
			if spec.Italic != nil {
				font.Italic = *spec.Italic
			}
			if spec.Underline != "" {
				font.Underline = spec.Underline
			}
			if spec.Strikethrough != nil {
				font.Strike = *spec.Strikethrough
			}
			if spec.FontName != "" {
				font.Family = spec.FontName
			}
			if spec.FontSize != nil {
				font.Size = *spec.FontSize
			}
			if spec.FontColor != "" {
				font.Color = trimHash(spec.FontColor)
			}
			style.Font = font
			hasStyle = true
		}
		if spec.BGColor != "" {
			style.Fill = excelize.Fill{Type: "pattern", Pattern: 1, Color: []string{trimHash(spec.BGColor)}}
			hasStyle = true
		}
		if spec.HAlign != "" || spec.VAlign != "" || spec.Wrap != nil || spec.Rotation != nil || spec.Indent != nil {
			align := &excelize.Alignment{}
			if spec.HAlign != "" {
				align.Horizontal = spec.HAlign
			}
			if spec.VAlign != "" {
				align.Vertical = spec.VAlign
			}
			if spec.Wrap != nil {
				align.WrapText = *spec.Wrap
			}
			if spec.Rotation != nil {
				align.TextRotation = *spec.Rotation
			}
			if spec.Indent != nil {
				align.Indent = *spec.Indent
			}
			style.Alignment = align
			hasStyle = true
		}
		if spec.NumberFormat != "" {
			if numFmt, ok := builtinNumFmt(spec.NumberFormat); ok {
				style.NumFmt = numFmt
			} else {
				custom := spec.NumberFormat
				style.CustomNumFmt = &custom
			}
			hasStyle = true
		}
		if !hasStyle {
			continue
		}
		styleID, err := workbook.NewStyle(style)
		if err != nil {
			return fmt.Errorf("new style %s!%s: %w", spec.Sheet, spec.Cell, err)
		}
		if err := workbook.SetCellStyle(spec.Sheet, spec.Cell, spec.Cell, styleID); err != nil {
			return fmt.Errorf("set style %s!%s: %w", spec.Sheet, spec.Cell, err)
		}
	}
	return nil
}

func applyColumns(workbook *excelize.File, columns []columnSpec) error {
	for _, column := range columns {
		if column.Sheet == "" || column.Start == "" || column.End == "" {
			return errors.New("column entries require sheet, start, and end")
		}
		if err := workbook.SetColWidth(column.Sheet, column.Start, column.End, column.Width); err != nil {
			return fmt.Errorf("set column width %s!%s:%s: %w", column.Sheet, column.Start, column.End, err)
		}
	}
	return nil
}

func applyRowHeights(workbook *excelize.File, rows []rowHeightSpec) error {
	for _, spec := range rows {
		if spec.Sheet == "" || spec.Row <= 0 {
			return errors.New("row height entries require sheet and row")
		}
		if err := workbook.SetRowHeight(spec.Sheet, spec.Row, spec.Height); err != nil {
			return fmt.Errorf("set row height %s!%d: %w", spec.Sheet, spec.Row, err)
		}
	}
	return nil
}

func applyMerges(workbook *excelize.File, merges []mergeSpec) error {
	for _, merge := range merges {
		if merge.Sheet == "" || merge.Range == "" {
			return errors.New("merge entries require sheet and range")
		}
		start, end, ok := splitRange(merge.Range)
		if !ok {
			return fmt.Errorf("invalid merge range %q", merge.Range)
		}
		if err := workbook.MergeCell(merge.Sheet, start, end); err != nil {
			return fmt.Errorf("merge cells %s!%s: %w", merge.Sheet, merge.Range, err)
		}
	}
	return nil
}

func applyDataValidations(workbook *excelize.File, validations []validationSpec) error {
	for _, spec := range validations {
		if spec.Sheet == "" || spec.Range == "" || spec.ValidationType == "" {
			return errors.New("validation entries require sheet, range, and validation_type")
		}
		dv := excelize.NewDataValidation(boolPtrDefault(spec.AllowBlank, true))
		dv.Sqref = spec.Range
		switch spec.ValidationType {
		case "list":
			if strings.HasPrefix(spec.Formula1, "\"") && strings.HasSuffix(spec.Formula1, "\"") {
				body := strings.Trim(spec.Formula1, "\"")
				if err := dv.SetDropList(strings.Split(body, ",")); err != nil {
					return fmt.Errorf("set validation drop list %s!%s: %w", spec.Sheet, spec.Range, err)
				}
			} else {
				dv.Type = "list"
				dv.Formula1 = spec.Formula1
			}
		case "custom":
			dv.Type = "custom"
			dv.Formula1 = spec.Formula1
		case "whole":
			if err := dv.SetRange(spec.Formula1, spec.Formula2, excelize.DataValidationTypeWhole, validationOperator(spec.Operator)); err != nil {
				return fmt.Errorf("set validation range %s!%s: %w", spec.Sheet, spec.Range, err)
			}
		default:
			return fmt.Errorf("unsupported validation type %q", spec.ValidationType)
		}
		if spec.ErrorTitle != "" || spec.Error != "" {
			dv.SetError(excelize.DataValidationErrorStyleStop, spec.ErrorTitle, spec.Error)
		}
		if err := workbook.AddDataValidation(spec.Sheet, dv); err != nil {
			return fmt.Errorf("add data validation %s!%s: %w", spec.Sheet, spec.Range, err)
		}
	}
	return nil
}

func applyHyperlinks(workbook *excelize.File, links []hyperlinkSpec) error {
	for _, spec := range links {
		if spec.Sheet == "" || spec.Cell == "" || spec.Target == "" {
			return errors.New("hyperlink entries require sheet, cell, and target")
		}
		display := spec.Display
		if display != "" {
			if err := workbook.SetCellValue(spec.Sheet, spec.Cell, display); err != nil {
				return fmt.Errorf("set hyperlink display %s!%s: %w", spec.Sheet, spec.Cell, err)
			}
		}
		tooltip := spec.Tooltip
		linkType := "External"
		if spec.Internal {
			linkType = "Location"
		}
		if err := workbook.SetCellHyperLink(spec.Sheet, spec.Cell, spec.Target, linkType, excelize.HyperlinkOpts{Display: stringPtrIfSet(display), Tooltip: stringPtrIfSet(tooltip)}); err != nil {
			return fmt.Errorf("set hyperlink %s!%s: %w", spec.Sheet, spec.Cell, err)
		}
	}
	return nil
}

func applyComments(workbook *excelize.File, comments []commentSpec) error {
	for _, spec := range comments {
		if spec.Sheet == "" || spec.Cell == "" {
			return errors.New("comment entries require sheet and cell")
		}
		comment := excelize.Comment{Cell: spec.Cell, Author: spec.Author, Paragraph: []excelize.RichTextRun{{Text: spec.Text}}}
		if err := workbook.AddComment(spec.Sheet, comment); err != nil {
			return fmt.Errorf("add comment %s!%s: %w", spec.Sheet, spec.Cell, err)
		}
	}
	return nil
}

func applyPanes(workbook *excelize.File, panes []paneSpec) error {
	for _, spec := range panes {
		if spec.Sheet == "" {
			return errors.New("pane entries require sheet")
		}
		mode := spec.Mode
		if mode == "" {
			mode = "freeze"
		}
		panesSpec := &excelize.Panes{Freeze: mode == "freeze", Split: mode == "split", XSplit: spec.XSplit, YSplit: spec.YSplit, TopLeftCell: spec.TopLeftCell, ActivePane: activePane(spec.XSplit, spec.YSplit, mode)}
		if err := workbook.SetPanes(spec.Sheet, panesSpec); err != nil {
			return fmt.Errorf("set panes %s: %w", spec.Sheet, err)
		}
	}
	return nil
}

func applyNamedRanges(workbook *excelize.File, names []namedRangeSpec) error {
	for _, spec := range names {
		if spec.Name == "" || spec.RefersTo == "" {
			return errors.New("named range entries require name and refers_to")
		}
		defined := &excelize.DefinedName{Name: spec.Name, RefersTo: spec.RefersTo}
		if spec.Scope == "sheet" {
			defined.Scope = spec.Sheet
		}
		if err := workbook.SetDefinedName(defined); err != nil {
			return fmt.Errorf("set defined name %s: %w", spec.Name, err)
		}
	}
	return nil
}

func applyTables(workbook *excelize.File, tables []tableSpec) error {
	for _, table := range tables {
		if table.Sheet == "" || table.Range == "" || table.Name == "" {
			return errors.New("table entries require sheet, range, and name")
		}
		style := table.Style
		if style == "" {
			style = "TableStyleMedium9"
		}
		if err := workbook.AddTable(table.Sheet, &excelize.Table{
			Range:          table.Range,
			Name:           table.Name,
			StyleName:      style,
			ShowHeaderRow:  table.ShowHeaderRow,
			ShowRowStripes: table.ShowRowStripes,
		}); err != nil {
			return fmt.Errorf("add table %q: %w", table.Name, err)
		}
	}
	return nil
}

func applyConditionalFormats(workbook *excelize.File, formats []conditionalFormatSpec) error {
	for _, format := range formats {
		if format.Sheet == "" || format.Range == "" || format.Type == "" {
			return errors.New("conditional format entries require sheet, range, and type")
		}
		option := excelize.ConditionalFormatOptions{
			Type:           format.Type,
			Criteria:       defaultString(format.Criteria, "="),
			Value:          format.Value,
			MinType:        defaultString(format.MinType, "min"),
			MidType:        format.MidType,
			MaxType:        defaultString(format.MaxType, "max"),
			MinColor:       defaultString(format.MinColor, "#F8696B"),
			MidColor:       format.MidColor,
			MaxColor:       defaultString(format.MaxColor, "#63BE7B"),
			BarColor:       defaultString(format.BarColor, "#638EC6"),
			BarBorderColor: format.BarBorderColor,
			IconStyle:      format.IconStyle,
			StopIfTrue:     format.StopIfTrue,
		}
		if option.Type == "formula" {
			option.Criteria = format.Criteria
		}
		if option.Type == "3_color_scale" && option.MidType == "" {
			option.MidType = "percentile"
			option.MidColor = defaultString(option.MidColor, "#FFEB84")
		}
		if option.Type == "icon_set" && option.IconStyle == "" {
			option.IconStyle = "3TrafficLights1"
		}
		if option.Type == "cell" && option.Criteria == "" {
			option.Criteria = ">"
		}
		if option.Type == "cell" && option.Value == "" {
			option.Value = "0"
		}
		if option.Type == "cell" || option.Type == "formula" {
			style, err := workbook.NewConditionalStyle(&excelize.Style{
				Font: &excelize.Font{Color: "9A0511"},
				Fill: excelize.Fill{Type: "pattern", Color: []string{"FEC7CE"}, Pattern: 1},
			})
			if err != nil {
				return fmt.Errorf("new conditional style: %w", err)
			}
			if format.BGColor != "" {
				style, err = workbook.NewConditionalStyle(&excelize.Style{
					Fill: excelize.Fill{Type: "pattern", Color: []string{format.BGColor}, Pattern: 1},
				})
				if err != nil {
					return fmt.Errorf("new conditional fill style: %w", err)
				}
			}
			option.Format = &style
		}
		if err := workbook.SetConditionalFormat(format.Sheet, format.Range, []excelize.ConditionalFormatOptions{option}); err != nil {
			return fmt.Errorf("set conditional format %s!%s: %w", format.Sheet, format.Range, err)
		}
	}
	return nil
}

func applyCharts(workbook *excelize.File, charts []chartSpec) error {
	for _, chart := range charts {
		if chart.Sheet == "" || chart.Cell == "" || chart.Type == "" {
			return errors.New("chart entries require sheet, cell, and type")
		}
		series := make([]excelize.ChartSeries, 0)
		if len(chart.Series) == 0 && chart.Categories != "" && chart.Values != "" {
			chart.Series = []seriesSpec{{Name: chart.Name, Categories: chart.Categories, Values: chart.Values, DataPoints: chart.DataPoints}}
		}
		for _, spec := range chart.Series {
			s := excelize.ChartSeries{
				Name:       spec.Name,
				Categories: spec.Categories,
				Values:     spec.Values,
			}
			if spec.FillColor != "" {
				s.Fill = excelize.Fill{Color: []string{spec.FillColor}}
			}
			for _, point := range spec.DataPoints {
				if point.FillColor != "" {
					s.DataPoint = append(s.DataPoint, excelize.ChartDataPoint{
						Index: point.Index,
						Fill:  excelize.Fill{Color: []string{point.FillColor}},
					})
				}
			}
			series = append(series, s)
		}
		chartType, err := chartTypeFromString(chart.Type)
		if err != nil {
			return err
		}
		width, height := chart.Width, chart.Height
		if width == 0 {
			width = 640
		}
		if height == 0 {
			height = 360
		}
		excelChart := &excelize.Chart{
			Type:   chartType,
			Series: series,
			Format: excelize.GraphicOptions{
				Name:    chart.Name,
				AltText: chart.AltText,
			},
			Dimension:  excelize.ChartDimension{Width: width, Height: height},
			Legend:     excelize.ChartLegend{Position: "right"},
			PlotArea:   excelize.ChartPlotArea{ShowVal: chart.ShowValues},
			VaryColors: chart.VaryColors,
		}
		if chart.Title != "" {
			excelChart.Title = []excelize.RichTextRun{{Text: chart.Title}}
		}
		if err := workbook.AddChart(chart.Sheet, chart.Cell, excelChart); err != nil {
			return fmt.Errorf("add chart %s!%s: %w", chart.Sheet, chart.Cell, err)
		}
	}
	return nil
}

func applyPivots(workbook *excelize.File, pivots []pivotSpec) error {
	for _, pivot := range pivots {
		if pivot.DataRange == "" || pivot.Range == "" || pivot.Name == "" {
			return errors.New("pivot entries require data_range, range, and name")
		}
		options := &excelize.PivotTableOptions{
			DataRange:           pivot.DataRange,
			PivotTableRange:     pivot.Range,
			Name:                pivot.Name,
			Rows:                pivotFields(pivot.Rows),
			Columns:             pivotFields(pivot.Columns),
			Data:                pivotFields(pivot.Data),
			Filter:              pivotFields(pivot.Filters),
			PivotTableStyleName: defaultString(pivot.Style, "PivotStyleMedium9"),
			ShowRowStripes:      pivot.ShowRowStripes,
			ShowColStripes:      pivot.ShowColStripes,
			RowGrandTotals:      boolDefault(pivot.RowGrandTotals, true),
			ColGrandTotals:      boolDefault(pivot.ColGrandTotals, true),
			ShowDrill:           true,
			UseAutoFormatting:   true,
			ShowRowHeaders:      true,
			ShowColHeaders:      true,
		}
		if err := workbook.AddPivotTable(options); err != nil {
			return fmt.Errorf("add pivot %q: %w", pivot.Name, err)
		}
	}
	return nil
}

func applySlicers(workbook *excelize.File, slicers []slicerSpec) error {
	for _, slicer := range slicers {
		if slicer.Sheet == "" || slicer.Name == "" || slicer.Cell == "" || slicer.TableName == "" {
			return errors.New("slicer entries require sheet, name, cell, and table_name")
		}
		width, height := slicer.Width, slicer.Height
		if width == 0 {
			width = 144
		}
		if height == 0 {
			height = 180
		}
		options := &excelize.SlicerOptions{
			Name:          slicer.Name,
			Cell:          slicer.Cell,
			TableSheet:    defaultString(slicer.TableSheet, slicer.Sheet),
			TableName:     slicer.TableName,
			Caption:       defaultString(slicer.Caption, slicer.Name),
			Width:         width,
			Height:        height,
			DisplayHeader: slicer.DisplayHeader,
		}
		if err := workbook.AddSlicer(slicer.Sheet, options); err != nil {
			return fmt.Errorf("add slicer %q: %w", slicer.Name, err)
		}
	}
	return nil
}

func applyPictures(workbook *excelize.File, pictures []pictureSpec) error {
	for _, picture := range pictures {
		if picture.Sheet == "" || picture.Cell == "" {
			return errors.New("picture entries require sheet and cell")
		}
		extension := defaultString(picture.Extension, ".png")
		encoded := picture.Base64
		if encoded == "" {
			encoded = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGP4z8AAAAMBAQDJ/pLvAAAAAElFTkSuQmCC"
		}
		bytes, err := base64.StdEncoding.DecodeString(encoded)
		if err != nil {
			return fmt.Errorf("decode picture base64: %w", err)
		}
		format := &excelize.GraphicOptions{
			Name:        picture.Name,
			AltText:     picture.AltText,
			ScaleX:      nonzeroFloat(picture.ScaleX, 1),
			ScaleY:      nonzeroFloat(picture.ScaleY, 1),
			Positioning: picture.Positioning,
		}
		if err := workbook.AddPictureFromBytes(picture.Sheet, picture.Cell, &excelize.Picture{
			Extension: extension,
			File:      bytes,
			Format:    format,
		}); err != nil {
			return fmt.Errorf("add picture %s!%s: %w", picture.Sheet, picture.Cell, err)
		}
	}
	return nil
}

func pivotFields(fields []pivotField) []excelize.PivotTableField {
	out := make([]excelize.PivotTableField, 0, len(fields))
	for _, field := range fields {
		out = append(out, excelize.PivotTableField{
			Name:     field.Name,
			Data:     field.Data,
			Subtotal: defaultString(field.Subtotal, "Sum"),
			NumFmt:   field.NumFmt,
		})
	}
	return out
}

func chartTypeFromString(value string) (excelize.ChartType, error) {
	switch value {
	case "area":
		return excelize.Area, nil
	case "bar":
		return excelize.Bar, nil
	case "col", "column":
		return excelize.Col, nil
	case "line":
		return excelize.Line, nil
	case "pie":
		return excelize.Pie, nil
	case "doughnut":
		return excelize.Doughnut, nil
	case "scatter":
		return excelize.Scatter, nil
	case "bubble":
		return excelize.Bubble, nil
	default:
		return 0, fmt.Errorf("unsupported chart type %q", value)
	}
}

func normalizeValue(value interface{}) interface{} {
	switch v := value.(type) {
	case json.Number:
		if i, err := strconv.ParseInt(v.String(), 10, 64); err == nil {
			return i
		}
		if f, err := strconv.ParseFloat(v.String(), 64); err == nil {
			return f
		}
		return v.String()
	default:
		return value
	}
}

func stringify(value interface{}) string {
	switch v := normalizeValue(value).(type) {
	case string:
		return v
	case fmt.Stringer:
		return v.String()
	default:
		return fmt.Sprint(v)
	}
}

func trimFormulaEquals(value string) string {
	return strings.TrimPrefix(value, "=")
}

func trimHash(value string) string {
	return strings.TrimPrefix(value, "#")
}

func splitRange(ref string) (string, string, bool) {
	parts := strings.Split(ref, ":")
	if len(parts) != 2 {
		return "", "", false
	}
	return parts[0], parts[1], true
}

func boolPtrDefault(value *bool, fallback bool) bool {
	if value == nil {
		return fallback
	}
	return *value
}

func stringPtrIfSet(value string) *string {
	if value == "" {
		return nil
	}
	return &value
}

func activePane(xSplit int, ySplit int, mode string) string {
	if mode == "split" {
		if xSplit > 0 && ySplit > 0 {
			return "bottomRight"
		}
		if xSplit > 0 {
			return "topRight"
		}
		if ySplit > 0 {
			return "bottomLeft"
		}
		return "topLeft"
	}
	if xSplit > 0 && ySplit == 0 {
		return "topRight"
	}
	if ySplit > 0 {
		return "bottomLeft"
	}
	return "topLeft"
}

func builtinNumFmt(value string) (int, bool) {
	switch value {
	case "0.00%":
		return 10, true
	case "0.00E+00":
		return 11, true
	default:
		return 0, false
	}
}

func buildBorders(border map[string]interface{}) []excelize.Border {
	if border == nil {
		return nil
	}
	order := []struct {
		key      string
		typeName string
	}{
		{"top", "top"},
		{"bottom", "bottom"},
		{"left", "left"},
		{"right", "right"},
		{"diagonal_up", "diagonalUp"},
		{"diagonal_down", "diagonalDown"},
	}
	result := make([]excelize.Border, 0)
	for _, item := range order {
		raw, ok := border[item.key].(map[string]interface{})
		if !ok || raw == nil {
			continue
		}
		styleName, _ := raw["style"].(string)
		if styleName == "" || styleName == "none" {
			continue
		}
		color, _ := raw["color"].(string)
		result = append(result, excelize.Border{Type: item.typeName, Color: trimHash(color), Style: borderStyle(styleName)})
	}
	return result
}

func borderStyle(value string) int {
	switch value {
	case "thin":
		return 1
	case "medium":
		return 2
	case "dashed":
		return 3
	case "dotted":
		return 4
	case "thick":
		return 5
	case "double":
		return 6
	case "hair":
		return 7
	case "mediumDashed":
		return 8
	case "dashDot":
		return 9
	case "mediumDashDot":
		return 10
	case "dashDotDot":
		return 11
	case "mediumDashDotDot":
		return 12
	case "slantDashDot":
		return 13
	default:
		return 1
	}
}

func normalizeExcelizeTables(path string, tables []tableSpec) error {
	if len(tables) == 0 {
		return nil
	}
	reader, err := zip.OpenReader(path)
	if err != nil {
		return fmt.Errorf("open generated workbook for table normalization: %w", err)
	}
	defer reader.Close()

	var buf bytes.Buffer
	writer := zip.NewWriter(&buf)
	for _, file := range reader.File {
		rc, err := file.Open()
		if err != nil {
			return fmt.Errorf("open zip member %s: %w", file.Name, err)
		}
		data, err := io.ReadAll(rc)
		_ = rc.Close()
		if err != nil {
			return fmt.Errorf("read zip member %s: %w", file.Name, err)
		}
		if strings.HasPrefix(file.Name, "xl/tables/table") && strings.HasSuffix(file.Name, ".xml") {
			idxText := strings.TrimSuffix(strings.TrimPrefix(file.Name, "xl/tables/table"), ".xml")
			if idx, convErr := strconv.Atoi(idxText); convErr == nil && idx >= 1 && idx <= len(tables) {
				data = []byte(patchTableXML(string(data), tables[idx-1]))
			}
		}
		hdr := file.FileHeader
		w, err := writer.CreateHeader(&hdr)
		if err != nil {
			return fmt.Errorf("create zip member %s: %w", file.Name, err)
		}
		if _, err := w.Write(data); err != nil {
			return fmt.Errorf("write zip member %s: %w", file.Name, err)
		}
	}
	if err := writer.Close(); err != nil {
		return fmt.Errorf("close normalized workbook zip: %w", err)
	}
	if err := os.WriteFile(path, buf.Bytes(), 0o644); err != nil {
		return fmt.Errorf("rewrite normalized workbook: %w", err)
	}
	return nil
}

func normalizeExcelizeConditionalFormats(path string, formats []conditionalFormatSpec) error {
	if len(formats) == 0 {
		return nil
	}
	reader, err := zip.OpenReader(path)
	if err != nil {
		return fmt.Errorf("open generated workbook for CF normalization: %w", err)
	}
	defer reader.Close()

	var buf bytes.Buffer
	writer := zip.NewWriter(&buf)
	styleColors := make([]string, 0)
	for _, spec := range formats {
		if spec.BGColor != "" {
			styleColors = append(styleColors, trimHash(spec.BGColor))
		}
	}
	for _, file := range reader.File {
		rc, err := file.Open()
		if err != nil {
			return fmt.Errorf("open zip member %s: %w", file.Name, err)
		}
		data, err := io.ReadAll(rc)
		_ = rc.Close()
		if err != nil {
			return fmt.Errorf("read zip member %s: %w", file.Name, err)
		}
		if file.Name == "xl/styles.xml" && len(styleColors) > 0 {
			data = []byte(patchStylesDxfColors(string(data), styleColors))
		}
		hdr := file.FileHeader
		w, err := writer.CreateHeader(&hdr)
		if err != nil {
			return fmt.Errorf("create zip member %s: %w", file.Name, err)
		}
		if _, err := w.Write(data); err != nil {
			return fmt.Errorf("write zip member %s: %w", file.Name, err)
		}
	}
	if err := writer.Close(); err != nil {
		return fmt.Errorf("close normalized workbook zip: %w", err)
	}
	if err := os.WriteFile(path, buf.Bytes(), 0o644); err != nil {
		return fmt.Errorf("rewrite normalized workbook: %w", err)
	}
	return nil
}

func patchTableXML(xml string, spec tableSpec) string {
	if spec.Range != "" {
		xml = replaceXMLAttr(xml, "ref", spec.Range)
		xml = replaceXMLAttr(xml, "xr:uid", "")
	}
	if spec.ShowHeaderRow != nil {
		xml = replaceXMLBoolAttr(xml, "headerRowCount", boolToIntString(*spec.ShowHeaderRow, 1, 0))
	}
	if spec.Style == "" {
		xml = removeTableStyleInfo(xml)
	}
	if spec.TotalsRow {
		xml = ensureXMLAttr(xml, "totalsRowCount", "1")
		xml = ensureXMLAttr(xml, "totalsRowShown", "1")
	} else {
		xml = ensureXMLAttr(xml, "totalsRowCount", "0")
		xml = ensureXMLAttr(xml, "totalsRowShown", "0")
	}
	xml = ensureAutoFilterRef(xml, spec.Range)
	return xml
}

func patchStylesDxfColors(xml string, colors []string) string {
	re := regexp.MustCompile(`<patternFill patternType="solid"><bgColor rgb="([A-Fa-f0-9]{6,8})"></bgColor>`)
	xml = re.ReplaceAllStringFunc(xml, func(match string) string {
		inner := re.FindStringSubmatch(match)
		if len(inner) < 2 {
			return match
		}
		argb := inner[1]
		if len(argb) == 6 {
			argb = "FF" + argb
		}
		return `<patternFill patternType="solid"><fgColor rgb="` + argb + `"/><bgColor rgb="` + argb + `"></bgColor>`
	})
	for _, color := range colors {
		argb := color
		if len(argb) == 6 {
			argb = "FF" + argb
		}
		if strings.Contains(xml, "<bgColor rgb=\"00000000\"") {
			xml = strings.Replace(xml, "<bgColor rgb=\"00000000\"", "<bgColor rgb=\""+argb+"\"", 1)
		}
	}
	return xml
}

func replaceXMLAttr(xml string, attr string, value string) string {
	prefix := attr + "=\""
	if value == "" {
		return xml
	}
	if start := strings.Index(xml, prefix); start >= 0 {
		start += len(prefix)
		end := strings.Index(xml[start:], "\"")
		if end >= 0 {
			return xml[:start] + value + xml[start+end:]
		}
	}
	return xml
}

func ensureXMLAttr(xml string, attr string, value string) string {
	prefix := attr + "=\""
	if start := strings.Index(xml, prefix); start >= 0 {
		start += len(prefix)
		end := strings.Index(xml[start:], "\"")
		if end >= 0 {
			return xml[:start] + value + xml[start+end:]
		}
	}
	tableStart := strings.Index(xml, "<table ")
	if tableStart >= 0 {
		if idx := strings.Index(xml[tableStart:], ">"); idx >= 0 {
			idx += tableStart
			return xml[:idx] + " " + attr + "=\"" + value + "\"" + xml[idx:]
		}
	}
	return xml
}

func replaceXMLBoolAttr(xml string, attr string, value string) string {
	return ensureXMLAttr(xml, attr, value)
}

func removeTableStyleInfo(xml string) string {
	start := strings.Index(xml, "<tableStyleInfo")
	if start < 0 {
		return xml
	}
	if end := strings.Index(xml[start:], "/>"); end >= 0 {
		end += start + 2
		return xml[:start] + xml[end:]
	}
	if end := strings.Index(xml[start:], "</tableStyleInfo>"); end >= 0 {
		end += start + len("</tableStyleInfo>")
		return xml[:start] + xml[end:]
	}
	return xml
}

func ensureAutoFilterRef(xml string, ref string) string {
	if ref == "" {
		return xml
	}
	prefix := "<autoFilter ref=\""
	if start := strings.Index(xml, prefix); start >= 0 {
		start += len(prefix)
		if end := strings.Index(xml[start:], "\""); end >= 0 {
			return xml[:start] + ref + xml[start+end:]
		}
	}
	return xml
}

func boolToIntString(value bool, trueValue int, falseValue int) string {
	if value {
		return strconv.Itoa(trueValue)
	}
	return strconv.Itoa(falseValue)
}

func validationOperator(value string) excelize.DataValidationOperator {
	switch value {
	case "between":
		return excelize.DataValidationOperatorBetween
	case "greaterThan":
		return excelize.DataValidationOperatorGreaterThan
	case "greaterThanOrEqual":
		return excelize.DataValidationOperatorGreaterThanOrEqual
	case "lessThan":
		return excelize.DataValidationOperatorLessThan
	case "lessThanOrEqual":
		return excelize.DataValidationOperatorLessThanOrEqual
	case "equal":
		return excelize.DataValidationOperatorEqual
	case "notBetween":
		return excelize.DataValidationOperatorNotBetween
	case "notEqual":
		return excelize.DataValidationOperatorNotEqual
	default:
		return excelize.DataValidationOperatorBetween
	}
}

func defaultString(value string, fallback string) string {
	if value == "" {
		return fallback
	}
	return value
}

func boolDefault(value *bool, fallback bool) bool {
	if value == nil {
		return fallback
	}
	return *value
}

func nonzeroFloat(value float64, fallback float64) float64 {
	if value == 0 {
		return fallback
	}
	return value
}

func parseTemporalCell(cellType string, value interface{}) (time.Time, error) {
	text := stringify(value)
	if cellType == "date" {
		return time.Parse("2006-01-02", text)
	}
	if parsed, err := time.Parse(time.RFC3339, text); err == nil {
		return parsed, nil
	}
	return time.Parse("2006-01-02T15:04:05", text)
}
