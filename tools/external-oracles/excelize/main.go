package main

import (
	"encoding/base64"
	"encoding/json"
	"errors"
	"fmt"
	_ "image/png"
	"io"
	"os"
	"path/filepath"
	"strconv"

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
	Columns            []columnSpec            `json:"columns"`
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

type columnSpec struct {
	Sheet string  `json:"sheet"`
	Start string  `json:"start"`
	End   string  `json:"end"`
	Width float64 `json:"width"`
}

type tableSpec struct {
	Sheet          string `json:"sheet"`
	Range          string `json:"range"`
	Name           string `json:"name"`
	Style          string `json:"style"`
	ShowHeaderRow  *bool  `json:"show_header_row"`
	ShowRowStripes *bool  `json:"show_row_stripes"`
}

type conditionalFormatSpec struct {
	Sheet          string `json:"sheet"`
	Range          string `json:"range"`
	Type           string `json:"type"`
	Criteria       string `json:"criteria"`
	Value          string `json:"value"`
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
	if err := applyColumns(workbook, payload.Columns); err != nil {
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

	return json.NewEncoder(output).Encode(map[string]interface{}{
		"fixture_id":  request.FixtureID,
		"operation":   request.Operation,
		"output_path": request.OutputPath,
		"tool":        "excelize",
		"counts": map[string]int{
			"sheets":              len(workbook.GetSheetList()),
			"cells":               len(payload.Cells),
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
		if cell.Type == "formula" || cell.Formula != "" {
			formula := cell.Formula
			if formula == "" {
				formula = stringify(cell.Value)
			}
			if err := workbook.SetCellFormula(cell.Sheet, cell.Cell, formula); err != nil {
				return fmt.Errorf("set formula %s!%s: %w", cell.Sheet, cell.Cell, err)
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
		if option.Type == "cell" {
			style, err := workbook.NewConditionalStyle(&excelize.Style{
				Font: &excelize.Font{Color: "9A0511"},
				Fill: excelize.Fill{Type: "pattern", Color: []string{"FEC7CE"}, Pattern: 1},
			})
			if err != nil {
				return fmt.Errorf("new conditional style: %w", err)
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
