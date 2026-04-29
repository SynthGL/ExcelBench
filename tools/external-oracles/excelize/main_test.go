package main

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"path/filepath"
	"testing"
)

func TestWriteFixtureCreatesAdvancedWorkbookParts(t *testing.T) {
	t.Parallel()
	outputPath := filepath.Join(t.TempDir(), "excelize-smoke.xlsx")
	request := map[string]interface{}{
		"fixture_id":  "excelize-smoke",
		"operation":   "write_fixture",
		"output_path": outputPath,
		"payload": map[string]interface{}{
			"sheets": []map[string]interface{}{{"name": "Data"}, {"name": "Pivot"}},
			"cells": []map[string]interface{}{
				{"sheet": "Data", "cell": "A1", "value": "Region"},
				{"sheet": "Data", "cell": "B1", "value": "Product"},
				{"sheet": "Data", "cell": "C1", "value": "Sales"},
				{"sheet": "Data", "cell": "A2", "value": "West"},
				{"sheet": "Data", "cell": "B2", "value": "Widgets"},
				{"sheet": "Data", "cell": "C2", "value": 120},
				{"sheet": "Data", "cell": "A3", "value": "East"},
				{"sheet": "Data", "cell": "B3", "value": "Services"},
				{"sheet": "Data", "cell": "C3", "value": 95},
			},
			"tables": []map[string]interface{}{
				{"sheet": "Data", "range": "A1:C3", "name": "SalesTable"},
			},
			"conditional_formats": []map[string]interface{}{
				{"sheet": "Data", "range": "C2:C3", "type": "3_color_scale"},
				{"sheet": "Data", "range": "C2:C3", "type": "data_bar"},
			},
			"charts": []map[string]interface{}{
				{
					"sheet":      "Data",
					"cell":       "E2",
					"type":       "col",
					"title":      "Sales",
					"categories": "Data!$A$2:$A$3",
					"values":     "Data!$C$2:$C$3",
				},
			},
			"pivots": []map[string]interface{}{
				{
					"data_range": "Data!A1:C3",
					"range":      "Pivot!A3:E10",
					"name":       "SalesPivot",
					"rows":       []map[string]interface{}{{"name": "Region"}},
					"data":       []map[string]interface{}{{"name": "Sales", "subtotal": "Sum"}},
				},
			},
			"slicers": []map[string]interface{}{
				{
					"sheet":       "Data",
					"name":        "Region",
					"cell":        "E15",
					"table_sheet": "Data",
					"table_name":  "SalesTable",
				},
			},
			"pictures": []map[string]interface{}{
				{"sheet": "Data", "cell": "H2", "name": "Pixel"},
			},
		},
	}
	stdin, err := json.Marshal(request)
	if err != nil {
		t.Fatalf("marshal request: %v", err)
	}

	var stdout bytes.Buffer
	if err := run(bytes.NewReader(stdin), &stdout); err != nil {
		t.Fatalf("run helper: %v", err)
	}

	var response map[string]interface{}
	if err := json.Unmarshal(stdout.Bytes(), &response); err != nil {
		t.Fatalf("decode response %q: %v", stdout.String(), err)
	}
	if response["tool"] != "excelize" {
		t.Fatalf("unexpected tool response: %#v", response)
	}
	assertZipContains(t, outputPath, []string{
		"xl/tables/table1.xml",
		"xl/pivotTables/pivotTable1.xml",
		"xl/pivotCache/pivotCacheDefinition1.xml",
		"xl/slicers/slicer1.xml",
		"xl/slicerCaches/slicerCache1.xml",
		"xl/charts/chart1.xml",
		"xl/drawings/drawing1.xml",
	})
}

func assertZipContains(t *testing.T, path string, names []string) {
	t.Helper()
	reader, err := zip.OpenReader(path)
	if err != nil {
		t.Fatalf("open workbook zip: %v", err)
	}
	defer reader.Close()

	seen := make(map[string]bool)
	for _, file := range reader.File {
		seen[file.Name] = true
	}
	for _, name := range names {
		if !seen[name] {
			t.Fatalf("workbook missing %s; saw %v", name, seen)
		}
	}
}

