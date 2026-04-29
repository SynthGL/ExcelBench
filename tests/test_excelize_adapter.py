"""Tests for the ExcelizeAdapter cross-language write path."""

from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile

import pytest

from excelbench.harness.adapters.excelize_adapter import ExcelizeAdapter
from excelbench.harness.adapters.openpyxl_adapter import OpenpyxlAdapter
from excelbench.models import BorderEdge, BorderInfo, BorderStyle, CellFormat, CellType, CellValue


@pytest.fixture
def opxl() -> OpenpyxlAdapter:
    return OpenpyxlAdapter()


@pytest.fixture
def excelize() -> ExcelizeAdapter:
    return ExcelizeAdapter()


class TestExcelizeAdapterInfo:
    def test_name(self, excelize: ExcelizeAdapter) -> None:
        assert excelize.info.name == "excelize"

    def test_capabilities(self, excelize: ExcelizeAdapter) -> None:
        assert excelize.info.capabilities == {"write"}
        assert excelize.can_write()
        assert not excelize.can_read()


class TestExcelizeAvailability:
    def test_save_raises_structured_error_when_helper_missing(
        self, excelize: ExcelizeAdapter, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        class MissingTool:
            name = "excelize"
            command = ("missing-excelize-helper",)

            def is_available(self) -> bool:
                return False

            def resolve_executable(self) -> None:
                return None

            def resolved_command(self) -> None:
                return None

        monkeypatch.setattr(ExcelizeAdapter, "tool", classmethod(lambda cls: MissingTool()))
        wb = excelize.create_workbook()
        excelize.add_sheet(wb, "S")
        excelize.write_cell_value(wb, "S", "A1", CellValue(type=CellType.STRING, value="x"))
        with pytest.raises(FileNotFoundError):
            excelize.save_workbook(wb, tmp_path / "missing.xlsx")


@pytest.mark.skipif(not ExcelizeAdapter.is_available(), reason="Excelize helper is not available")
class TestExcelizeWriteRoundtrip:
    def test_values_and_multiple_sheets(
        self, excelize: ExcelizeAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "excelize_multi.xlsx"
        wb = excelize.create_workbook()
        excelize.add_sheet(wb, "Data")
        excelize.add_sheet(wb, "More")
        excelize.write_cell_value(wb, "Data", "A1", CellValue(type=CellType.STRING, value="hello"))
        excelize.write_cell_value(wb, "Data", "B1", CellValue(type=CellType.NUMBER, value=42.5))
        excelize.write_cell_value(
            wb,
            "Data",
            "C1",
            CellValue(type=CellType.FORMULA, value="=B1*2", formula="=B1*2"),
        )
        excelize.write_cell_value(wb, "More", "A1", CellValue(type=CellType.BOOLEAN, value=True))
        excelize.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        assert opxl.read_cell_value(rb, "Data", "A1").value == "hello"
        assert opxl.read_cell_value(rb, "Data", "B1").value == 42.5
        assert opxl.read_cell_value(rb, "Data", "C1").type == CellType.FORMULA
        assert opxl.read_cell_value(rb, "Data", "C1").formula == "=B1*2"
        assert opxl.read_cell_value(rb, "More", "A1").value is True
        opxl.close_workbook(rb)

    def test_blank_cell_roundtrip(
        self, excelize: ExcelizeAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "excelize_blank.xlsx"
        wb = excelize.create_workbook()
        excelize.add_sheet(wb, "Data")
        excelize.write_cell_value(wb, "Data", "A1", CellValue(type=CellType.BLANK, value=None))
        excelize.write_cell_value(wb, "Data", "A2", CellValue(type=CellType.STRING, value=""))
        excelize.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        assert opxl.read_cell_value(rb, "Data", "A1").type == CellType.BLANK
        assert opxl.read_cell_value(rb, "Data", "A2").type == CellType.BLANK
        opxl.close_workbook(rb)

    def test_text_format_alignment_validation_hyperlink_comment_and_panes(
        self, excelize: ExcelizeAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "excelize_features.xlsx"
        wb = excelize.create_workbook()
        excelize.add_sheet(wb, "S")
        excelize.write_cell_value(wb, "S", "A1", CellValue(type=CellType.STRING, value="Styled"))
        excelize.write_cell_format(
            wb,
            "S",
            "A1",
            CellFormat(
                bold=True,
                italic=True,
                underline="double",
                strikethrough=True,
                font_name="Arial",
                font_size=16,
                font_color="#FF0000",
                bg_color="#FFFF00",
                h_align="center",
                v_align="bottom",
                wrap=True,
                rotation=45,
                indent=2,
            ),
        )
        excelize.write_cell_value(wb, "S", "B2", CellValue(type=CellType.STRING, value="Link"))
        excelize.add_hyperlink(
            wb,
            "S",
            {
                "hyperlink": {
                    "cell": "B2",
                    "target": "https://example.com/docs",
                    "display": "Example Docs",
                    "tooltip": "Go to docs",
                    "internal": False,
                }
            },
        )
        excelize.add_comment(
            wb, "S", {"comment": {"cell": "C3", "text": "note", "author": "ExcelBench"}}
        )
        excelize.add_data_validation(
            wb,
            "S",
            {
                "validation": {
                    "range": "D4",
                    "validation_type": "list",
                    "formula1": '"Open,Closed"',
                    "allow_blank": True,
                }
            },
        )
        excelize.set_freeze_panes(
            wb,
            "S",
            {"freeze": {"mode": "freeze", "x_split": 1, "y_split": 1, "top_left_cell": "B2"}},
        )
        excelize.merge_cells(wb, "S", "E1:F1")
        excelize.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        fmt = opxl.read_cell_format(rb, "S", "A1")
        assert fmt.bold is True
        assert fmt.italic is True
        assert fmt.underline == "double"
        assert fmt.strikethrough is True
        assert fmt.font_name == "Arial"
        assert float(fmt.font_size or 0) == 16.0
        assert fmt.font_color and fmt.font_color.upper() == "#FF0000"
        assert fmt.bg_color and fmt.bg_color.upper() == "#FFFF00"
        assert fmt.h_align == "center"
        assert fmt.v_align == "bottom"
        assert fmt.wrap is True
        assert fmt.rotation == 45
        assert fmt.indent == 2
        ws = rb["S"]
        assert ws["B2"].hyperlink is not None
        assert ws["C3"].comment is not None
        assert len(list(ws.data_validations.dataValidation)) == 1
        assert ws.freeze_panes == "B2"
        assert "E1:F1" in {str(rng) for rng in ws.merged_cells.ranges}
        opxl.close_workbook(rb)

    def test_border_row_height_and_named_range_roundtrip(
        self, excelize: ExcelizeAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "excelize_border_named.xlsx"
        wb = excelize.create_workbook()
        excelize.add_sheet(wb, "named_ranges")
        excelize.write_cell_value(
            wb, "named_ranges", "B2", CellValue(type=CellType.STRING, value="Bordered")
        )
        excelize.write_cell_border(
            wb,
            "named_ranges",
            "B2",
            BorderInfo(
                top=BorderEdge(style=BorderStyle.THIN, color="#FF0000"),
                bottom=BorderEdge(style=BorderStyle.DOUBLE, color="#00FF00"),
                left=BorderEdge(style=BorderStyle.MEDIUM, color="#0000FF"),
                right=BorderEdge(style=BorderStyle.DASHED, color="#FFFF00"),
            ),
        )
        excelize.set_row_height(wb, "named_ranges", 2, 30)
        excelize.set_column_width(wb, "named_ranges", "B", 20)
        excelize.add_named_range(
            wb,
            "named_ranges",
            {
                "named_range": {
                    "name": "SingleCell",
                    "scope": "workbook",
                    "refers_to": "named_ranges!$B$2",
                }
            },
        )
        excelize.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        border = opxl.read_cell_border(rb, "named_ranges", "B2")
        assert border.top and border.top.style.value == "thin"
        assert border.bottom and border.bottom.style.value == "double"
        ws = rb["named_ranges"]
        assert round(ws.row_dimensions[2].height or 0) == 30
        assert round(ws.column_dimensions["B"].width or 0) == 20
        names = opxl.read_named_ranges(rb, "named_ranges")
        assert any(item.get("name") == "SingleCell" for item in names)
        opxl.close_workbook(rb)

    def test_conditional_format_roundtrip(
        self, excelize: ExcelizeAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "excelize_cf.xlsx"
        wb = excelize.create_workbook()
        excelize.add_sheet(wb, "conditional_formatting")
        for row in range(2, 7):
            excelize.write_cell_value(
                wb,
                "conditional_formatting",
                f"B{row}",
                CellValue(type=CellType.NUMBER, value=row - 1),
            )
        excelize.add_conditional_format(
            wb,
            "conditional_formatting",
            {
                "cf_rule": {
                    "range": "B2:B6",
                    "rule_type": "cellIs",
                    "operator": "greaterThan",
                    "formula": "5",
                    "format": {"bg_color": "#FFFF00"},
                }
            },
        )
        excelize.add_conditional_format(
            wb,
            "conditional_formatting",
            {
                "cf_rule": {
                    "range": "B2:B6",
                    "rule_type": "expression",
                    "formula": '=ISNUMBER(SEARCH("foo",B2))',
                    "format": {"bg_color": "#FF00FF"},
                }
            },
        )
        excelize.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        rules = opxl.read_conditional_formats(rb, "conditional_formatting")
        assert any(
            rule.get("rule_type") == "cellIs" and rule.get("formula") == "5" for rule in rules
        )
        assert any(rule.get("rule_type") == "expression" for rule in rules)
        opxl.close_workbook(rb)

    def test_tables_conditional_format_picture_and_pivot_parts(
        self, excelize: ExcelizeAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "excelize_structural.xlsx"
        wb = excelize.create_workbook()
        excelize.add_sheet(wb, "Data")
        excelize.add_sheet(wb, "Pivot")
        excelize.write_cell_value(wb, "Data", "A1", CellValue(type=CellType.STRING, value="Region"))
        excelize.write_cell_value(wb, "Data", "B1", CellValue(type=CellType.STRING, value="Sales"))
        excelize.write_cell_value(wb, "Data", "A2", CellValue(type=CellType.STRING, value="West"))
        excelize.write_cell_value(wb, "Data", "B2", CellValue(type=CellType.NUMBER, value=120))
        excelize.write_cell_value(wb, "Data", "A3", CellValue(type=CellType.STRING, value="East"))
        excelize.write_cell_value(wb, "Data", "B3", CellValue(type=CellType.NUMBER, value=95))
        excelize.add_table(wb, "Data", {"range": "A1:B3", "name": "SalesTable"})
        excelize.add_conditional_format(wb, "Data", {"range": "B2:B3", "type": "3_color_scale"})
        excelize.add_image(wb, "Data", {"cell": "D2", "name": "Pixel"})
        excelize.add_pivot_table(
            wb,
            "Pivot",
            {
                "data_range": "Data!A1:B3",
                "range": "Pivot!A3:D8",
                "name": "SalesPivot",
                "rows": [{"name": "Region"}],
                "data": [{"name": "Sales", "subtotal": "Sum"}],
            },
        )
        excelize.save_workbook(wb, path)

        with ZipFile(path) as workbook:
            names = set(workbook.namelist())
        assert "xl/tables/table1.xml" in names
        assert "xl/pivotTables/pivotTable1.xml" in names
        assert any(name.startswith("xl/media/") for name in names)
        assert any(name.startswith("xl/drawings/") for name in names)
