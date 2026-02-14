from __future__ import annotations

from pathlib import Path

import pytest


def test_excelbench_rust_smoke(tmp_path: Path) -> None:
    pytest.importorskip("excelbench_rust")
    from excelbench_rust import UmyaBook

    # Create a tiny PNG for image insertion.
    from PIL import Image

    img_path = tmp_path / "logo.png"
    Image.new("RGBA", (1, 1), (255, 0, 0, 255)).save(img_path)

    out_path = tmp_path / "smoke.xlsx"

    book = UmyaBook()
    book.add_sheet("Data")

    book.write_cell_value("Data", "A1", {"type": "string", "value": "Hello"})
    book.merge_cells("Data", "A1:B1")

    book.add_comment("Data", {"cell": "A1", "text": "Note", "author": "Me"})
    book.add_hyperlink(
        "Data",
        {"cell": "A2", "target": "https://example.com", "display": "Example"},
    )

    # Ergonomic alias input: row/column -> top_left_cell for freeze panes.
    book.set_freeze_panes("Data", {"row": 1, "column": 1})

    book.add_image("Data", {"path": str(img_path), "cell": "C3"})

    # Alias inputs: ranges/type/error_message.
    book.add_data_validation(
        "Data",
        {
            "ranges": ["B2:B5"],
            "type": "list",
            "formula1": "A,B,C",
            "allow_blank": True,
            "error_message": "Choose from the list",
            "show_error": True,
        },
    )
    book.add_conditional_format(
        "Data",
        {
            "ranges": ["D1:D5"],
            "type": "cellIs",
            "operator": "greaterThan",
            "formula": "1",
            "format": {"bg_color": "#C6EFCE"},
        },
    )

    book.save(str(out_path))

    reopened = UmyaBook.open(str(out_path))
    assert "Data" in reopened.sheet_names()

    panes = reopened.read_freeze_panes("Data")
    assert panes.get("mode") == "freeze"
    assert panes.get("top_left_cell") == "B2"

    comments = reopened.read_comments("Data")
    assert any(c.get("cell") == "A1" for c in comments)

    links = reopened.read_hyperlinks("Data")
    assert any(link.get("target") == "https://example.com" for link in links)

    dvs = reopened.read_data_validations("Data")
    assert any(d.get("validation_type") == "list" for d in dvs)

    cfs = reopened.read_conditional_formats("Data")
    assert any(cf.get("rule_type") == "cellIs" for cf in cfs)

    imgs = reopened.read_images("Data")
    assert any(i.get("cell") == "C3" for i in imgs if i.get("cell"))
