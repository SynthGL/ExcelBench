"""Tests for to_expected() methods on Tier 2 spec models."""

from excelbench.models import (
    CommentSpec,
    ConditionalFormatSpec,
    DataValidationSpec,
    FreezePaneSpec,
    HyperlinkSpec,
    ImageSpec,
    MergeSpec,
    PivotSpec,
)

# ─────────────────────────────────────────────────
# MergeSpec
# ─────────────────────────────────────────────────


def test_merge_spec_minimal() -> None:
    spec = MergeSpec(range="A1:B2")
    result = spec.to_expected()
    assert result == {"merged_range": "A1:B2"}


def test_merge_spec_full() -> None:
    spec = MergeSpec(
        range="A1:C3",
        top_left_value="Hello",
        non_top_left_nonempty=2,
        top_left_bg_color="#FF0000",
        non_top_left_bg_color="#00FF00",
    )
    result = spec.to_expected()
    assert result["merged_range"] == "A1:C3"
    assert result["top_left_value"] == "Hello"
    assert result["non_top_left_nonempty"] == 2
    assert result["top_left_bg_color"] == "#FF0000"
    assert result["non_top_left_bg_color"] == "#00FF00"


# ─────────────────────────────────────────────────
# ConditionalFormatSpec
# ─────────────────────────────────────────────────


def test_cf_spec_minimal() -> None:
    spec = ConditionalFormatSpec(range="B2:B6", rule_type="cellIs")
    result = spec.to_expected()
    assert result == {"cf_rule": {"range": "B2:B6", "rule_type": "cellIs"}}


def test_cf_spec_full() -> None:
    spec = ConditionalFormatSpec(
        range="B2:B6",
        rule_type="cellIs",
        operator="greaterThan",
        formula="100",
        priority=1,
        stop_if_true=True,
        format={"font_color": "#FF0000"},
    )
    rule = spec.to_expected()["cf_rule"]
    assert rule["operator"] == "greaterThan"
    assert rule["formula"] == "100"
    assert rule["priority"] == 1
    assert rule["stop_if_true"] is True
    assert rule["format"] == {"font_color": "#FF0000"}


# ─────────────────────────────────────────────────
# DataValidationSpec
# ─────────────────────────────────────────────────


def test_dv_spec_minimal() -> None:
    spec = DataValidationSpec(range="A1:A10", validation_type="list")
    result = spec.to_expected()
    assert result == {"validation": {"range": "A1:A10", "validation_type": "list"}}


def test_dv_spec_full() -> None:
    spec = DataValidationSpec(
        range="A1:A10",
        validation_type="whole",
        operator="between",
        formula1="1",
        formula2="100",
        allow_blank=True,
        show_input=True,
        show_error=True,
        prompt_title="Enter value",
        prompt="Must be 1-100",
        error_title="Invalid",
        error="Out of range",
    )
    v = spec.to_expected()["validation"]
    assert v["operator"] == "between"
    assert v["formula1"] == "1"
    assert v["formula2"] == "100"
    assert v["allow_blank"] is True
    assert v["show_input"] is True
    assert v["show_error"] is True
    assert v["prompt_title"] == "Enter value"
    assert v["prompt"] == "Must be 1-100"
    assert v["error_title"] == "Invalid"
    assert v["error"] == "Out of range"


# ─────────────────────────────────────────────────
# HyperlinkSpec
# ─────────────────────────────────────────────────


def test_hyperlink_spec_minimal() -> None:
    spec = HyperlinkSpec(cell="A1", target="https://example.com")
    result = spec.to_expected()
    assert result == {
        "hyperlink": {"cell": "A1", "target": "https://example.com"}
    }


def test_hyperlink_spec_full() -> None:
    spec = HyperlinkSpec(
        cell="A1",
        target="https://example.com",
        display="Example",
        tooltip="Click here",
        internal=False,
    )
    h = spec.to_expected()["hyperlink"]
    assert h["display"] == "Example"
    assert h["tooltip"] == "Click here"
    assert h["internal"] is False


# ─────────────────────────────────────────────────
# ImageSpec
# ─────────────────────────────────────────────────


def test_image_spec_minimal() -> None:
    spec = ImageSpec(cell="A1", path="/img.png")
    result = spec.to_expected()
    assert result == {"image": {"cell": "A1", "path": "/img.png"}}


def test_image_spec_with_offset_converts_tuple_to_list() -> None:
    spec = ImageSpec(cell="A1", path="/img.png", offset=(10, 20))
    img = spec.to_expected()["image"]
    assert img["offset"] == [10, 20]
    assert isinstance(img["offset"], list)


def test_image_spec_full() -> None:
    spec = ImageSpec(
        cell="B3",
        path="/img.png",
        anchor="twoCellAnchor",
        offset=(5, 10),
        alt_text="A photo",
    )
    img = spec.to_expected()["image"]
    assert img["anchor"] == "twoCellAnchor"
    assert img["alt_text"] == "A photo"


# ─────────────────────────────────────────────────
# PivotSpec
# ─────────────────────────────────────────────────


def test_pivot_spec_minimal() -> None:
    spec = PivotSpec(
        name="PivotTable1",
        source_range="A1:D10",
        target_cell="F1",
        row_fields=["Region"],
        column_fields=["Product"],
        data_fields=["Sales"],
    )
    p = spec.to_expected()["pivot"]
    assert p["name"] == "PivotTable1"
    assert p["row_fields"] == ["Region"]
    assert "filter_fields" not in p


def test_pivot_spec_with_filters() -> None:
    spec = PivotSpec(
        name="PT",
        source_range="A1:D10",
        target_cell="F1",
        row_fields=["A"],
        column_fields=["B"],
        data_fields=["C"],
        filter_fields=["D"],
    )
    p = spec.to_expected()["pivot"]
    assert p["filter_fields"] == ["D"]


# ─────────────────────────────────────────────────
# CommentSpec
# ─────────────────────────────────────────────────


def test_comment_spec_minimal() -> None:
    spec = CommentSpec(cell="A1", text="Hello")
    result = spec.to_expected()
    assert result == {"comment": {"cell": "A1", "text": "Hello"}}


def test_comment_spec_full() -> None:
    spec = CommentSpec(cell="A1", text="Note", author="User1", threaded=True)
    c = spec.to_expected()["comment"]
    assert c["author"] == "User1"
    assert c["threaded"] is True


# ─────────────────────────────────────────────────
# FreezePaneSpec
# ─────────────────────────────────────────────────


def test_freeze_pane_spec_minimal() -> None:
    spec = FreezePaneSpec(mode="freeze")
    result = spec.to_expected()
    assert result == {"freeze": {"mode": "freeze"}}


def test_freeze_pane_spec_full() -> None:
    spec = FreezePaneSpec(
        mode="freeze",
        top_left_cell="B2",
        x_split=1,
        y_split=1,
        active_pane="bottomRight",
    )
    f = spec.to_expected()["freeze"]
    assert f["top_left_cell"] == "B2"
    assert f["x_split"] == 1
    assert f["y_split"] == 1
    assert f["active_pane"] == "bottomRight"
