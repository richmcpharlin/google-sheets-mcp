import pytest

from google_sheets_mcp.formatting import (
    assert_bounded_range,
    collapse_runs,
    color_to_hex,
    runs_from_grid,
    style_from_cell,
)


def test_color_to_hex_rounds():
    assert color_to_hex({"red": 1, "green": 1, "blue": 0}) == "#ffff00"
    assert color_to_hex({"green": 0.5}) == "#008000"


def test_style_drops_defaults():
    style = style_from_cell(
        {
            "effectiveFormat": {
                "backgroundColor": {"red": 1, "green": 1, "blue": 1},
                "textFormat": {
                    "bold": False,
                    "fontSize": 10,
                    "foregroundColor": {"red": 0, "green": 0, "blue": 0},
                },
            }
        }
    )
    assert style == {}


def test_style_keeps_color_emphasis_and_number_format():
    style = style_from_cell(
        {
            "effectiveFormat": {
                "backgroundColorStyle": {"rgbColor": {"red": 1, "green": 1}},
                "textFormat": {"bold": True, "fontSize": 12, "foregroundColor": {"blue": 1}},
                "numberFormat": {"type": "CURRENCY", "pattern": "$#,##0.00"},
            },
            "note": "input",
        }
    )
    assert style == {
        "backgroundColor": "#ffff00",
        "textColor": "#0000ff",
        "bold": True,
        "fontSize": 12,
        "numberFormat": {"type": "CURRENCY", "pattern": "$#,##0.00"},
        "note": "input",
    }


def test_collapse_merges_a_rectangle_and_keeps_an_l():
    yellow = {"backgroundColor": "#ffff00", "bold": True}
    runs = collapse_runs(
        [
            (0, 0, yellow),
            (0, 1, yellow),
            (1, 0, yellow),
            (1, 1, yellow),
            (1, 2, {"numberFormat": {"type": "CURRENCY"}}),
        ],
        "Eng & Product",
    )
    assert runs[0]["range"] == "'Eng & Product'!A1:B2"
    assert runs[0]["backgroundColor"] == "#ffff00"
    assert runs[1]["range"] == "'Eng & Product'!C2"

    l_shape = collapse_runs([(0, 0, yellow), (0, 1, yellow), (1, 0, yellow)], "Inputs")
    assert [run["range"] for run in l_shape] == ["Inputs!A1:B1", "Inputs!A2"]


def test_runs_from_grid_skips_plain_cells():
    response = {
        "sheets": [
            {
                "properties": {"title": "Inputs"},
                "data": [
                    {
                        "rowData": [
                            {
                                "values": [
                                    {
                                        "effectiveFormat": {
                                            "backgroundColor": {"red": 1, "green": 1},
                                            "textFormat": {"bold": True},
                                        }
                                    },
                                    {
                                        "effectiveFormat": {
                                            "backgroundColor": {"red": 1, "green": 1, "blue": 1},
                                            "textFormat": {"fontSize": 10},
                                        }
                                    },
                                ]
                            }
                        ]
                    }
                ],
            }
        ]
    }
    runs = runs_from_grid(response)
    assert len(runs) == 1
    assert runs[0]["range"] == "Inputs!A1"
    assert runs[0]["bold"] is True


def test_open_range_rejected():
    with pytest.raises(ValueError):
        assert_bounded_range("Inputs!A:C")


def test_oversized_range_rejected():
    with pytest.raises(ValueError):
        assert_bounded_range("Inputs!A1:Z200")


def test_sheet_defaults_are_not_reported():
    plain = {
        "effectiveFormat": {
            "textFormat": {"foregroundColorStyle": {"themeColor": "TEXT"}, "fontSize": 11},
            "backgroundColorStyle": {"themeColor": "BACKGROUND"},
        }
    }
    bold = {"effectiveFormat": {"textFormat": {"bold": True, "fontSize": 11}}}
    response = {
        "properties": {"defaultFormat": {"textFormat": {"fontSize": 11}}},
        "sheets": [{"properties": {"title": "Budget"}, "data": [{"rowData": [{"values": [plain, bold]}]}]}],
    }
    runs = runs_from_grid(response)
    assert runs == [{"range": "Budget!B1", "bold": True}]
