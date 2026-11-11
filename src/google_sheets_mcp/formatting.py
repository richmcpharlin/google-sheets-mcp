"""Read displayed cell formatting and collapse it into rectangular runs.

The Sheets API returns one format object per cell. Callers want "the header
row is yellow and bold", so identical neighbors are merged into one A1 range.
"""

from __future__ import annotations

from .a1 import a1_to_grid_range, col_index_to_letter, split_sheet_and_range
from .refs import quote_sheet

# ponytail: one spreadsheets.get of the whole range, then an in-memory merge.
# Ceiling is MAX_CELLS. Upgrade path is paging the grid and emitting row-runs.
MAX_CELLS = 2500

DEFAULT_FONT_SIZE = 10

FORMAT_FIELDS = (
    "properties.defaultFormat.textFormat.fontSize,"
    "sheets(properties.title,data(startRow,startColumn,"
    "rowData(values(effectiveFormat(backgroundColor,backgroundColorStyle,"
    "textFormat(foregroundColor,foregroundColorStyle,bold,italic,underline,"
    "strikethrough,fontSize),numberFormat),note))))"
)


def color_to_hex(color: dict | None) -> str | None:
    """Turn a Sheets ``Color`` (channels 0–1) into ``#rrggbb``.

    Returns ``None`` when the color has no RGB channels (a bare theme color
    is handled by the caller).
    """
    if not color:
        return None
    if not any(k in color for k in ("red", "green", "blue")):
        return None

    def channel(name: str) -> int:
        value = float(color.get(name, 0))
        return max(0, min(255, round(value * 255)))

    return f"#{channel('red'):02x}{channel('green'):02x}{channel('blue'):02x}"


def _color_value(rgb: dict | None, style: dict | None) -> str | None:
    """Prefer a resolved hex color, then a theme-color name."""
    if style and style.get("rgbColor"):
        return color_to_hex(style["rgbColor"])
    if rgb:
        hex_color = color_to_hex(rgb)
        if hex_color:
            return hex_color
    if style and style.get("themeColor"):
        return style["themeColor"]
    return None


def style_from_cell(cell: dict, default_font_size: int = DEFAULT_FONT_SIZE) -> dict:
    """Pick the formatting worth reporting from one grid cell.

    White backgrounds, black text, theme default colors, and the spreadsheet's
    default font size are omitted. What remains is what a reader would notice:
    color, emphasis, a number format, or a cell note. Colors come from the
    effective format, so conditional formatting is included.
    """
    fmt = cell.get("effectiveFormat") or {}
    text = fmt.get("textFormat") or {}
    style: dict = {}

    background = _color_value(fmt.get("backgroundColor"), fmt.get("backgroundColorStyle"))
    if background and background.lower() not in ("#ffffff", "background"):
        style["backgroundColor"] = background

    foreground = _color_value(text.get("foregroundColor"), text.get("foregroundColorStyle"))
    if foreground and foreground.lower() not in ("#000000", "text"):
        style["textColor"] = foreground

    for flag in ("bold", "italic", "underline", "strikethrough"):
        if text.get(flag):
            style[flag] = True

    font_size = text.get("fontSize")
    if font_size is not None and font_size != default_font_size:
        style["fontSize"] = font_size

    number_format = fmt.get("numberFormat") or {}
    if number_format.get("type"):
        entry = {"type": number_format["type"]}
        if number_format.get("pattern"):
            entry["pattern"] = number_format["pattern"]
        style["numberFormat"] = entry

    if cell.get("note"):
        style["note"] = cell["note"]

    return style


def assert_bounded_range(range_a1: str, limit: int = MAX_CELLS) -> None:
    """Reject open-ended or oversized ranges before calling the API.

    Whole columns (``A:C``), whole rows (``2:5``), and a bare sheet name have
    no finite cell count. A bounded range over ``limit`` cells is also rejected.
    """
    _, cell_range = split_sheet_and_range(range_a1)
    grid = a1_to_grid_range(cell_range)
    end_row = grid.get("endRowIndex")
    end_col = grid.get("endColumnIndex")
    if end_row is None or end_col is None:
        raise ValueError(
            "Range must be bounded, like A1:C20. Whole rows, whole columns, "
            "and whole sheets are too large to read formatting for."
        )
    count = (end_row - grid.get("startRowIndex", 0)) * (end_col - grid.get("startColumnIndex", 0))
    if count > limit:
        raise ValueError(f"Range has {count} cells; the limit is {limit}. Narrow the range.")


def _span(row: int, col: int, end_row: int, end_col: int) -> str:
    """A1 for a rectangle. Indices are zero-based and inclusive."""
    start = f"{col_index_to_letter(col)}{row + 1}"
    if row == end_row and col == end_col:
        return start
    end = f"{col_index_to_letter(end_col)}{end_row + 1}"
    return f"{start}:{end}"


def collapse_runs(cells: list[tuple[int, int, dict]], sheet_title: str | None) -> list[dict]:
    """Merge same-style neighbors into rectangles, scanned row by row.

    ``cells`` is ``(row_index, col_index, style)``. An L shape stays two runs.
    """
    by_pos = {(row, col): style for row, col, style in cells}
    visited: set[tuple[int, int]] = set()
    prefix = f"{quote_sheet(sheet_title)}!" if sheet_title else ""
    runs: list[dict] = []

    for row, col in sorted(by_pos):
        if (row, col) in visited:
            continue
        style = by_pos[(row, col)]
        end_col = col
        while by_pos.get((row, end_col + 1)) == style and (row, end_col + 1) not in visited:
            end_col += 1
        end_row = row
        while all(
            by_pos.get((end_row + 1, cc)) == style and (end_row + 1, cc) not in visited
            for cc in range(col, end_col + 1)
        ):
            end_row += 1
        for rr in range(row, end_row + 1):
            for cc in range(col, end_col + 1):
                visited.add((rr, cc))
        runs.append({"range": prefix + _span(row, col, end_row, end_col), **style})

    return runs


def runs_from_grid(response: dict, include_defaults: bool = False) -> list[dict]:
    """Build collapsed runs from a ``spreadsheets.get`` grid-data response.

    The spreadsheet's own default font size is the baseline, so a sheet set to
    11pt does not report every cell as 11pt.
    """
    default_format = response.get("properties", {}).get("defaultFormat", {})
    default_font_size = default_format.get("textFormat", {}).get("fontSize", DEFAULT_FONT_SIZE)
    runs: list[dict] = []
    for sheet in response.get("sheets", []):
        title = sheet.get("properties", {}).get("title")
        cells: list[tuple[int, int, dict]] = []
        for block in sheet.get("data", []):
            start_row = block.get("startRow", 0)
            start_col = block.get("startColumn", 0)
            for row_offset, row in enumerate(block.get("rowData", [])):
                for col_offset, cell in enumerate(row.get("values", [])):
                    style = style_from_cell(cell, default_font_size)
                    if style or include_defaults:
                        cells.append((start_row + row_offset, start_col + col_offset, style))
        runs.extend(collapse_runs(cells, title))
    return runs
