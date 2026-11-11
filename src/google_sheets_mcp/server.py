"""FastMCP server exposing cell-level and formatting-level Google Sheets tools.

Run locally over stdio (the default transport) — e.g. configured as an MCP
server in Claude Desktop / Claude Code, or `google-sheets-mcp` from a shell.
"""

from __future__ import annotations

import re
from typing import Any

from googleapiclient.errors import HttpError
from mcp.server.fastmcp import FastMCP

from .a1 import a1_to_grid_range, split_sheet_and_range
from .client import SheetsClient
from .colors import parse_color
from .config import mode as _mode
from .formatting import FORMAT_FIELDS, assert_bounded_range, runs_from_grid
from .refs import resolve_spreadsheet_id

# ---------------------------------------------------------------------------
# Server construction (mode-aware) and per-request client resolution
# ---------------------------------------------------------------------------


def _build_mcp() -> FastMCP:
    """Build the FastMCP server for the active mode.

    local (default): plain stdio server using one shared on-disk identity.
    group: a Streamable-HTTP server that acts as an OAuth Authorization Server
    bridging each user to Google (see ``oauth_bridge.py``). Auth has to be
    configured at construction time, so we branch here at import time on
    ``$GSHEETS_MCP_MODE``.
    """
    if _mode() == "group":
        from mcp.server.auth.settings import (
            AuthSettings,
            ClientRegistrationOptions,
            RevocationOptions,
        )

        from .config import load_group_config
        from .oauth_bridge import GoogleBridgeProvider, register_callback

        cfg = load_group_config()
        provider = GoogleBridgeProvider(cfg)
        server = FastMCP(
            "google-sheets",
            host=cfg.host,
            port=cfg.port,
            auth_server_provider=provider,
            auth=AuthSettings(
                issuer_url=cfg.issuer_url,
                resource_server_url=cfg.issuer_url,
                client_registration_options=ClientRegistrationOptions(
                    enabled=True,
                    valid_scopes=cfg.scopes,
                    default_scopes=cfg.scopes,
                ),
                revocation_options=RevocationOptions(enabled=True),
                required_scopes=cfg.scopes,
            ),
        )
        register_callback(server, provider)
        return server

    return FastMCP("google-sheets")


mcp = _build_mcp()

_local_singleton: SheetsClient | None = None


def current_client() -> SheetsClient:
    """Return the ``SheetsClient`` for the current request.

    group: build a client from the caller's verified Google access token
    (carried in the MCP access token's claims). local: a shared singleton.
    """
    if _mode() == "group":
        from mcp.server.auth.middleware.auth_context import get_access_token

        from .auth import credentials_from_access_token

        token = get_access_token()
        if token is None or not token.claims or "g_at" not in token.claims:
            raise RuntimeError("No authenticated Google credentials on this request.")
        return SheetsClient(credentials_from_access_token(token.claims["g_at"]))

    global _local_singleton
    if _local_singleton is None:
        _local_singleton = SheetsClient()
    return _local_singleton

# ---------------------------------------------------------------------------
# Validation helpers
# ---------------------------------------------------------------------------

_SPREADSHEET_ID_RE = re.compile(r"^[A-Za-z0-9_-]{20,100}$")
_RENDER_OPTIONS = {"FORMATTED_VALUE", "UNFORMATTED_VALUE", "FORMULA"}
_INPUT_OPTIONS = {"USER_ENTERED", "RAW"}
_BASIC_CHART_TYPES = {"BAR", "LINE", "AREA", "COLUMN", "SCATTER", "COMBO", "STEPPED_AREA"}
_PIE_CHART_TYPES = {"PIE", "DONUT"}
_BASIC_LEGEND_POSITIONS = {
    "BOTTOM_LEGEND",
    "LEFT_LEGEND",
    "RIGHT_LEGEND",
    "TOP_LEGEND",
    "NO_LEGEND",
}
_PIE_LEGEND_POSITIONS = _BASIC_LEGEND_POSITIONS | {"LABELED_LEGEND"}
_STACKED_TYPES = {"NOT_STACKED", "STACKED", "PERCENT_STACKED"}
_ADVANCED_BATCH_MAX_REQUESTS = 50
_BLOCKED_BATCH_REQUEST_KINDS = {
    "addDataSource",
    "updateDataSource",
    "deleteDataSource",
    "refreshDataSource",
    "cancelDataSourceRefresh",
}
_ALLOWED_BATCH_REQUEST_KINDS = {
    "updateSpreadsheetProperties",
    "updateSheetProperties",
    "updateDimensionProperties",
    "updateNamedRange",
    "repeatCell",
    "addNamedRange",
    "deleteNamedRange",
    "addSheet",
    "deleteSheet",
    "autoFill",
    "cutPaste",
    "copyPaste",
    "mergeCells",
    "unmergeCells",
    "updateBorders",
    "updateCells",
    "addFilterView",
    "appendCells",
    "clearBasicFilter",
    "deleteDimension",
    "deleteEmbeddedObject",
    "deleteFilterView",
    "duplicateFilterView",
    "duplicateSheet",
    "findReplace",
    "insertDimension",
    "insertRange",
    "moveDimension",
    "updateEmbeddedObjectPosition",
    "pasteData",
    "textToColumns",
    "updateFilterView",
    "deleteRange",
    "appendDimension",
    "addConditionalFormatRule",
    "updateConditionalFormatRule",
    "deleteConditionalFormatRule",
    "sortRange",
    "setDataValidation",
    "setBasicFilter",
    "addProtectedRange",
    "updateProtectedRange",
    "deleteProtectedRange",
    "autoResizeDimensions",
    "addChart",
    "updateChartSpec",
    "updateBanding",
    "addBanding",
    "deleteBanding",
    "createDeveloperMetadata",
    "updateDeveloperMetadata",
    "deleteDeveloperMetadata",
    "randomizeRange",
    "addDimensionGroup",
    "deleteDimensionGroup",
    "updateDimensionGroup",
    "trimWhitespace",
    "deleteDuplicates",
    "updateEmbeddedObjectBorder",
    "addSlicer",
    "updateSlicerSpec",
    "addTable",
    "updateTable",
    "deleteTable",
}

# Functions that make outbound HTTP requests when evaluated by Sheets.
# Blocked inside CUSTOM_FORMULA conditions to prevent data exfiltration.
_NETWORK_FORMULA_RE = re.compile(
    r"\b(IMPORTDATA|IMPORTRANGE|IMPORTHTML|IMPORTFEED|IMPORTXML|IMAGE)\s*\(",
    re.IGNORECASE,
)


def _sid(spreadsheet_id: str) -> str:
    """Validate and return a spreadsheet ID, accepting a full Sheets URL too."""
    candidate = resolve_spreadsheet_id(spreadsheet_id)
    if not _SPREADSHEET_ID_RE.fullmatch(candidate):
        raise ValueError(
            "Invalid spreadsheet_id — expected a Google Sheets ID "
            "(20–100 chars of letters/digits/hyphens/underscores) or a full "
            "https://docs.google.com/spreadsheets/d/… URL."
        )
    return candidate


def _render_option(opt: str) -> str:
    upper = opt.upper()
    if upper not in _RENDER_OPTIONS:
        raise ValueError(
            f"Invalid value_render_option {opt!r}. "
            f"Choose from: {', '.join(sorted(_RENDER_OPTIONS))}."
        )
    return upper


def _input_option(opt: str) -> str:
    upper = opt.upper()
    if upper not in _INPUT_OPTIONS:
        raise ValueError(
            f"Invalid value_input_option {opt!r}. "
            f"Choose from: {', '.join(sorted(_INPUT_OPTIONS))}."
        )
    return upper


def _api(fn, *args, **kwargs):
    """Execute a Google API call, stripping URL details from errors."""
    try:
        return fn(*args, **kwargs).execute()
    except HttpError as exc:
        raise RuntimeError(
            f"Google Sheets API error {exc.status_code}: {exc.reason}"
        ) from None


def _request_kind(req: dict) -> str:
    if not isinstance(req, dict):
        raise ValueError("Each batchUpdate request must be an object.")
    kinds = [k for k, v in req.items() if v is not None]
    if len(kinds) != 1:
        raise ValueError(
            "Each batchUpdate request must specify exactly one request kind "
            f"(got {len(kinds)}: {', '.join(kinds) or 'none'})."
        )
    kind = kinds[0]
    if kind in _BLOCKED_BATCH_REQUEST_KINDS:
        raise ValueError(f"Request kind {kind!r} is blocked by this MCP server.")
    if kind not in _ALLOWED_BATCH_REQUEST_KINDS:
        allowed = ", ".join(sorted(_ALLOWED_BATCH_REQUEST_KINDS))
        raise ValueError(f"Request kind {kind!r} is not allowed. Allowed kinds: {allowed}.")
    return kind


def _validate_advanced_requests(requests: list[dict]) -> list[str]:
    if not requests:
        raise ValueError("Provide at least one request.")
    if len(requests) > _ADVANCED_BATCH_MAX_REQUESTS:
        raise ValueError(
            f"Too many requests: {len(requests)}. Maximum is {_ADVANCED_BATCH_MAX_REQUESTS}."
        )
    return [_request_kind(req) for req in requests]


def _execute_batch_update(
    spreadsheet_id: str,
    requests: list[dict],
    response_ranges: list[str] | None = None,
    include_spreadsheet_in_response: bool = False,
) -> dict:
    body: dict[str, Any] = {
        "requests": requests,
        "includeSpreadsheetInResponse": include_spreadsheet_in_response,
    }
    if response_ranges:
        body["responseRanges"] = response_ranges
    return _api(
        current_client().spreadsheets.batchUpdate,
        spreadsheetId=_sid(spreadsheet_id),
        body=body,
    )


def _dimension(dimension: str) -> str:
    upper = dimension.upper()
    if upper not in {"ROWS", "COLUMNS"}:
        raise ValueError("dimension must be ROWS or COLUMNS.")
    return upper


def _dimension_range(
    spreadsheet_id: str,
    sheet: str,
    dimension: str,
    start_index: int,
    end_index: int,
) -> dict:
    if start_index < 0 or end_index <= start_index:
        raise ValueError("Dimension indices must be zero-based and end_index must be greater than start_index.")
    return {
        "sheetId": current_client().resolve_sheet_id(_sid(spreadsheet_id), sheet),
        "dimension": _dimension(dimension),
        "startIndex": start_index,
        "endIndex": end_index,
    }


def _single_cell_coordinate(spreadsheet_id: str, cell: str, default_sheet: str | int | None = None) -> dict:
    grid = current_client().resolve_grid_range(_sid(spreadsheet_id), cell, default_sheet=default_sheet)
    if (
        "startRowIndex" not in grid
        or "startColumnIndex" not in grid
        or grid.get("endRowIndex") != grid["startRowIndex"] + 1
        or grid.get("endColumnIndex") != grid["startColumnIndex"] + 1
    ):
        raise ValueError("Expected a single cell reference like 'A1' or 'Sheet1!A1'.")
    return {
        "sheetId": grid["sheetId"],
        "rowIndex": grid["startRowIndex"],
        "columnIndex": grid["startColumnIndex"],
    }


def _overlay_position(
    spreadsheet_id: str,
    sheet: str,
    anchor_cell: str,
    width_pixels: int = 600,
    height_pixels: int = 371,
    offset_x_pixels: int = 0,
    offset_y_pixels: int = 0,
) -> dict:
    sheet_id = current_client().resolve_sheet_id(_sid(spreadsheet_id), sheet)
    if width_pixels <= 0 or height_pixels <= 0:
        raise ValueError("width_pixels and height_pixels must be positive.")
    return {
        "overlayPosition": {
            "anchorCell": _anchor_coordinate(sheet_id, anchor_cell),
            "offsetXPixels": offset_x_pixels,
            "offsetYPixels": offset_y_pixels,
            "widthPixels": width_pixels,
            "heightPixels": height_pixels,
        }
    }


def _field_mask(fields: list[str] | None, default: str = "*") -> str:
    if not fields:
        return default
    return ",".join(fields)


def _condition(condition_type: str, values: list[str] | None = None) -> dict:
    ctype = condition_type.upper()
    if ctype == "CUSTOM_FORMULA" and values:
        for v in values:
            match = _NETWORK_FORMULA_RE.search(v)
            if match:
                raise ValueError(
                    f"CUSTOM_FORMULA value contains a network-fetching function "
                    f"({match.group(1)}), which is not allowed."
                )
    condition: dict[str, Any] = {"type": ctype}
    if values:
        condition["values"] = [{"userEnteredValue": v} for v in values]
    return condition


def _grid_span(grid: dict[str, int], start_key: str, end_key: str) -> int | None:
    if start_key not in grid or end_key not in grid:
        return None
    return grid[end_key] - grid[start_key]


def _validate_chart_source_grid(range_a1: str, grid: dict[str, int]) -> None:
    """Validate that a source range is suitable for a Sheets chart series/domain."""
    row_span = _grid_span(grid, "startRowIndex", "endRowIndex")
    col_span = _grid_span(grid, "startColumnIndex", "endColumnIndex")
    if row_span is None and col_span is None:
        raise ValueError(f"Chart source range {range_a1!r} must identify cells, not a whole sheet.")
    if row_span is not None and row_span <= 0:
        raise ValueError(f"Chart source range {range_a1!r} has no rows.")
    if col_span is not None and col_span <= 0:
        raise ValueError(f"Chart source range {range_a1!r} has no columns.")
    if row_span != 1 and col_span != 1:
        raise ValueError(
            f"Chart source range {range_a1!r} must be a single row or single column "
            "(for example, Sheet1!A2:A10 or Sheet1!B1:F1)."
        )


def _anchor_coordinate(sheet_id: int, anchor_cell: str) -> dict[str, int]:
    sheet_name, cell_ref = split_sheet_and_range(anchor_cell)
    if sheet_name is not None:
        raise ValueError("anchor_cell must be a sheet-relative single cell like 'E2'; use sheet to choose the tab.")

    grid = a1_to_grid_range(cell_ref)
    if (
        "startRowIndex" not in grid
        or "startColumnIndex" not in grid
        or grid.get("endRowIndex") != grid["startRowIndex"] + 1
        or grid.get("endColumnIndex") != grid["startColumnIndex"] + 1
    ):
        raise ValueError("anchor_cell must be a single cell like 'E2'.")

    return {
        "sheetId": sheet_id,
        "rowIndex": grid["startRowIndex"],
        "columnIndex": grid["startColumnIndex"],
    }


def _chart_data(grid: dict[str, int]) -> dict:
    return {"sourceRange": {"sources": [grid]}}


def _normalize_chart_type(chart_type: str) -> str:
    ctype = chart_type.upper()
    if ctype not in _BASIC_CHART_TYPES and ctype not in _PIE_CHART_TYPES:
        allowed = ", ".join(sorted(_BASIC_CHART_TYPES | _PIE_CHART_TYPES))
        raise ValueError(f"Invalid chart_type {chart_type!r}. Choose from: {allowed}.")
    return ctype


def _build_add_chart_request(
    *,
    chart_type: str,
    domain_grid: dict[str, int],
    series_grids: list[dict[str, int]],
    position_sheet_id: int,
    anchor_cell: str,
    title: str | None = None,
    legend_position: str = "RIGHT_LEGEND",
    x_axis_title: str | None = None,
    y_axis_title: str | None = None,
    header_count: int | None = None,
    stacked_type: str | None = None,
    width_pixels: int = 600,
    height_pixels: int = 371,
    offset_x_pixels: int = 0,
    offset_y_pixels: int = 0,
    pie_hole: float | None = None,
    three_dimensional: bool = False,
) -> dict:
    """Build a Sheets API ``addChart`` request from already-resolved grid ranges."""
    ctype = _normalize_chart_type(chart_type)
    if not series_grids:
        raise ValueError("Provide at least one series range.")
    if width_pixels <= 0 or height_pixels <= 0:
        raise ValueError("width_pixels and height_pixels must be positive.")
    if header_count is not None and header_count < 0:
        raise ValueError("header_count must be >= 0.")
    if pie_hole is not None and not 0 <= pie_hole <= 1:
        raise ValueError("pie_hole must be between 0 and 1.")

    legend = legend_position.upper()
    spec: dict[str, Any] = {}
    if title:
        spec["title"] = title

    if ctype in _PIE_CHART_TYPES:
        if len(series_grids) != 1:
            raise ValueError("PIE and DONUT charts require exactly one series range.")
        if legend not in _PIE_LEGEND_POSITIONS:
            allowed = ", ".join(sorted(_PIE_LEGEND_POSITIONS))
            raise ValueError(f"Invalid pie legend_position {legend_position!r}. Choose from: {allowed}.")
        pie_spec: dict[str, Any] = {
            "domain": _chart_data(domain_grid),
            "series": _chart_data(series_grids[0]),
            "legendPosition": legend,
        }
        if ctype == "DONUT":
            pie_spec["pieHole"] = 0.4 if pie_hole is None else pie_hole
        elif pie_hole is not None:
            pie_spec["pieHole"] = pie_hole
        if three_dimensional:
            pie_spec["threeDimensional"] = True
        spec["pieChart"] = pie_spec
    else:
        if legend not in _BASIC_LEGEND_POSITIONS:
            allowed = ", ".join(sorted(_BASIC_LEGEND_POSITIONS))
            raise ValueError(f"Invalid legend_position {legend_position!r}. Choose from: {allowed}.")
        basic_spec: dict[str, Any] = {
            "chartType": ctype,
            "legendPosition": legend,
            "domains": [{"domain": _chart_data(domain_grid)}],
            "series": [
                {"series": _chart_data(grid), "targetAxis": "LEFT_AXIS"}
                for grid in series_grids
            ],
        }
        if header_count is not None:
            basic_spec["headerCount"] = header_count
        if stacked_type is not None:
            stacked = stacked_type.upper()
            if stacked not in _STACKED_TYPES:
                allowed = ", ".join(sorted(_STACKED_TYPES))
                raise ValueError(f"Invalid stacked_type {stacked_type!r}. Choose from: {allowed}.")
            basic_spec["stackedType"] = stacked

        axes = []
        if x_axis_title:
            axes.append({"position": "BOTTOM_AXIS", "title": x_axis_title})
        if y_axis_title:
            axes.append({"position": "LEFT_AXIS", "title": y_axis_title})
        if axes:
            basic_spec["axis"] = axes
        spec["basicChart"] = basic_spec

    return {
        "addChart": {
            "chart": {
                "spec": spec,
                "position": {
                    "overlayPosition": {
                        "anchorCell": _anchor_coordinate(position_sheet_id, anchor_cell),
                        "offsetXPixels": offset_x_pixels,
                        "offsetYPixels": offset_y_pixels,
                        "widthPixels": width_pixels,
                        "heightPixels": height_pixels,
                    }
                },
            }
        }
    }


# ---------------------------------------------------------------------------
# Spreadsheet / sheet structure
# ---------------------------------------------------------------------------


@mcp.tool()
def get_spreadsheet_info(spreadsheet_id: str) -> dict:
    """Get a spreadsheet's title, URL, and the list of its sheets.

    Returns each sheet's title, numeric ``sheetId``, index, dimensions, and any
    frozen rows/columns. Use this first to discover sheet names and sizes.
    """
    meta = current_client().get_metadata(_sid(spreadsheet_id))
    sheets = []
    for s in meta.get("sheets", []):
        p = s["properties"]
        grid = p.get("gridProperties", {})
        sheets.append(
            {
                "title": p["title"],
                "sheetId": p["sheetId"],
                "index": p.get("index"),
                "rowCount": grid.get("rowCount"),
                "columnCount": grid.get("columnCount"),
                "frozenRowCount": grid.get("frozenRowCount", 0),
                "frozenColumnCount": grid.get("frozenColumnCount", 0),
                "charts": [
                    {
                        "chartId": c.get("chartId"),
                        "title": c.get("spec", {}).get("title"),
                        "position": c.get("position"),
                    }
                    for c in s.get("charts", [])
                ],
                "filterViews": [
                    {"filterViewId": f.get("filterViewId"), "title": f.get("title"), "range": f.get("range")}
                    for f in s.get("filterViews", [])
                ],
                "protectedRanges": [
                    {
                        "protectedRangeId": pr.get("protectedRangeId"),
                        "description": pr.get("description"),
                        "warningOnly": pr.get("warningOnly"),
                        "range": pr.get("range"),
                    }
                    for pr in s.get("protectedRanges", [])
                ],
                "bandedRanges": [
                    {"bandedRangeId": br.get("bandedRangeId"), "range": br.get("range")}
                    for br in s.get("bandedRanges", [])
                ],
                "slicers": [
                    {
                        "slicerId": sl.get("slicerId"),
                        "title": sl.get("spec", {}).get("title"),
                        "position": sl.get("position"),
                    }
                    for sl in s.get("slicers", [])
                ],
                "tables": [
                    {"tableId": t.get("tableId"), "name": t.get("name"), "range": t.get("range")}
                    for t in s.get("tables", [])
                ],
            }
        )
    return {
        "spreadsheetId": meta.get("spreadsheetId"),
        "title": meta.get("properties", {}).get("title"),
        "url": meta.get("spreadsheetUrl"),
        "namedRanges": [
            {"namedRangeId": nr.get("namedRangeId"), "name": nr.get("name"), "range": nr.get("range")}
            for nr in meta.get("namedRanges", [])
        ],
        "sheets": sheets,
    }


@mcp.tool()
def create_spreadsheet(title: str, sheet_titles: list[str] | None = None) -> dict:
    """Create a new spreadsheet and return its id and URL.

    Optionally provide ``sheet_titles`` to seed multiple named sheets/tabs.
    """
    body: dict[str, Any] = {"properties": {"title": title}}
    if sheet_titles:
        body["sheets"] = [{"properties": {"title": t}} for t in sheet_titles]
    result = _api(
        current_client().spreadsheets.create,
        body=body,
        fields="spreadsheetId,spreadsheetUrl,properties.title",
    )
    return {
        "spreadsheetId": result["spreadsheetId"],
        "url": result["spreadsheetUrl"],
        "title": result["properties"]["title"],
    }


@mcp.tool()
def add_sheet(
    spreadsheet_id: str,
    title: str,
    row_count: int = 1000,
    column_count: int = 26,
) -> dict:
    """Add a new sheet/tab to an existing spreadsheet."""
    req = {
        "addSheet": {
            "properties": {
                "title": title,
                "gridProperties": {"rowCount": row_count, "columnCount": column_count},
            }
        }
    }
    result = current_client().batch_update(_sid(spreadsheet_id), [req])
    props = result["replies"][0]["addSheet"]["properties"]
    return {"title": props["title"], "sheetId": props["sheetId"]}


@mcp.tool()
def rename_sheet(spreadsheet_id: str, sheet: str, new_title: str) -> dict:
    """Rename a sheet/tab."""
    sheet_id = current_client().resolve_sheet_id(_sid(spreadsheet_id), sheet)
    req = {
        "updateSheetProperties": {
            "properties": {"sheetId": sheet_id, "title": new_title},
            "fields": "title",
        }
    }
    current_client().batch_update(spreadsheet_id, [req])
    return {"sheetId": sheet_id, "title": new_title}


@mcp.tool()
def delete_sheet(spreadsheet_id: str, sheet: str) -> dict:
    """Delete a sheet/tab by name or numeric sheetId."""
    sheet_id = current_client().resolve_sheet_id(_sid(spreadsheet_id), sheet)
    _execute_batch_update(spreadsheet_id, [{"deleteSheet": {"sheetId": sheet_id}}])
    return {"sheetId": sheet_id, "deleted": True}


@mcp.tool()
def duplicate_sheet(
    spreadsheet_id: str,
    sheet: str,
    new_title: str | None = None,
    insert_index: int | None = None,
) -> dict:
    """Duplicate a sheet/tab, optionally setting the new title and index."""
    source_sheet_id = current_client().resolve_sheet_id(_sid(spreadsheet_id), sheet)
    body: dict[str, Any] = {"sourceSheetId": source_sheet_id}
    if new_title is not None:
        body["newSheetName"] = new_title
    if insert_index is not None:
        body["insertSheetIndex"] = insert_index
    result = _execute_batch_update(spreadsheet_id, [{"duplicateSheet": body}])
    props = result.get("replies", [{}])[0].get("duplicateSheet", {}).get("properties", {})
    return {
        "sourceSheetId": source_sheet_id,
        "sheetId": props.get("sheetId"),
        "title": props.get("title", new_title),
    }


@mcp.tool()
def batch_update_advanced(
    spreadsheet_id: str,
    requests: list[dict],
    response_ranges: list[str] | None = None,
    include_spreadsheet_in_response: bool = False,
) -> dict:
    """Run allowlisted raw Sheets API ``batchUpdate`` requests.

    Each request must contain exactly one allowed request kind. Data-source
    lifecycle requests are blocked, and a call may contain at most 50 requests.
    Prefer the typed tools when one exists; use this for advanced API coverage.
    """
    kinds = _validate_advanced_requests(requests)
    result = _execute_batch_update(
        spreadsheet_id,
        requests,
        response_ranges=response_ranges,
        include_spreadsheet_in_response=include_spreadsheet_in_response,
    )
    return {"requestKinds": kinds, "replies": result.get("replies", []), "spreadsheetId": result.get("spreadsheetId")}


# ---------------------------------------------------------------------------
# Cell values
# ---------------------------------------------------------------------------


@mcp.tool()
def read_range(
    spreadsheet_id: str,
    range: str,
    value_render_option: str = "FORMATTED_VALUE",
) -> dict:
    """Read cell values from an A1 range (e.g. ``Sheet1!A1:C10``).

    ``value_render_option``: ``FORMATTED_VALUE`` (as shown, default),
    ``UNFORMATTED_VALUE`` (typed values), or ``FORMULA`` (underlying formulas).
    Returns ``{"range", "values"}`` where ``values`` is a list of rows.
    """
    result = _api(
        current_client().spreadsheets.values().get,
        spreadsheetId=_sid(spreadsheet_id),
        range=range,
        valueRenderOption=_render_option(value_render_option),
    )
    return {"range": result.get("range"), "values": result.get("values", [])}


@mcp.tool()
def batch_read(
    spreadsheet_id: str,
    ranges: list[str],
    value_render_option: str = "FORMATTED_VALUE",
) -> dict:
    """Read several A1 ranges at once. Returns ``{"valueRanges": [...]}``."""
    result = _api(
        current_client().spreadsheets.values().batchGet,
        spreadsheetId=_sid(spreadsheet_id),
        ranges=ranges,
        valueRenderOption=_render_option(value_render_option),
    )
    return {
        "valueRanges": [
            {"range": vr.get("range"), "values": vr.get("values", [])}
            for vr in result.get("valueRanges", [])
        ]
    }


@mcp.tool()
def write_range(
    spreadsheet_id: str,
    range: str,
    values: list[list[Any]],
    value_input_option: str = "RAW",
) -> dict:
    """Write a 2D array of values to an A1 range.

    ``value_input_option``: ``RAW`` (store verbatim as strings; default) or
    ``USER_ENTERED`` (parse like a user typing — formulas, dates, and numbers
    are evaluated). Opt in to ``USER_ENTERED`` explicitly; it evaluates formulas
    including those that can make external network requests.
    """
    result = _api(
        current_client().spreadsheets.values().update,
        spreadsheetId=_sid(spreadsheet_id),
        range=range,
        valueInputOption=_input_option(value_input_option),
        body={"values": values},
    )
    return {
        "updatedRange": result.get("updatedRange"),
        "updatedRows": result.get("updatedRows"),
        "updatedColumns": result.get("updatedColumns"),
        "updatedCells": result.get("updatedCells"),
    }


@mcp.tool()
def append_rows(
    spreadsheet_id: str,
    range: str,
    values: list[list[Any]],
    value_input_option: str = "RAW",
) -> dict:
    """Append rows after the last row of data within/near the given A1 range.

    The ``range`` selects the table to append to (e.g. ``Sheet1!A1`` or
    ``Sheet1!A:D``); new rows are added below the existing data.

    ``value_input_option``: ``RAW`` (verbatim strings; default) or
    ``USER_ENTERED`` (evaluate formulas/dates/numbers — opt in explicitly).
    """
    result = _api(
        current_client().spreadsheets.values().append,
        spreadsheetId=_sid(spreadsheet_id),
        range=range,
        valueInputOption=_input_option(value_input_option),
        insertDataOption="INSERT_ROWS",
        body={"values": values},
    )
    updates = result.get("updates", {})
    return {
        "updatedRange": updates.get("updatedRange"),
        "updatedRows": updates.get("updatedRows"),
        "updatedCells": updates.get("updatedCells"),
    }


@mcp.tool()
def clear_range(spreadsheet_id: str, range: str) -> dict:
    """Clear the *values* in an A1 range (formatting is preserved)."""
    result = _api(
        current_client().spreadsheets.values().clear,
        spreadsheetId=_sid(spreadsheet_id),
        range=range,
        body={},
    )
    return {"clearedRange": result.get("clearedRange")}


@mcp.tool()
def batch_write(
    spreadsheet_id: str,
    data: list[dict],
    value_input_option: str = "RAW",
) -> dict:
    """Write values to multiple ranges in one values ``batchUpdate`` call.

    ``data`` items use the Sheets API ``ValueRange`` shape, e.g.
    ``{"range": "Sheet1!A1:B2", "values": [[1, 2], [3, 4]]}``.
    """
    result = _api(
        current_client().spreadsheets.values().batchUpdate,
        spreadsheetId=_sid(spreadsheet_id),
        body={"valueInputOption": _input_option(value_input_option), "data": data},
    )
    return {
        "totalUpdatedRows": result.get("totalUpdatedRows"),
        "totalUpdatedColumns": result.get("totalUpdatedColumns"),
        "totalUpdatedCells": result.get("totalUpdatedCells"),
        "responses": result.get("responses", []),
    }


@mcp.tool()
def batch_clear(spreadsheet_id: str, ranges: list[str]) -> dict:
    """Clear values from multiple A1 ranges in one call."""
    result = _api(
        current_client().spreadsheets.values().batchClear,
        spreadsheetId=_sid(spreadsheet_id),
        body={"ranges": ranges},
    )
    return {"clearedRanges": result.get("clearedRanges", [])}


# ---------------------------------------------------------------------------
# Cell formatting
# ---------------------------------------------------------------------------


@mcp.tool()
def read_formatting(
    spreadsheet_id: str,
    range: str,
    include_defaults: bool = False,
) -> dict:
    """Read the formatting you see in a bounded A1 range.

    Adjacent cells with the same formatting come back as one run, so a yellow
    bold header is a single range instead of one entry per cell. Colors are
    ``#rrggbb`` from the effective format, which includes conditional formatting.

    White backgrounds, black text, and the default 10pt font are left out.
    A run is returned when a cell has a color, bold/italic/underline/
    strikethrough, a number format, or a note. Set ``include_defaults`` to
    keep the plain cells too.

    The range must be bounded (``A1:C20``, not ``A:C``) and at most 2500 cells.
    """
    assert_bounded_range(range)
    response = _api(
        current_client().spreadsheets.get,
        spreadsheetId=_sid(spreadsheet_id),
        ranges=[range],
        includeGridData=True,
        fields=FORMAT_FIELDS,
    )
    return {"range": range, "runs": runs_from_grid(response, include_defaults)}


@mcp.tool()
def format_cells(
    spreadsheet_id: str,
    range: str,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    strikethrough: bool | None = None,
    font_size: int | None = None,
    font_family: str | None = None,
    text_color: str | None = None,
    background_color: str | None = None,
    horizontal_alignment: str | None = None,
    vertical_alignment: str | None = None,
    wrap_strategy: str | None = None,
    number_format_type: str | None = None,
    number_format_pattern: str | None = None,
) -> dict:
    """Apply cell formatting to an A1 range. Only the arguments you pass change.

    * ``text_color`` / ``background_color``: ``"#RRGGBB"``, a named color
      (red, blue, lightgray, ...), or omit to leave unchanged.
    * ``horizontal_alignment``: ``LEFT`` | ``CENTER`` | ``RIGHT``.
    * ``vertical_alignment``: ``TOP`` | ``MIDDLE`` | ``BOTTOM``.
    * ``wrap_strategy``: ``OVERFLOW_CELL`` | ``CLIP`` | ``WRAP``.
    * ``number_format_type``: ``TEXT`` | ``NUMBER`` | ``PERCENT`` | ``CURRENCY``
      | ``DATE`` | ``TIME`` | ``DATE_TIME`` | ``SCIENTIFIC``; pair with an
      optional ``number_format_pattern`` (e.g. ``"$#,##0.00"``, ``"0.0%"``).
    """
    cell_format: dict[str, Any] = {}
    fields: list[str] = []
    text_format: dict[str, Any] = {}

    if bold is not None:
        text_format["bold"] = bold
        fields.append("userEnteredFormat.textFormat.bold")
    if italic is not None:
        text_format["italic"] = italic
        fields.append("userEnteredFormat.textFormat.italic")
    if underline is not None:
        text_format["underline"] = underline
        fields.append("userEnteredFormat.textFormat.underline")
    if strikethrough is not None:
        text_format["strikethrough"] = strikethrough
        fields.append("userEnteredFormat.textFormat.strikethrough")
    if font_size is not None:
        text_format["fontSize"] = font_size
        fields.append("userEnteredFormat.textFormat.fontSize")
    if font_family is not None:
        text_format["fontFamily"] = font_family
        fields.append("userEnteredFormat.textFormat.fontFamily")
    if text_color is not None:
        text_format["foregroundColor"] = parse_color(text_color)
        fields.append("userEnteredFormat.textFormat.foregroundColor")

    if text_format:
        cell_format["textFormat"] = text_format

    if background_color is not None:
        cell_format["backgroundColor"] = parse_color(background_color)
        fields.append("userEnteredFormat.backgroundColor")
    if horizontal_alignment is not None:
        cell_format["horizontalAlignment"] = horizontal_alignment.upper()
        fields.append("userEnteredFormat.horizontalAlignment")
    if vertical_alignment is not None:
        cell_format["verticalAlignment"] = vertical_alignment.upper()
        fields.append("userEnteredFormat.verticalAlignment")
    if wrap_strategy is not None:
        cell_format["wrapStrategy"] = wrap_strategy.upper()
        fields.append("userEnteredFormat.wrapStrategy")
    if number_format_type is not None:
        nf: dict[str, Any] = {"type": number_format_type.upper()}
        if number_format_pattern is not None:
            nf["pattern"] = number_format_pattern
        cell_format["numberFormat"] = nf
        fields.append("userEnteredFormat.numberFormat")

    if not fields:
        raise ValueError("No formatting options provided — pass at least one.")

    grid = current_client().resolve_grid_range(_sid(spreadsheet_id), range)
    req = {
        "repeatCell": {
            "range": grid,
            "cell": {"userEnteredFormat": cell_format},
            "fields": ",".join(fields),
        }
    }
    current_client().batch_update(spreadsheet_id, [req])
    return {"range": range, "appliedFields": fields}


@mcp.tool()
def set_borders(
    spreadsheet_id: str,
    range: str,
    style: str = "SOLID",
    color: str = "#000000",
    top: bool = True,
    bottom: bool = True,
    left: bool = True,
    right: bool = True,
    inner_horizontal: bool = False,
    inner_vertical: bool = False,
) -> dict:
    """Set borders on an A1 range.

    ``style``: ``SOLID`` | ``SOLID_MEDIUM`` | ``SOLID_THICK`` | ``DASHED`` |
    ``DOTTED`` | ``DOUBLE`` | ``NONE`` (use NONE to clear). ``top``/``bottom``/
    ``left``/``right`` are the outer edges; ``inner_*`` set borders between cells.
    """
    border = {"style": style.upper(), "color": parse_color(color)}
    body: dict[str, Any] = {"range": current_client().resolve_grid_range(_sid(spreadsheet_id), range)}
    for edge, enabled in (
        ("top", top),
        ("bottom", bottom),
        ("left", left),
        ("right", right),
        ("innerHorizontal", inner_horizontal),
        ("innerVertical", inner_vertical),
    ):
        if enabled:
            body[edge] = dict(border)

    current_client().batch_update(spreadsheet_id, [{"updateBorders": body}])
    return {"range": range, "style": style.upper()}


@mcp.tool()
def merge_cells(spreadsheet_id: str, range: str, merge_type: str = "MERGE_ALL") -> dict:
    """Merge cells in an A1 range.

    ``merge_type``: ``MERGE_ALL`` (one cell), ``MERGE_COLUMNS`` (merge each column
    vertically), or ``MERGE_ROWS`` (merge each row horizontally).
    """
    grid = current_client().resolve_grid_range(_sid(spreadsheet_id), range)
    current_client().batch_update(
        spreadsheet_id, [{"mergeCells": {"range": grid, "mergeType": merge_type.upper()}}]
    )
    return {"range": range, "mergeType": merge_type.upper()}


@mcp.tool()
def unmerge_cells(spreadsheet_id: str, range: str) -> dict:
    """Unmerge any merged cells overlapping an A1 range."""
    grid = current_client().resolve_grid_range(_sid(spreadsheet_id), range)
    current_client().batch_update(spreadsheet_id, [{"unmergeCells": {"range": grid}}])
    return {"range": range, "unmerged": True}


# ---------------------------------------------------------------------------
# Dimensions: column widths, row heights, freezing
# ---------------------------------------------------------------------------


@mcp.tool()
def set_dimension_size(
    spreadsheet_id: str,
    sheet: str,
    dimension: str,
    start_index: int,
    end_index: int,
    pixel_size: int,
) -> dict:
    """Set a fixed pixel size for a band of columns or rows.

    ``dimension``: ``COLUMNS`` or ``ROWS``. Indices are zero-based and half-open
    (``start_index=0, end_index=3`` covers the first three columns/rows).
    """
    sheet_id = current_client().resolve_sheet_id(_sid(spreadsheet_id), sheet)
    req = {
        "updateDimensionProperties": {
            "range": {
                "sheetId": sheet_id,
                "dimension": dimension.upper(),
                "startIndex": start_index,
                "endIndex": end_index,
            },
            "properties": {"pixelSize": pixel_size},
            "fields": "pixelSize",
        }
    }
    current_client().batch_update(spreadsheet_id, [req])
    return {"sheetId": sheet_id, "dimension": dimension.upper(), "pixelSize": pixel_size}


@mcp.tool()
def auto_resize_dimensions(
    spreadsheet_id: str,
    sheet: str,
    dimension: str,
    start_index: int,
    end_index: int,
) -> dict:
    """Auto-fit columns or rows to their contents (zero-based, half-open range)."""
    sheet_id = current_client().resolve_sheet_id(_sid(spreadsheet_id), sheet)
    req = {
        "autoResizeDimensions": {
            "dimensions": {
                "sheetId": sheet_id,
                "dimension": dimension.upper(),
                "startIndex": start_index,
                "endIndex": end_index,
            }
        }
    }
    current_client().batch_update(spreadsheet_id, [req])
    return {"sheetId": sheet_id, "dimension": dimension.upper(), "autoResized": True}


@mcp.tool()
def freeze(
    spreadsheet_id: str,
    sheet: str,
    rows: int | None = None,
    columns: int | None = None,
) -> dict:
    """Freeze the first ``rows`` rows and/or ``columns`` columns of a sheet.

    Pass ``0`` to unfreeze. Leaving an argument as ``None`` keeps it unchanged.
    """
    sheet_id = current_client().resolve_sheet_id(_sid(spreadsheet_id), sheet)
    grid_props: dict[str, int] = {}
    fields: list[str] = []
    if rows is not None:
        grid_props["frozenRowCount"] = rows
        fields.append("gridProperties.frozenRowCount")
    if columns is not None:
        grid_props["frozenColumnCount"] = columns
        fields.append("gridProperties.frozenColumnCount")
    if not fields:
        raise ValueError("Pass rows and/or columns to freeze.")

    req = {
        "updateSheetProperties": {
            "properties": {"sheetId": sheet_id, "gridProperties": grid_props},
            "fields": ",".join(fields),
        }
    }
    current_client().batch_update(spreadsheet_id, [req])
    return {"sheetId": sheet_id, "frozenRows": rows, "frozenColumns": columns}


# ---------------------------------------------------------------------------
# Sheet data operations
# ---------------------------------------------------------------------------


@mcp.tool()
def insert_rows(
    spreadsheet_id: str,
    sheet: str,
    start_index: int,
    end_index: int,
    inherit_from_before: bool = False,
) -> dict:
    """Insert rows before ``start_index``; indices are zero-based and half-open."""
    rng = _dimension_range(spreadsheet_id, sheet, "ROWS", start_index, end_index)
    req = {"insertDimension": {"range": rng, "inheritFromBefore": inherit_from_before}}
    _execute_batch_update(spreadsheet_id, [req])
    return {"sheetId": rng["sheetId"], "dimension": "ROWS", "startIndex": start_index, "endIndex": end_index}


@mcp.tool()
def delete_rows(spreadsheet_id: str, sheet: str, start_index: int, end_index: int) -> dict:
    """Delete rows; indices are zero-based and half-open."""
    rng = _dimension_range(spreadsheet_id, sheet, "ROWS", start_index, end_index)
    _execute_batch_update(spreadsheet_id, [{"deleteDimension": {"range": rng}}])
    return {"sheetId": rng["sheetId"], "dimension": "ROWS", "startIndex": start_index, "endIndex": end_index}


@mcp.tool()
def insert_columns(
    spreadsheet_id: str,
    sheet: str,
    start_index: int,
    end_index: int,
    inherit_from_before: bool = False,
) -> dict:
    """Insert columns before ``start_index``; indices are zero-based and half-open."""
    rng = _dimension_range(spreadsheet_id, sheet, "COLUMNS", start_index, end_index)
    req = {"insertDimension": {"range": rng, "inheritFromBefore": inherit_from_before}}
    _execute_batch_update(spreadsheet_id, [req])
    return {"sheetId": rng["sheetId"], "dimension": "COLUMNS", "startIndex": start_index, "endIndex": end_index}


@mcp.tool()
def delete_columns(spreadsheet_id: str, sheet: str, start_index: int, end_index: int) -> dict:
    """Delete columns; indices are zero-based and half-open."""
    rng = _dimension_range(spreadsheet_id, sheet, "COLUMNS", start_index, end_index)
    _execute_batch_update(spreadsheet_id, [{"deleteDimension": {"range": rng}}])
    return {"sheetId": rng["sheetId"], "dimension": "COLUMNS", "startIndex": start_index, "endIndex": end_index}


@mcp.tool()
def append_rows_to_sheet(spreadsheet_id: str, sheet: str, length: int) -> dict:
    """Append empty rows to the end of a sheet."""
    if length <= 0:
        raise ValueError("length must be positive.")
    sheet_id = current_client().resolve_sheet_id(_sid(spreadsheet_id), sheet)
    _execute_batch_update(spreadsheet_id, [{"appendDimension": {"sheetId": sheet_id, "dimension": "ROWS", "length": length}}])
    return {"sheetId": sheet_id, "dimension": "ROWS", "length": length}


@mcp.tool()
def append_columns_to_sheet(spreadsheet_id: str, sheet: str, length: int) -> dict:
    """Append empty columns to the end of a sheet."""
    if length <= 0:
        raise ValueError("length must be positive.")
    sheet_id = current_client().resolve_sheet_id(_sid(spreadsheet_id), sheet)
    _execute_batch_update(spreadsheet_id, [{"appendDimension": {"sheetId": sheet_id, "dimension": "COLUMNS", "length": length}}])
    return {"sheetId": sheet_id, "dimension": "COLUMNS", "length": length}


@mcp.tool()
def move_rows_or_columns(
    spreadsheet_id: str,
    sheet: str,
    dimension: str,
    start_index: int,
    end_index: int,
    destination_index: int,
) -> dict:
    """Move rows or columns to ``destination_index``; indices are zero-based."""
    rng = _dimension_range(spreadsheet_id, sheet, dimension, start_index, end_index)
    req = {"moveDimension": {"source": rng, "destinationIndex": destination_index}}
    _execute_batch_update(spreadsheet_id, [req])
    return {**rng, "destinationIndex": destination_index}


@mcp.tool()
def insert_range(spreadsheet_id: str, range: str, shift_dimension: str = "ROWS") -> dict:
    """Insert cells in an A1 range, shifting existing cells by ROWS or COLUMNS."""
    grid = current_client().resolve_grid_range(_sid(spreadsheet_id), range)
    shift = _dimension(shift_dimension)
    _execute_batch_update(spreadsheet_id, [{"insertRange": {"range": grid, "shiftDimension": shift}}])
    return {"range": range, "shiftDimension": shift}


@mcp.tool()
def delete_range(spreadsheet_id: str, range: str, shift_dimension: str = "ROWS") -> dict:
    """Delete cells in an A1 range, shifting remaining cells by ROWS or COLUMNS."""
    grid = current_client().resolve_grid_range(_sid(spreadsheet_id), range)
    shift = _dimension(shift_dimension)
    _execute_batch_update(spreadsheet_id, [{"deleteRange": {"range": grid, "shiftDimension": shift}}])
    return {"range": range, "shiftDimension": shift}


@mcp.tool()
def copy_range(
    spreadsheet_id: str,
    source_range: str,
    destination_range: str,
    paste_type: str = "PASTE_NORMAL",
    paste_orientation: str = "NORMAL",
) -> dict:
    """Copy a range to another range."""
    client = current_client()
    sid = _sid(spreadsheet_id)
    req = {
        "copyPaste": {
            "source": client.resolve_grid_range(sid, source_range),
            "destination": client.resolve_grid_range(sid, destination_range),
            "pasteType": paste_type.upper(),
            "pasteOrientation": paste_orientation.upper(),
        }
    }
    _execute_batch_update(spreadsheet_id, [req])
    return {"sourceRange": source_range, "destinationRange": destination_range, "pasteType": paste_type.upper()}


@mcp.tool()
def cut_paste_range(
    spreadsheet_id: str,
    source_range: str,
    destination_cell: str,
    paste_type: str = "PASTE_NORMAL",
) -> dict:
    """Cut a range and paste it starting at a destination cell."""
    client = current_client()
    sid = _sid(spreadsheet_id)
    req = {
        "cutPaste": {
            "source": client.resolve_grid_range(sid, source_range),
            "destination": _single_cell_coordinate(sid, destination_cell),
            "pasteType": paste_type.upper(),
        }
    }
    _execute_batch_update(spreadsheet_id, [req])
    return {"sourceRange": source_range, "destinationCell": destination_cell, "pasteType": paste_type.upper()}


@mcp.tool()
def paste_delimited_data(
    spreadsheet_id: str,
    start_cell: str,
    data: str,
    delimiter: str = ",",
    paste_type: str = "PASTE_NORMAL",
    html: bool = False,
) -> dict:
    """Paste delimited text/HTML starting at a single cell."""
    req = {
        "pasteData": {
            "coordinate": _single_cell_coordinate(_sid(spreadsheet_id), start_cell),
            "data": data,
            "delimiter": delimiter,
            "type": paste_type.upper(),
            "html": html,
        }
    }
    _execute_batch_update(spreadsheet_id, [req])
    return {"startCell": start_cell, "pasteType": paste_type.upper()}


@mcp.tool()
def sort_range(spreadsheet_id: str, range: str, sort_specs: list[dict]) -> dict:
    """Sort a range using Sheets API ``SortSpec`` objects."""
    if not sort_specs:
        raise ValueError("Provide at least one sort spec.")
    grid = current_client().resolve_grid_range(_sid(spreadsheet_id), range)
    _execute_batch_update(spreadsheet_id, [{"sortRange": {"range": grid, "sortSpecs": sort_specs}}])
    return {"range": range, "sortSpecs": sort_specs}


@mcp.tool()
def find_replace(
    spreadsheet_id: str,
    find: str,
    replacement: str,
    range: str | None = None,
    all_sheets: bool = False,
    match_case: bool = False,
    match_entire_cell: bool = False,
    search_by_regex: bool = False,
    include_formulas: bool = False,
) -> dict:
    """Find and replace text in one range or all sheets."""
    body: dict[str, Any] = {
        "find": find,
        "replacement": replacement,
        "matchCase": match_case,
        "matchEntireCell": match_entire_cell,
        "searchByRegex": search_by_regex,
        "includeFormulas": include_formulas,
    }
    if range:
        body["range"] = current_client().resolve_grid_range(_sid(spreadsheet_id), range)
    else:
        body["allSheets"] = all_sheets or True
    result = _execute_batch_update(spreadsheet_id, [{"findReplace": body}])
    reply = result.get("replies", [{}])[0].get("findReplace", {})
    return {"occurrencesChanged": reply.get("occurrencesChanged"), "rowsChanged": reply.get("rowsChanged"), "sheetsChanged": reply.get("sheetsChanged")}


@mcp.tool()
def trim_whitespace(spreadsheet_id: str, range: str) -> dict:
    """Trim leading/trailing/repeated whitespace in each cell in an A1 range."""
    grid = current_client().resolve_grid_range(_sid(spreadsheet_id), range)
    result = _execute_batch_update(spreadsheet_id, [{"trimWhitespace": {"range": grid}}])
    reply = result.get("replies", [{}])[0].get("trimWhitespace", {})
    return {"range": range, "cellsChangedCount": reply.get("cellsChangedCount")}


@mcp.tool()
def delete_duplicates(
    spreadsheet_id: str,
    range: str,
    comparison_columns: list[int] | None = None,
) -> dict:
    """Remove duplicate rows from a range, optionally comparing only column indexes."""
    grid = current_client().resolve_grid_range(_sid(spreadsheet_id), range)
    body: dict[str, Any] = {"range": grid}
    if comparison_columns:
        body["comparisonColumns"] = [
            {
                "sheetId": grid["sheetId"],
                "dimension": "COLUMNS",
                "startIndex": col,
                "endIndex": col + 1,
            }
            for col in comparison_columns
        ]
    result = _execute_batch_update(spreadsheet_id, [{"deleteDuplicates": body}])
    reply = result.get("replies", [{}])[0].get("deleteDuplicates", {})
    return {"range": range, "duplicatesRemovedCount": reply.get("duplicatesRemovedCount")}


@mcp.tool()
def text_to_columns(
    spreadsheet_id: str,
    source_range: str,
    delimiter_type: str = "AUTODETECT",
    delimiter: str | None = None,
) -> dict:
    """Split a single-column range into multiple columns."""
    body: dict[str, Any] = {
        "source": current_client().resolve_grid_range(_sid(spreadsheet_id), source_range),
        "delimiterType": delimiter_type.upper(),
    }
    if delimiter is not None:
        body["delimiter"] = delimiter
    _execute_batch_update(spreadsheet_id, [{"textToColumns": body}])
    return {"sourceRange": source_range, "delimiterType": delimiter_type.upper()}


@mcp.tool()
def auto_fill(
    spreadsheet_id: str,
    range: str,
    use_alternate_series: bool = False,
) -> dict:
    """Autofill within an A1 range by detecting existing source data."""
    grid = current_client().resolve_grid_range(_sid(spreadsheet_id), range)
    req = {
        "autoFill": {
            "range": grid,
            "useAlternateSeries": use_alternate_series,
        }
    }
    _execute_batch_update(spreadsheet_id, [req])
    return {"range": range, "useAlternateSeries": use_alternate_series}


@mcp.tool()
def randomize_range(spreadsheet_id: str, range: str) -> dict:
    """Randomize row order inside an A1 range."""
    grid = current_client().resolve_grid_range(_sid(spreadsheet_id), range)
    _execute_batch_update(spreadsheet_id, [{"randomizeRange": {"range": grid}}])
    return {"range": range, "randomized": True}


# ---------------------------------------------------------------------------
# Filters and validation
# ---------------------------------------------------------------------------


@mcp.tool()
def set_basic_filter(
    spreadsheet_id: str,
    range: str,
    sort_specs: list[dict] | None = None,
    filter_specs: list[dict] | None = None,
) -> dict:
    """Set the sheet's basic filter over an A1 range."""
    basic_filter: dict[str, Any] = {"range": current_client().resolve_grid_range(_sid(spreadsheet_id), range)}
    if sort_specs:
        basic_filter["sortSpecs"] = sort_specs
    if filter_specs:
        basic_filter["filterSpecs"] = filter_specs
    _execute_batch_update(spreadsheet_id, [{"setBasicFilter": {"filter": basic_filter}}])
    return {"range": range, "set": True}


@mcp.tool()
def clear_basic_filter(spreadsheet_id: str, sheet: str) -> dict:
    """Clear the basic filter on a sheet."""
    sheet_id = current_client().resolve_sheet_id(_sid(spreadsheet_id), sheet)
    _execute_batch_update(spreadsheet_id, [{"clearBasicFilter": {"sheetId": sheet_id}}])
    return {"sheetId": sheet_id, "cleared": True}


@mcp.tool()
def add_filter_view(
    spreadsheet_id: str,
    range: str,
    title: str | None = None,
    sort_specs: list[dict] | None = None,
    filter_specs: list[dict] | None = None,
) -> dict:
    """Add a filter view over an A1 range."""
    filter_view: dict[str, Any] = {"range": current_client().resolve_grid_range(_sid(spreadsheet_id), range)}
    if title is not None:
        filter_view["title"] = title
    if sort_specs:
        filter_view["sortSpecs"] = sort_specs
    if filter_specs:
        filter_view["filterSpecs"] = filter_specs
    result = _execute_batch_update(spreadsheet_id, [{"addFilterView": {"filter": filter_view}}])
    added = result.get("replies", [{}])[0].get("addFilterView", {}).get("filter", {})
    return {"filterViewId": added.get("filterViewId"), "title": added.get("title", title)}


@mcp.tool()
def update_filter_view(
    spreadsheet_id: str,
    filter_view: dict,
    fields: list[str] | None = None,
) -> dict:
    """Update a filter view using a Sheets API FilterView object."""
    _execute_batch_update(spreadsheet_id, [{"updateFilterView": {"filter": filter_view, "fields": _field_mask(fields)}}])
    return {"filterViewId": filter_view.get("filterViewId"), "updatedFields": fields or ["*"]}


@mcp.tool()
def delete_filter_view(spreadsheet_id: str, filter_view_id: int) -> dict:
    """Delete a filter view by ID."""
    _execute_batch_update(spreadsheet_id, [{"deleteFilterView": {"filterId": filter_view_id}}])
    return {"filterViewId": filter_view_id, "deleted": True}


@mcp.tool()
def set_data_validation(
    spreadsheet_id: str,
    range: str,
    condition_type: str,
    values: list[str] | None = None,
    input_message: str | None = None,
    strict: bool = True,
    show_custom_ui: bool = True,
    filtered_rows_included: bool = False,
) -> dict:
    """Set a data validation rule on an A1 range."""
    rule: dict[str, Any] = {
        "condition": _condition(condition_type, values),
        "strict": strict,
        "showCustomUi": show_custom_ui,
    }
    if input_message is not None:
        rule["inputMessage"] = input_message
    req = {
        "setDataValidation": {
            "range": current_client().resolve_grid_range(_sid(spreadsheet_id), range),
            "rule": rule,
            "filteredRowsIncluded": filtered_rows_included,
        }
    }
    _execute_batch_update(spreadsheet_id, [req])
    return {"range": range, "conditionType": condition_type.upper()}


@mcp.tool()
def clear_data_validation(spreadsheet_id: str, range: str) -> dict:
    """Clear data validation from an A1 range."""
    req = {"setDataValidation": {"range": current_client().resolve_grid_range(_sid(spreadsheet_id), range)}}
    _execute_batch_update(spreadsheet_id, [req])
    return {"range": range, "cleared": True}


# ---------------------------------------------------------------------------
# Named/protected ranges, banding, groups, and tables
# ---------------------------------------------------------------------------


@mcp.tool()
def add_named_range(spreadsheet_id: str, name: str, range: str) -> dict:
    """Create a named range over an A1 range."""
    named_range = {"name": name, "range": current_client().resolve_grid_range(_sid(spreadsheet_id), range)}
    result = _execute_batch_update(spreadsheet_id, [{"addNamedRange": {"namedRange": named_range}}])
    added = result.get("replies", [{}])[0].get("addNamedRange", {}).get("namedRange", {})
    return {"namedRangeId": added.get("namedRangeId"), "name": added.get("name", name)}


@mcp.tool()
def update_named_range(
    spreadsheet_id: str,
    named_range_id: str,
    name: str | None = None,
    range: str | None = None,
) -> dict:
    """Update a named range's name and/or range."""
    named_range: dict[str, Any] = {"namedRangeId": named_range_id}
    fields = []
    if name is not None:
        named_range["name"] = name
        fields.append("name")
    if range is not None:
        named_range["range"] = current_client().resolve_grid_range(_sid(spreadsheet_id), range)
        fields.append("range")
    if not fields:
        raise ValueError("Provide name and/or range.")
    req = {"updateNamedRange": {"namedRange": named_range, "fields": ",".join(fields)}}
    _execute_batch_update(spreadsheet_id, [req])
    return {"namedRangeId": named_range_id, "updatedFields": fields}


@mcp.tool()
def delete_named_range(spreadsheet_id: str, named_range_id: str) -> dict:
    """Delete a named range by ID."""
    _execute_batch_update(spreadsheet_id, [{"deleteNamedRange": {"namedRangeId": named_range_id}}])
    return {"namedRangeId": named_range_id, "deleted": True}


@mcp.tool()
def add_protected_range(
    spreadsheet_id: str,
    range: str,
    description: str | None = None,
    warning_only: bool = True,
    editors: dict | None = None,
) -> dict:
    """Protect an A1 range; defaults to warning-only protection."""
    protected_range: dict[str, Any] = {
        "range": current_client().resolve_grid_range(_sid(spreadsheet_id), range),
        "warningOnly": warning_only,
    }
    if description is not None:
        protected_range["description"] = description
    if editors is not None:
        protected_range["editors"] = editors
    result = _execute_batch_update(spreadsheet_id, [{"addProtectedRange": {"protectedRange": protected_range}}])
    added = result.get("replies", [{}])[0].get("addProtectedRange", {}).get("protectedRange", {})
    return {"protectedRangeId": added.get("protectedRangeId"), "description": added.get("description", description)}


@mcp.tool()
def update_protected_range(
    spreadsheet_id: str,
    protected_range: dict,
    fields: list[str] | None = None,
) -> dict:
    """Update a protected range using a Sheets API ProtectedRange object."""
    req = {"updateProtectedRange": {"protectedRange": protected_range, "fields": _field_mask(fields)}}
    _execute_batch_update(spreadsheet_id, [req])
    return {"protectedRangeId": protected_range.get("protectedRangeId"), "updatedFields": fields or ["*"]}


@mcp.tool()
def delete_protected_range(spreadsheet_id: str, protected_range_id: int) -> dict:
    """Delete a protected range by ID."""
    _execute_batch_update(spreadsheet_id, [{"deleteProtectedRange": {"protectedRangeId": protected_range_id}}])
    return {"protectedRangeId": protected_range_id, "deleted": True}


@mcp.tool()
def add_banding(
    spreadsheet_id: str,
    range: str,
    row_properties: dict | None = None,
    column_properties: dict | None = None,
) -> dict:
    """Add alternating-color banding to an A1 range."""
    banded_range: dict[str, Any] = {"range": current_client().resolve_grid_range(_sid(spreadsheet_id), range)}
    if row_properties is not None:
        banded_range["rowProperties"] = row_properties
    if column_properties is not None:
        banded_range["columnProperties"] = column_properties
    if "rowProperties" not in banded_range and "columnProperties" not in banded_range:
        banded_range["rowProperties"] = {
            "headerColor": parse_color("#d9ead3"),
            "firstBandColor": parse_color("#ffffff"),
            "secondBandColor": parse_color("#f3f3f3"),
        }
    result = _execute_batch_update(spreadsheet_id, [{"addBanding": {"bandedRange": banded_range}}])
    added = result.get("replies", [{}])[0].get("addBanding", {}).get("bandedRange", {})
    return {"bandedRangeId": added.get("bandedRangeId"), "range": range}


@mcp.tool()
def update_banding(spreadsheet_id: str, banded_range: dict, fields: list[str] | None = None) -> dict:
    """Update banding using a Sheets API BandedRange object."""
    _execute_batch_update(spreadsheet_id, [{"updateBanding": {"bandedRange": banded_range, "fields": _field_mask(fields)}}])
    return {"bandedRangeId": banded_range.get("bandedRangeId"), "updatedFields": fields or ["*"]}


@mcp.tool()
def delete_banding(spreadsheet_id: str, banded_range_id: int) -> dict:
    """Delete banding by ID."""
    _execute_batch_update(spreadsheet_id, [{"deleteBanding": {"bandedRangeId": banded_range_id}}])
    return {"bandedRangeId": banded_range_id, "deleted": True}


@mcp.tool()
def add_dimension_group(
    spreadsheet_id: str,
    sheet: str,
    dimension: str,
    start_index: int,
    end_index: int,
) -> dict:
    """Group rows or columns over a zero-based, half-open range."""
    rng = _dimension_range(spreadsheet_id, sheet, dimension, start_index, end_index)
    result = _execute_batch_update(spreadsheet_id, [{"addDimensionGroup": {"range": rng}}])
    group = result.get("replies", [{}])[0].get("addDimensionGroup", {}).get("dimensionGroups", [])
    return {"range": rng, "dimensionGroups": group}


@mcp.tool()
def update_dimension_group(
    spreadsheet_id: str,
    sheet: str,
    dimension: str,
    start_index: int,
    end_index: int,
    collapsed: bool,
) -> dict:
    """Collapse or expand a row/column group."""
    group = {
        "range": _dimension_range(spreadsheet_id, sheet, dimension, start_index, end_index),
        "collapsed": collapsed,
    }
    _execute_batch_update(spreadsheet_id, [{"updateDimensionGroup": {"dimensionGroup": group, "fields": "collapsed"}}])
    return {"range": group["range"], "collapsed": collapsed}


@mcp.tool()
def delete_dimension_group(
    spreadsheet_id: str,
    sheet: str,
    dimension: str,
    start_index: int,
    end_index: int,
) -> dict:
    """Delete a row/column group."""
    rng = _dimension_range(spreadsheet_id, sheet, dimension, start_index, end_index)
    _execute_batch_update(spreadsheet_id, [{"deleteDimensionGroup": {"range": rng}}])
    return {"range": rng, "deleted": True}


@mcp.tool()
def add_table(
    spreadsheet_id: str,
    range: str,
    name: str | None = None,
    column_properties: list[dict] | None = None,
    rows_properties: dict | None = None,
) -> dict:
    """Add a Sheets table over an A1 range."""
    table: dict[str, Any] = {"range": current_client().resolve_grid_range(_sid(spreadsheet_id), range)}
    if name is not None:
        table["name"] = name
    if column_properties:
        table["columnProperties"] = column_properties
    if rows_properties:
        table["rowsProperties"] = rows_properties
    result = _execute_batch_update(spreadsheet_id, [{"addTable": {"table": table}}])
    added = result.get("replies", [{}])[0].get("addTable", {}).get("table", {})
    return {"tableId": added.get("tableId"), "name": added.get("name", name)}


@mcp.tool()
def update_table(spreadsheet_id: str, table: dict, fields: list[str] | None = None) -> dict:
    """Update a table using a Sheets API Table object."""
    _execute_batch_update(spreadsheet_id, [{"updateTable": {"table": table, "fields": _field_mask(fields)}}])
    return {"tableId": table.get("tableId"), "updatedFields": fields or ["*"]}


@mcp.tool()
def delete_table(spreadsheet_id: str, table_id: str) -> dict:
    """Delete a table by ID."""
    _execute_batch_update(spreadsheet_id, [{"deleteTable": {"tableId": table_id}}])
    return {"tableId": table_id, "deleted": True}


# ---------------------------------------------------------------------------
# Charts
# ---------------------------------------------------------------------------


@mcp.tool()
def create_chart(
    spreadsheet_id: str,
    sheet: str,
    chart_type: str,
    domain_range: str,
    series_ranges: list[str],
    title: str | None = None,
    anchor_cell: str = "E2",
    legend_position: str = "RIGHT_LEGEND",
    x_axis_title: str | None = None,
    y_axis_title: str | None = None,
    header_count: int | None = None,
    stacked_type: str | None = None,
    width_pixels: int = 600,
    height_pixels: int = 371,
    offset_x_pixels: int = 0,
    offset_y_pixels: int = 0,
    pie_hole: float | None = None,
    three_dimensional: bool = False,
) -> dict:
    """Create an embedded chart from existing sheet data.

    ``chart_type``: ``COLUMN``, ``BAR``, ``LINE``, ``AREA``, ``SCATTER``,
    ``COMBO``, ``STEPPED_AREA``, ``PIE``, or ``DONUT``.

    ``domain_range`` is the category/label range; ``series_ranges`` is one or
    more value ranges. Each source must be a single row or single column, e.g.
    ``Sheet1!A2:A10`` plus ``Sheet1!B2:B10``. Ranges without a sheet prefix use
    ``sheet``. ``anchor_cell`` is where the chart is overlaid on that sheet.
    """
    sid = _sid(spreadsheet_id)
    client = current_client()
    position_sheet_id = client.resolve_sheet_id(sid, sheet)

    domain_grid = client.resolve_grid_range(sid, domain_range, default_sheet=sheet)
    _validate_chart_source_grid(domain_range, domain_grid)
    series_grids = []
    for range_a1 in series_ranges:
        grid = client.resolve_grid_range(sid, range_a1, default_sheet=sheet)
        _validate_chart_source_grid(range_a1, grid)
        series_grids.append(grid)

    req = _build_add_chart_request(
        chart_type=chart_type,
        domain_grid=domain_grid,
        series_grids=series_grids,
        position_sheet_id=position_sheet_id,
        anchor_cell=anchor_cell,
        title=title,
        legend_position=legend_position,
        x_axis_title=x_axis_title,
        y_axis_title=y_axis_title,
        header_count=header_count,
        stacked_type=stacked_type,
        width_pixels=width_pixels,
        height_pixels=height_pixels,
        offset_x_pixels=offset_x_pixels,
        offset_y_pixels=offset_y_pixels,
        pie_hole=pie_hole,
        three_dimensional=three_dimensional,
    )
    result = client.batch_update(sid, [req])
    chart = result.get("replies", [{}])[0].get("addChart", {}).get("chart", {})
    return {
        "chartId": chart.get("chartId"),
        "sheetId": position_sheet_id,
        "title": title,
        "chartType": _normalize_chart_type(chart_type),
    }


@mcp.tool()
def create_chart_from_spec(
    spreadsheet_id: str,
    sheet: str,
    chart_spec: dict,
    anchor_cell: str = "E2",
    width_pixels: int = 600,
    height_pixels: int = 371,
    offset_x_pixels: int = 0,
    offset_y_pixels: int = 0,
) -> dict:
    """Create an embedded chart from a full Sheets API ChartSpec object.

    Use this for histogram, scorecard, bubble, candlestick, waterfall, treemap,
    org, or other chart specs not covered by ``create_chart``.
    """
    req = {
        "addChart": {
            "chart": {
                "spec": chart_spec,
                "position": _overlay_position(
                    spreadsheet_id,
                    sheet,
                    anchor_cell,
                    width_pixels,
                    height_pixels,
                    offset_x_pixels,
                    offset_y_pixels,
                ),
            }
        }
    }
    result = _execute_batch_update(spreadsheet_id, [req])
    chart = result.get("replies", [{}])[0].get("addChart", {}).get("chart", {})
    return {"chartId": chart.get("chartId"), "sheet": sheet}


@mcp.tool()
def update_chart(spreadsheet_id: str, chart_id: int, chart_spec: dict) -> dict:
    """Replace a chart's spec. Use ``move_chart`` to move or resize it."""
    _execute_batch_update(spreadsheet_id, [{"updateChartSpec": {"chartId": chart_id, "spec": chart_spec}}])
    return {"chartId": chart_id, "updated": True}


@mcp.tool()
def delete_chart(spreadsheet_id: str, chart_id: int) -> dict:
    """Delete a chart or other embedded object by ID."""
    _execute_batch_update(spreadsheet_id, [{"deleteEmbeddedObject": {"objectId": chart_id}}])
    return {"objectId": chart_id, "deleted": True}


@mcp.tool()
def move_chart(
    spreadsheet_id: str,
    chart_id: int,
    sheet: str,
    anchor_cell: str = "E2",
    width_pixels: int = 600,
    height_pixels: int = 371,
    offset_x_pixels: int = 0,
    offset_y_pixels: int = 0,
) -> dict:
    """Move and/or resize a chart using overlay positioning."""
    req = {
        "updateEmbeddedObjectPosition": {
            "objectId": chart_id,
            "newPosition": _overlay_position(
                spreadsheet_id,
                sheet,
                anchor_cell,
                width_pixels,
                height_pixels,
                offset_x_pixels,
                offset_y_pixels,
            ),
            "fields": "anchorCell,offsetXPixels,offsetYPixels,widthPixels,heightPixels",
        }
    }
    _execute_batch_update(spreadsheet_id, [req])
    return {"chartId": chart_id, "sheet": sheet, "anchorCell": anchor_cell}


@mcp.tool()
def set_chart_border(
    spreadsheet_id: str,
    chart_id: int,
    color: str = "#000000",
) -> dict:
    """Set an embedded object's border color."""
    req = {
        "updateEmbeddedObjectBorder": {
            "objectId": chart_id,
            "border": {"color": parse_color(color)},
            "fields": "color",
        }
    }
    _execute_batch_update(spreadsheet_id, [req])
    return {"objectId": chart_id, "borderColor": color}


@mcp.tool()
def create_pivot_table(
    spreadsheet_id: str,
    source_range: str,
    destination_cell: str,
    rows: list[dict] | None = None,
    columns: list[dict] | None = None,
    values: list[dict] | None = None,
    filter_specs: list[dict] | None = None,
    value_layout: str = "HORIZONTAL",
) -> dict:
    """Create a pivot table at a destination cell.

    ``rows``, ``columns``, ``values``, and ``filter_specs`` use the Sheets API
    PivotGroup/PivotValue/PivotFilterSpec shapes.
    """
    pivot_table: dict[str, Any] = {
        "source": current_client().resolve_grid_range(_sid(spreadsheet_id), source_range),
        "valueLayout": value_layout.upper(),
    }
    if rows:
        pivot_table["rows"] = rows
    if columns:
        pivot_table["columns"] = columns
    if values:
        pivot_table["values"] = values
    if filter_specs:
        pivot_table["filterSpecs"] = filter_specs

    req = {
        "updateCells": {
            "start": _single_cell_coordinate(_sid(spreadsheet_id), destination_cell),
            "rows": [{"values": [{"pivotTable": pivot_table}]}],
            "fields": "pivotTable",
        }
    }
    _execute_batch_update(spreadsheet_id, [req])
    return {"sourceRange": source_range, "destinationCell": destination_cell, "valueLayout": value_layout.upper()}


@mcp.tool()
def add_slicer(
    spreadsheet_id: str,
    sheet: str,
    data_range: str,
    column_index: int,
    title: str | None = None,
    filter_criteria: dict | None = None,
    anchor_cell: str = "E2",
    width_pixels: int = 300,
    height_pixels: int = 200,
) -> dict:
    """Add a slicer for a data range."""
    spec: dict[str, Any] = {
        "dataRange": current_client().resolve_grid_range(_sid(spreadsheet_id), data_range, default_sheet=sheet),
        "columnIndex": column_index,
    }
    if title is not None:
        spec["title"] = title
    if filter_criteria is not None:
        spec["filterCriteria"] = filter_criteria
    req = {
        "addSlicer": {
            "slicer": {
                "spec": spec,
                "position": _overlay_position(spreadsheet_id, sheet, anchor_cell, width_pixels, height_pixels),
            }
        }
    }
    result = _execute_batch_update(spreadsheet_id, [req])
    slicer = result.get("replies", [{}])[0].get("addSlicer", {}).get("slicer", {})
    return {"slicerId": slicer.get("slicerId"), "title": title}


@mcp.tool()
def update_slicer(spreadsheet_id: str, slicer_id: int, spec: dict, fields: list[str] | None = None) -> dict:
    """Update a slicer spec by ID."""
    _execute_batch_update(
        spreadsheet_id,
        [{"updateSlicerSpec": {"slicerId": slicer_id, "spec": spec, "fields": _field_mask(fields)}}],
    )
    return {"slicerId": slicer_id, "updatedFields": fields or ["*"]}


@mcp.tool()
def delete_slicer(spreadsheet_id: str, slicer_id: int) -> dict:
    """Delete a slicer by embedded object ID."""
    _execute_batch_update(spreadsheet_id, [{"deleteEmbeddedObject": {"objectId": slicer_id}}])
    return {"slicerId": slicer_id, "deleted": True}


# ---------------------------------------------------------------------------
# Conditional formatting
# ---------------------------------------------------------------------------


@mcp.tool()
def add_conditional_format(
    spreadsheet_id: str,
    range: str,
    condition_type: str,
    values: list[str] | None = None,
    background_color: str | None = None,
    text_color: str | None = None,
    bold: bool | None = None,
) -> dict:
    """Add a boolean conditional-format rule to an A1 range.

    ``condition_type`` is a Sheets ``ConditionType`` such as ``NUMBER_GREATER``,
    ``NUMBER_LESS``, ``NUMBER_BETWEEN``, ``TEXT_CONTAINS``, ``TEXT_EQ``,
    ``CUSTOM_FORMULA``, ``BLANK``, ``NOT_BLANK``. ``values`` supplies the
    comparison operands (e.g. ``["100"]`` or ``["10","20"]``; for
    ``CUSTOM_FORMULA`` a single ``["=$A1>5"]``). The format applies when the
    condition is true.
    """
    grid = current_client().resolve_grid_range(_sid(spreadsheet_id), range)

    fmt: dict[str, Any] = {}
    if background_color is not None:
        fmt["backgroundColor"] = parse_color(background_color)
    text_fmt: dict[str, Any] = {}
    if text_color is not None:
        text_fmt["foregroundColor"] = parse_color(text_color)
    if bold is not None:
        text_fmt["bold"] = bold
    if text_fmt:
        fmt["textFormat"] = text_fmt
    if not fmt:
        raise ValueError("Provide at least one of background_color, text_color, bold.")

    ctype = condition_type.upper()
    if ctype == "CUSTOM_FORMULA" and values:
        for v in values:
            if _NETWORK_FORMULA_RE.search(v):
                raise ValueError(
                    f"CUSTOM_FORMULA value contains a network-fetching function "
                    f"({_NETWORK_FORMULA_RE.search(v).group(1)}), which is not "
                    "allowed. Use a formula that does not make external requests."
                )

    condition: dict[str, Any] = {"type": ctype}
    if values:
        condition["values"] = [{"userEnteredValue": v} for v in values]

    req = {
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [grid],
                "booleanRule": {"condition": condition, "format": fmt},
            },
            "index": 0,
        }
    }
    current_client().batch_update(spreadsheet_id, [req])
    return {"range": range, "conditionType": condition_type.upper()}


@mcp.tool()
def update_conditional_format(
    spreadsheet_id: str,
    sheet: str,
    index: int,
    range: str,
    condition_type: str,
    values: list[str] | None = None,
    background_color: str | None = None,
    text_color: str | None = None,
    bold: bool | None = None,
    new_index: int | None = None,
) -> dict:
    """Replace a conditional-format rule at ``index``."""
    grid = current_client().resolve_grid_range(_sid(spreadsheet_id), range)
    fmt: dict[str, Any] = {}
    if background_color is not None:
        fmt["backgroundColor"] = parse_color(background_color)
    text_fmt: dict[str, Any] = {}
    if text_color is not None:
        text_fmt["foregroundColor"] = parse_color(text_color)
    if bold is not None:
        text_fmt["bold"] = bold
    if text_fmt:
        fmt["textFormat"] = text_fmt
    if not fmt:
        raise ValueError("Provide at least one of background_color, text_color, bold.")

    body: dict[str, Any] = {
        "rule": {
            "ranges": [grid],
            "booleanRule": {"condition": _condition(condition_type, values), "format": fmt},
        },
        "index": index,
        "sheetId": current_client().resolve_sheet_id(_sid(spreadsheet_id), sheet),
    }
    if new_index is not None:
        body["newIndex"] = new_index
    _execute_batch_update(spreadsheet_id, [{"updateConditionalFormatRule": body}])
    return {"range": range, "sheet": sheet, "index": index, "newIndex": new_index, "conditionType": condition_type.upper()}


@mcp.tool()
def delete_conditional_format(spreadsheet_id: str, sheet: str, index: int) -> dict:
    """Delete a conditional-format rule by sheet and index."""
    body: dict[str, Any] = {
        "index": index,
        "sheetId": current_client().resolve_sheet_id(_sid(spreadsheet_id), sheet),
    }
    _execute_batch_update(spreadsheet_id, [{"deleteConditionalFormatRule": body}])
    return {"index": index, "sheet": sheet, "deleted": True}


# Register the natural-interaction layer (composite tools + prompts) on top of
# the low-level tools above. Done here, at the bottom, so these modules can
# import server helpers without a circular import.
from .highlevel import register_highlevel  # noqa: E402
from .prompts import register_prompts  # noqa: E402

register_highlevel(mcp)
register_prompts(mcp)


def main() -> None:
    """Entry point — run the MCP server.

    local mode: stdio. group mode: Streamable HTTP (configured in _build_mcp).
    """
    if _mode() == "group":
        mcp.run(transport="streamable-http")
    else:
        mcp.run()


if __name__ == "__main__":
    main()
