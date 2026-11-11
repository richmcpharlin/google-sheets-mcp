# google-sheets-mcp

An [MCP](https://modelcontextprotocol.io) server that gives an LLM **cell-level**
and **formatting-level** control over Google Sheets via the Sheets API v4. Two
deployment modes:

- **local** (default) — runs over stdio for one user; authenticates with your own
  Google account via a desktop OAuth flow and caches the token on disk.
- **group** — centrally hosted over HTTP for many users; each connects from Claude
  via the native **"Connect"** button and acts as their **own** Google identity.
  The server is **stateless** — it persists no per-user tokens (see
  [Group mode](#group-hosted-mode)).

## Capabilities

### Natural interactions (start here)

A high-level layer sits on top of the ~75 low-level API wrappers so common
requests map to a single call phrased the way a user thinks. These accept a
spreadsheet **ID or a full URL**, and natural range references (a header name
like `Revenue`, `last row`, `top 10 rows`, `whole sheet`).

- `describe_spreadsheet` — one-call snapshot: sheets, dimensions, column headers
  with inferred types, sample rows, tables/charts. **Call this first** to ground
  a request before acting.
- `build_table` — write data + style the header + freeze + banding + auto-fit +
  per-column number formats in one shot (`preview=True` to see the plan first).
- `apply_style_preset` — `clean` / `financial` / `report-header` / `input`.
- `highlight_where` — conditional highlighting from a plain predicate:
  `"> 100"`, `"between 10 and 20"`, `"contains overdue"`, `"blank"`,
  `"duplicates"`, `"top 10%"`.
- `format_numbers` — `currency` / `percent` / `date` / `thousands` / `plain`.
- `add_totals_row`, `autofit`, `sort_by` (by header name).

**Prompts** (in the client's slash/prompt menu): `format_as_report`,
`clean_sheet`, `build_dashboard`, `analyze` — natural-language workflows that
orchestrate the tools.

The low-level wrappers below remain available for precise control.

### Low-level tools

**Structure**
- `get_spreadsheet_info` — list sheets with ids, dimensions, frozen rows/cols
- `create_spreadsheet`, `add_sheet`, `rename_sheet`, `delete_sheet`,
  `duplicate_sheet`

**Cell values**
- `read_range`, `batch_read` — read A1 ranges (formatted, unformatted, or formulas)
- `write_range`, `batch_write`, `append_rows`, `clear_range`, `batch_clear`

**Data operations**
- `insert_rows`, `delete_rows`, `insert_columns`, `delete_columns`
- `append_rows_to_sheet`, `append_columns_to_sheet`, `move_rows_or_columns`
- `insert_range`, `delete_range`, `copy_range`, `cut_paste_range`,
  `paste_delimited_data`, `auto_fill`
- `sort_range`, `find_replace`, `trim_whitespace`, `delete_duplicates`,
  `text_to_columns`, `randomize_range`

**Formatting**
- `read_formatting` — read displayed colors, emphasis, number formats, and notes
  for a bounded range, collapsed into rectangular runs
- `format_cells` — bold/italic/underline/strikethrough, font size & family,
  text & background color, horizontal/vertical alignment, wrap strategy, number formats
- `set_borders` — outer + inner borders, styles, colors
- `merge_cells`, `unmerge_cells`
- `set_dimension_size` (column width / row height), `auto_resize_dimensions`
- `freeze` rows/columns
- `add_conditional_format`, `update_conditional_format`,
  `delete_conditional_format`
- `add_banding`, `update_banding`, `delete_banding`

**Filters, validation, and ranges**
- `set_basic_filter`, `clear_basic_filter`
- `add_filter_view`, `update_filter_view`, `delete_filter_view`
- `set_data_validation`, `clear_data_validation`
- `add_named_range`, `update_named_range`, `delete_named_range`
- `add_protected_range`, `update_protected_range`, `delete_protected_range`
- `add_dimension_group`, `update_dimension_group`, `delete_dimension_group`
- `add_table`, `update_table`, `delete_table`

**Charts**
- `create_chart` — create embedded column, bar, line, area, scatter, combo,
  stepped-area, pie, and donut charts from existing row/column ranges
- `create_chart_from_spec` — create any Sheets API chart spec, including
  histogram, scorecard, bubble, candlestick, waterfall, treemap, and org charts
- `update_chart`, `delete_chart`, `move_chart`, `set_chart_border`
- `create_pivot_table`
- `add_slicer`, `update_slicer`, `delete_slicer`

**Advanced**
- `batch_update_advanced` — allowlisted raw Sheets API `batchUpdate` requests
  (max 50 per call; data-source lifecycle requests are blocked)

## Setup (local mode)

### 1. Get an OAuth client secret

1. In the [Google Cloud Console](https://console.cloud.google.com/), create (or pick)
   a project and **enable the Google Sheets API** (APIs & Services → Library).
2. Configure the OAuth consent screen (External is fine for personal use; add your
   own Google account as a *Test user*).
3. APIs & Services → Credentials → **Create Credentials → OAuth client ID** →
   application type **Desktop app**. Download the JSON.
4. Save it as `credentials.json` under the config directory:

   ```
   ~/.config/google-sheets-mcp/credentials.json
   ```

   (Or set `$GOOGLE_SHEETS_CREDENTIALS` to its path.)

### 2. Install

```bash
uv sync
```

### 3. Authenticate once

```bash
uv run google-sheets-mcp auth
```

This opens a browser, you log in, and the token is cached at
`~/.config/google-sheets-mcp/token.json`. After this the server runs without
prompting (the token auto-refreshes).

## Running

The server speaks MCP over **stdio** — point your MCP client at it.

### Claude Code

```bash
claude mcp add google-sheets -- uv run --directory /path/to/google-sheets-mcp google-sheets-mcp
```

### Claude Desktop (`claude_desktop_config.json`)

```json
{
  "mcpServers": {
    "google-sheets": {
      "command": "uv",
      "args": ["run", "--directory", "/path/to/google-sheets-mcp", "google-sheets-mcp"]
    }
  }
}
```

## Group (hosted) mode

One server, many users. Each user clicks **Connect** in Claude, consents with
their own Google account, and from then on every tool call runs as **that** user
— so Google's own sharing permissions and audit trail apply per person.

### How auth works (and why nothing is stored)

The server is both an OAuth **Authorization Server** to Claude and an OAuth
**client** to Google (a "bridge"). When a user connects, the flow is:

```
Claude  --register/authorize-->  this server  --redirect-->  Google consent
Google  --code-->  /oauth/google/callback  (exchange for Google tokens)
this server  --issues-->  MCP access + refresh tokens  -->  Claude stores them
Claude  --Bearer token-->  tool calls  -->  per-request Google client
```

Every MCP token the server issues is a **self-contained, encrypted blob** that
*wraps* the user's Google token. The user's Google **refresh token lives inside
the MCP refresh token that Claude stores client-side** — the server keeps **no
per-user token database**. The only durable secrets are app-level: the Google
client secret and a symmetric **wrap key**. This is the secure-but-good-UX
middle ground: the same one-click Connect experience as a first-party connector,
without a central store of user credentials.

### Setup

1. **Google Cloud:** enable the Sheets API; configure the OAuth consent screen;
   create an OAuth **Web application** client. Add this authorized redirect URI:

   ```
   https://YOUR_HOST/oauth/google/callback
   ```

   Download the client JSON (or note the client id/secret).

2. **Generate a wrap key:**

   ```bash
   uv run google-sheets-mcp genkey
   ```

3. **Configure and run** (put secrets in a secret manager / env, not in files):

   ```bash
   export GSHEETS_MCP_MODE=group
   export GSHEETS_PUBLIC_URL=https://YOUR_HOST          # public https base URL
   export GSHEETS_GOOGLE_CLIENT_SECRETS=/path/web_client.json   # or *_CLIENT_ID/_SECRET
   export GSHEETS_FERNET_KEY=<key from genkey>          # or GSHEETS_FERNET_KEYS=new,old
   export GSHEETS_HOST=0.0.0.0 GSHEETS_PORT=8000
   uv run google-sheets-mcp serve --mode group
   ```

   Terminate **TLS** in front of it (nginx/caddy/cloud LB) so `GSHEETS_PUBLIC_URL`
   is `https://`. Then add `https://YOUR_HOST/mcp` as a custom/remote connector in
   Claude and click **Connect**.

### Security notes

- **TLS is mandatory** — bearer tokens travel in request headers.
- **The wrap key is the crown jewel.** Keep it in a KMS/secret manager. To rotate
  with zero downtime, set `GSHEETS_FERNET_KEYS=<new>,<old>` (encrypt with the new
  key, still accept the old) until old tokens expire, then drop the old key.
- **Least privilege:** the server requests only the `spreadsheets` scope (plus
  `openid`/`email` for per-user identity in logs).
- **Revocation:** `/revoke` revokes the user's Google grant. Because issued access
  tokens are self-contained, they can't be individually revoked before expiry
  without a denylist — mitigated by short (≈1h) access-token lifetimes.
- **Scaling caveat:** dynamic client registrations are held in memory (public
  redirect URIs only, never tokens). For multi-instance deployments, run behind a
  sticky-session LB or add a shared client store; tokens themselves need no shared
  state.

## Notes on ranges

- **A1 notation** (`Sheet1!A1:C10`, `A:C`, `2:5`, `B2`) is used by all
  value/formatting tools. If you omit the `Sheet!` prefix the first sheet is used.
- **Dimension tools** (`set_dimension_size`, `auto_resize_dimensions`) use
  zero-based, half-open indices: `start_index=0, end_index=3` = first three.
- Colors accept `#RRGGBB`, short `#RGB`, named colors (`red`, `lightgray`, …),
  or a `{"red":..,"green":..,"blue":..}` dict (0.0–1.0).

## Config / environment variables

| Variable | Default | Purpose |
| --- | --- | --- |
| `GOOGLE_SHEETS_MCP_HOME` | `~/.config/google-sheets-mcp` | Config directory (local) |
| `GOOGLE_SHEETS_CREDENTIALS` | `<home>/credentials.json` | OAuth client secret (local) |
| `GOOGLE_SHEETS_TOKEN` | `<home>/token.json` | Cached user token (local) |
| `GSHEETS_MCP_MODE` | `local` | `local` or `group` |
| `GSHEETS_PUBLIC_URL` | — | Public https base URL (group, required) |
| `GSHEETS_GOOGLE_CLIENT_SECRETS` | — | Path to Google **Web** client JSON (group) |
| `GSHEETS_GOOGLE_CLIENT_ID` / `_SECRET` | — | Google client creds if not using the JSON (group) |
| `GSHEETS_FERNET_KEY` / `GSHEETS_FERNET_KEYS` | — | Token wrap key(s); first is primary (group, required) |
| `GSHEETS_HOST` / `GSHEETS_PORT` | `0.0.0.0` / `8000` | Bind address (group) |

## Development

```bash
uv run pytest
```

Secrets (`credentials.json`, `token.json`) are git-ignored — never commit them.
