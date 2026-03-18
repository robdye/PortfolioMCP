"""
PortfolioMCP — SharePoint Portfolio Excel MCP Server

A Model Context Protocol (MCP) server built with Python FastMCP that provides
CRUD operations on a portfolio Excel workbook stored in SharePoint.  Uses the
Microsoft Graph Excel API with client credentials (app-only) authentication.

Credentials are read from environment variables so they are never exposed to
callers or hard-coded in source.
"""

import hashlib
import json
import os
import re
import time

import httpx
from mcp.server.fastmcp import FastMCP

AZURE_HOST = os.environ.get(
    "WEBSITE_HOSTNAME", "portfolio-mcp.azurewebsites.net"
)

mcp = FastMCP(
    "PortfolioMCP",
    instructions=(
        "SharePoint portfolio server.  Provides tools to read, add, update, "
        "and remove stock holdings from an Excel workbook on SharePoint."
    ),
    host="0.0.0.0",
    port=int(os.environ.get("PORT", "8000")),
    stateless_http=True,
    transport_security={"enable_dns_rebinding_protection": False},
)

# ---------------------------------------------------------------------------
#  HTTP client
# ---------------------------------------------------------------------------

_http_client: httpx.AsyncClient | None = None


def _get_http_client() -> httpx.AsyncClient:
    """Return a shared HTTP client with connection pooling and timeouts."""
    global _http_client
    if _http_client is None or _http_client.is_closed:
        _http_client = httpx.AsyncClient(
            timeout=httpx.Timeout(30.0, connect=10.0),
        )
    return _http_client


# ---------------------------------------------------------------------------
#  Input validation
# ---------------------------------------------------------------------------

_TICKER_RE = re.compile(r"^[A-Za-z0-9.:\-/]{1,20}$")


def _validate_ticker(symbol: str) -> str:
    """Sanitise and validate a ticker symbol."""
    symbol = symbol.strip().upper()
    if not symbol or not _TICKER_RE.match(symbol):
        raise ValueError(f"Invalid ticker symbol: {symbol!r}")
    return symbol


# ---------------------------------------------------------------------------
#  Microsoft Graph — OAuth2 client credentials
# ---------------------------------------------------------------------------

_graph_token: str | None = None
_graph_token_expiry: float = 0


async def _get_graph_token() -> str:
    """Obtain or reuse a Graph access token via client credentials."""
    global _graph_token, _graph_token_expiry

    if _graph_token and time.time() < _graph_token_expiry - 60:
        return _graph_token

    tenant = os.environ.get("GRAPH_TENANT_ID", "")
    client_id = os.environ.get("GRAPH_CLIENT_ID", "")
    client_secret = os.environ.get("GRAPH_CLIENT_SECRET", "")
    if not all([tenant, client_id, client_secret]):
        raise RuntimeError(
            "Graph credentials not configured.  Set GRAPH_TENANT_ID, "
            "GRAPH_CLIENT_ID, and GRAPH_CLIENT_SECRET."
        )

    client = _get_http_client()
    resp = await client.post(
        f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token",
        data={
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": "https://graph.microsoft.com/.default",
        },
    )
    resp.raise_for_status()
    data = resp.json()
    _graph_token = data["access_token"]
    _graph_token_expiry = time.time() + int(data.get("expires_in", 3600))
    return _graph_token


# ---------------------------------------------------------------------------
#  Microsoft Graph — HTTP helpers
# ---------------------------------------------------------------------------


async def _graph_get(url: str) -> dict:
    """Authenticated GET against Microsoft Graph."""
    token = await _get_graph_token()
    client = _get_http_client()
    resp = await client.get(
        url, headers={"Authorization": f"Bearer {token}"}
    )
    resp.raise_for_status()
    return resp.json()


async def _graph_patch(url: str, body: dict) -> dict:
    """Authenticated PATCH against Microsoft Graph."""
    token = await _get_graph_token()
    client = _get_http_client()
    resp = await client.patch(
        url,
        json=body,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
    )
    resp.raise_for_status()
    return resp.json()


async def _graph_post(url: str, body: dict) -> dict:
    """Authenticated POST against Microsoft Graph."""
    token = await _get_graph_token()
    client = _get_http_client()
    resp = await client.post(
        url,
        json=body,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
    )
    resp.raise_for_status()
    return resp.json()


# ---------------------------------------------------------------------------
#  Workbook URL resolution
# ---------------------------------------------------------------------------


async def _get_workbook_url() -> str:
    """Resolve the Graph workbook URL from site + file path."""
    site = os.environ.get("SHAREPOINT_SITE", "")
    file_path = os.environ.get("PORTFOLIO_FILE_PATH", "")
    if not site or not file_path:
        raise RuntimeError(
            "Set SHAREPOINT_SITE and PORTFOLIO_FILE_PATH env vars."
        )

    token = await _get_graph_token()
    client = _get_http_client()
    headers = {"Authorization": f"Bearer {token}"}

    site_resp = await client.get(
        f"https://graph.microsoft.com/v1.0/sites/{site}:/",
        headers=headers,
    )
    site_resp.raise_for_status()
    site_id = site_resp.json()["id"]

    drives_resp = await client.get(
        f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives",
        headers=headers,
    )
    drives_resp.raise_for_status()
    drives = drives_resp.json().get("value", [])
    if not drives:
        raise RuntimeError("No drives found on the SharePoint site")

    for drive in drives:
        try:
            item_resp = await client.get(
                f"https://graph.microsoft.com/v1.0/drives/{drive['id']}/root:/{file_path}",
                headers=headers,
            )
            if item_resp.status_code == 200:
                item = item_resp.json()
                return f"https://graph.microsoft.com/v1.0/drives/{drive['id']}/items/{item['id']}/workbook"
        except Exception:
            continue

    raise RuntimeError(
        f"File '{file_path}' not found on any drive in site '{site}'"
    )


# ---------------------------------------------------------------------------
#  Row helpers
# ---------------------------------------------------------------------------


def _row_to_holding(row: list) -> dict:
    """Convert an Excel row to a holding dict."""
    def _safe(idx: int) -> str:
        return str(row[idx]) if idx < len(row) and row[idx] is not None else ""

    return {
        "company": _safe(0),
        "ticker": _safe(1),
        "sector": _safe(2),
        "shares": _safe(3),
        "costPerShare": _safe(4),
        "totalCost": _safe(5),
        "website": _safe(6),
        "mediaPressRelease": _safe(7),
        "currentPrice": _safe(8) if len(row) > 8 else "",
    }


# ---------------------------------------------------------------------------
#  Markdown formatting helpers (chat-friendly output)
# ---------------------------------------------------------------------------


def _fmt_usd(val: str) -> str:
    """Format a string value as USD currency."""
    try:
        return f"${float(val):,.2f}"
    except (ValueError, TypeError):
        return "—"


def _fmt_int(val: str) -> str:
    """Format a string value as an integer with commas."""
    try:
        return f"{int(float(val)):,}"
    except (ValueError, TypeError):
        return "—"


def _calc_pl(holding: dict) -> tuple[str, str, bool]:
    """Calculate P/L for a holding. Returns (pl_amount, pl_pct, is_gain)."""
    try:
        cost = float(holding["costPerShare"])
        current = float(holding["currentPrice"])
        shares = int(float(holding["shares"]))
        pl_per_share = current - cost
        pl_total = pl_per_share * shares
        pl_pct = (pl_per_share / cost) * 100 if cost else 0
        return (
            f"${pl_total:+,.2f}",
            f"{pl_pct:+.1f}%",
            pl_total >= 0,
        )
    except (ValueError, TypeError):
        return ("—", "—", True)


def _format_portfolio_md(holdings: list[dict]) -> str:
    """Build a rich, accessible Markdown summary of the portfolio."""
    n = len(holdings)
    lines: list[str] = []

    lines.append("## 📊 Portfolio Summary")
    lines.append(f"**{n} holding{'s' if n != 1 else ''}** across your portfolio\n")

    if n == 0:
        lines.append("_Your portfolio is empty. Use the add_holding tool to get started._")
        return "\n".join(lines)

    # ── KPI summary ──
    total_invested = 0.0
    total_current = 0.0
    gains = 0
    losses = 0
    for h in holdings:
        try:
            tc = float(h["totalCost"])
            cp = float(h["currentPrice"])
            sh = int(float(h["shares"]))
            total_invested += tc
            total_current += cp * sh
            if cp * sh >= tc:
                gains += 1
            else:
                losses += 1
        except (ValueError, TypeError):
            pass

    total_pl = total_current - total_invested
    total_pl_pct = (total_pl / total_invested * 100) if total_invested else 0
    pl_emoji = "📈" if total_pl >= 0 else "📉"

    lines.append(f"| Metric | Value |")
    lines.append(f"|--------|------:|")
    lines.append(f"| Total Invested | {_fmt_usd(str(total_invested))} |")
    lines.append(f"| Current Value | {_fmt_usd(str(total_current))} |")
    lines.append(f"| {pl_emoji} Total P/L | **${total_pl:+,.2f}** ({total_pl_pct:+.1f}%) |")
    lines.append(f"| Positions in Profit | {gains} |")
    lines.append(f"| Positions in Loss | {losses} |")
    lines.append("")

    # ── Group by sector ──
    sectors: dict[str, list[dict]] = {}
    for h in holdings:
        sector = h.get("sector", "").strip() or "Other"
        sectors.setdefault(sector, []).append(h)

    for sector, items in sorted(sectors.items()):
        lines.append(f"### {sector}")
        lines.append("")
        for h in items:
            ticker = h["ticker"]
            company = h["company"]
            shares = _fmt_int(h["shares"])
            cost = _fmt_usd(h["costPerShare"])
            current = _fmt_usd(h["currentPrice"])
            pl_amt, pl_pct, is_gain = _calc_pl(h)
            emoji = "🟢" if is_gain else "🔴"

            lines.append(f"**{company}** (`{ticker}`) — {shares} shares")
            lines.append(f"  Cost: {cost}/share → Current: {current}/share")
            lines.append(f"  {emoji} P/L: **{pl_amt}** ({pl_pct})")
            lines.append("")

    return "\n".join(lines)


def _format_holding_md(h: dict) -> str:
    """Format a single holding as accessible Markdown."""
    ticker = h["ticker"]
    company = h["company"]
    sector = h.get("sector", "") or "—"
    shares = _fmt_int(h["shares"])
    cost = _fmt_usd(h["costPerShare"])
    total = _fmt_usd(h["totalCost"])
    current = _fmt_usd(h["currentPrice"])
    pl_amt, pl_pct, is_gain = _calc_pl(h)
    emoji = "📈" if is_gain else "📉"

    lines = [
        f"## {company} (`{ticker}`)",
        "",
        f"| Detail | Value |",
        f"|--------|------:|",
        f"| Sector | {sector} |",
        f"| Shares | {shares} |",
        f"| Cost per Share | {cost} |",
        f"| Total Cost | {total} |",
        f"| Current Price | {current} |",
        f"| {emoji} P/L | **{pl_amt}** ({pl_pct}) |",
    ]

    website = h.get("website", "")
    if website:
        url = website if website.startswith("http") else f"https://{website}"
        lines.append(f"| Website | [{website}]({url}) |")

    return "\n".join(lines)


def _format_action_md(action: str, h: dict) -> str:
    """Format a CRUD action result as Markdown."""
    ticker = h.get("ticker", "")
    company = h.get("company", "")
    shares = _fmt_int(h.get("shares", "0"))
    cost = _fmt_usd(h.get("costPerShare", "0"))

    if action == "added":
        return (
            f"## ✅ Holding Added\n\n"
            f"**{company}** (`{ticker}`) — {shares} shares at {cost}/share\n\n"
            f"The holding has been added to your SharePoint portfolio."
        )
    elif action == "updated":
        return (
            f"## ✅ Holding Updated\n\n"
            f"**{company}** (`{ticker}`) — {shares} shares at {cost}/share\n\n"
            f"Total Cost: {_fmt_usd(h.get('totalCost', '0'))}"
        )
    elif action == "removed":
        return (
            f"## 🗑️ Holding Removed\n\n"
            f"**{company}** (`{ticker}`) has been removed from your portfolio."
        )
    return ""


# ---------------------------------------------------------------------------
#  MCP Tools — Portfolio CRUD
# ---------------------------------------------------------------------------


@mcp.tool()
async def get_portfolio() -> str:
    """Get all holdings from the SharePoint portfolio Excel file.

    Returns the complete list of stocks in the portfolio with company name,
    ticker, sector, shares held, cost per share, total cost, and current price.
    """
    wb_url = await _get_workbook_url()
    data = await _graph_get(f"{wb_url}/worksheets/Sheet1/usedRange")
    rows = data.get("values", [])
    if len(rows) <= 1:
        return _format_portfolio_md([])

    holdings = [_row_to_holding(r) for r in rows[1:]]
    return _format_portfolio_md(holdings)


@mcp.tool()
async def get_holding(ticker: str) -> str:
    """Get a single holding from the portfolio by ticker symbol.

    Args:
        ticker: Stock ticker symbol (e.g. AAPL, MSFT).
    """
    ticker = ticker.strip().upper()
    wb_url = await _get_workbook_url()
    data = await _graph_get(f"{wb_url}/worksheets/Sheet1/usedRange")
    rows = data.get("values", [])

    for row in rows[1:]:
        if len(row) > 1 and str(row[1]).strip().upper() == ticker:
            return _format_holding_md(_row_to_holding(row))

    return f"⚠️ Ticker **{ticker}** was not found in your portfolio."


@mcp.tool()
async def add_holding(
    company: str,
    ticker: str,
    sector: str = "",
    shares: int = 0,
    cost_per_share: float = 0.0,
    website: str = "",
    media_press_release: str = "",
) -> str:
    """Add a new stock holding to the SharePoint portfolio.

    Args:
        company:             Company name (required).
        ticker:              Stock ticker symbol (required).
        sector:              Industry sector.
        shares:              Number of shares held.
        cost_per_share:      Purchase price per share in USD.
        website:             Company website URL.
        media_press_release: Media/press release URL.
    """
    ticker = _validate_ticker(ticker)
    total_cost = shares * cost_per_share

    wb_url = await _get_workbook_url()
    data = await _graph_get(f"{wb_url}/worksheets/Sheet1/usedRange")
    rows = data.get("values", [])
    next_row = len(rows) + 1

    for row in rows[1:]:
        if len(row) > 1 and str(row[1]).strip().upper() == ticker:
            return f"⚠️ Ticker **{ticker}** already exists in your portfolio. Use update_holding to modify it."

    new_row = [
        company.strip(), ticker, sector.strip(), shares,
        cost_per_share, total_cost,
        website.strip(), media_press_release.strip(), "",
    ]

    cell_range = f"A{next_row}:I{next_row}"
    await _graph_patch(
        f"{wb_url}/worksheets/Sheet1/range(address='{cell_range}')",
        {"values": [new_row]},
    )

    return _format_action_md("added", _row_to_holding(new_row))


@mcp.tool()
async def update_holding(
    ticker: str,
    shares: int | None = None,
    cost_per_share: float | None = None,
    sector: str | None = None,
    website: str | None = None,
    media_press_release: str | None = None,
) -> str:
    """Update an existing holding in the portfolio.

    Only non-null fields are updated.  The total cost is recalculated
    automatically when shares or cost_per_share change.

    Args:
        ticker:              Ticker of the holding to update (required).
        shares:              New number of shares.
        cost_per_share:      New cost per share in USD.
        sector:              New sector.
        website:             New website URL.
        media_press_release: New media/press release URL.
    """
    ticker = _validate_ticker(ticker)
    wb_url = await _get_workbook_url()
    data = await _graph_get(f"{wb_url}/worksheets/Sheet1/usedRange")
    rows = data.get("values", [])

    for idx, row in enumerate(rows):
        if idx == 0:
            continue
        if len(row) > 1 and str(row[1]).strip().upper() == ticker:
            if shares is not None:
                row[3] = shares
            if cost_per_share is not None:
                row[4] = cost_per_share
            if sector is not None:
                row[2] = sector.strip()
            if website is not None:
                row[6] = website.strip()
            if media_press_release is not None:
                row[7] = media_press_release.strip()

            try:
                row[5] = float(row[3]) * float(row[4])
            except (ValueError, TypeError):
                pass

            excel_row = idx + 1
            await _graph_patch(
                f"{wb_url}/worksheets/Sheet1/range(address='A{excel_row}:I{excel_row}')",
                {"values": [row[:9]]},
            )
            return _format_action_md("updated", _row_to_holding(row))

    return f"⚠️ Ticker **{ticker}** was not found in your portfolio."


@mcp.tool()
async def remove_holding(ticker: str) -> str:
    """Remove a stock holding from the portfolio by ticker.

    Args:
        ticker: Ticker of the holding to remove.
    """
    ticker = _validate_ticker(ticker)
    wb_url = await _get_workbook_url()
    data = await _graph_get(f"{wb_url}/worksheets/Sheet1/usedRange")
    rows = data.get("values", [])

    for idx, row in enumerate(rows):
        if idx == 0:
            continue
        if len(row) > 1 and str(row[1]).strip().upper() == ticker:
            excel_row = idx + 1
            token = await _get_graph_token()
            client = _get_http_client()
            resp = await client.post(
                f"{wb_url}/worksheets/Sheet1/range(address='A{excel_row}:I{excel_row}')/clear",
                json={"applyTo": "Contents"},
                headers={
                    "Authorization": f"Bearer {token}",
                    "Content-Type": "application/json",
                },
            )
            resp.raise_for_status()
            return _format_action_md("removed", {"ticker": ticker, "company": str(row[0]) if row else ""})

    return f"⚠️ Ticker **{ticker}** was not found in your portfolio."


# ---------------------------------------------------------------------------
#  MCP Tools — Column Management
# ---------------------------------------------------------------------------

_COLUMN_NAME_RE = re.compile(r"^[A-Za-z0-9 _\-/.&()]{1,50}$")


def _validate_column_name(name: str) -> str:
    """Sanitise and validate a column name."""
    name = name.strip()
    if not name or not _COLUMN_NAME_RE.match(name):
        raise ValueError(f"Invalid column name: {name!r}")
    return name


@mcp.tool()
async def add_column(name: str) -> str:
    """Add a new column to the portfolio table.

    Adds a column header to the right of the last existing column.
    The column is added to Table1 in the portfolio workbook.

    Args:
        name: Name for the new column header (e.g. 'Current Price', 'P/L %').
    """
    name = _validate_column_name(name)
    wb_url = await _get_workbook_url()

    # Get current table columns to find the next position
    table_data = await _graph_get(f"{wb_url}/tables/Table1/columns")
    existing_cols = table_data.get("value", [])
    col_names = [c.get("name", "") for c in existing_cols]

    if name in col_names:
        return f"⚠️ Column **{name}** already exists in the table."

    # Add column via the Table1 columns endpoint
    result = await _graph_post(
        f"{wb_url}/tables/Table1/columns",
        {"name": name, "values": [[name]]},
    )

    col_index = result.get("index", len(existing_cols))
    return (
        f"✅ Column **{name}** added to the portfolio table.\n\n"
        f"Column index: {col_index} (0-based)\n"
        f"Total columns: {len(col_names) + 1}\n\n"
        f"Use `update_column` to populate the column with values."
    )


@mcp.tool()
async def update_column(col_index: int, values: list) -> str:
    """Update all values in a portfolio table column.

    Writes a complete set of values for a column identified by its
    zero-based index.  The values array should include the header as
    the first element, followed by one value per data row.

    Example: update_column(8, [["Current Price"],["131.99"],["6.36"]])

    Args:
        col_index: Zero-based column index in Table1.
        values:    2D array of values — [[header],[row1],[row2],...].
    """
    if col_index < 0:
        return "⚠️ Column index must be >= 0."

    wb_url = await _get_workbook_url()

    # Validate column exists
    table_data = await _graph_get(f"{wb_url}/tables/Table1/columns")
    existing_cols = table_data.get("value", [])
    if col_index >= len(existing_cols):
        return (
            f"⚠️ Column index {col_index} is out of range. "
            f"Table has {len(existing_cols)} columns (0–{len(existing_cols) - 1})."
        )

    col_name = existing_cols[col_index].get("name", f"Column {col_index}")

    # Write values to the column
    # The Graph API expects the full column range including header
    await _graph_patch(
        f"{wb_url}/tables/Table1/columns/{col_index}/range",
        {"values": values},
    )

    data_rows = len(values) - 1 if len(values) > 1 else 0
    return (
        f"✅ Column **{col_name}** (index {col_index}) updated.\n\n"
        f"Wrote {data_rows} data value{'s' if data_rows != 1 else ''} "
        f"(plus header).\n"
    )


@mcp.tool()
async def list_columns() -> str:
    """List all columns in the portfolio table with their indices.

    Returns the column names and zero-based indices from Table1.
    """
    wb_url = await _get_workbook_url()
    table_data = await _graph_get(f"{wb_url}/tables/Table1/columns")
    existing_cols = table_data.get("value", [])

    if not existing_cols:
        return "The portfolio table has no columns."

    lines = [
        "## Portfolio Table Columns\n",
        f"**{len(existing_cols)} columns** in Table1\n",
        "| Index | Column Name |",
        "|------:|-------------|",
    ]
    for col in existing_cols:
        idx = col.get("index", "?")
        name = col.get("name", "?")
        lines.append(f"| {idx} | {name} |")

    return "\n".join(lines)


@mcp.tool()
async def rename_column(col_index: int, new_name: str) -> str:
    """Rename a column in the portfolio table.

    Args:
        col_index: Zero-based column index to rename.
        new_name:  New name for the column header.
    """
    new_name = _validate_column_name(new_name)
    wb_url = await _get_workbook_url()

    table_data = await _graph_get(f"{wb_url}/tables/Table1/columns")
    existing_cols = table_data.get("value", [])
    if col_index < 0 or col_index >= len(existing_cols):
        return (
            f"⚠️ Column index {col_index} is out of range. "
            f"Table has {len(existing_cols)} columns (0–{len(existing_cols) - 1})."
        )

    old_name = existing_cols[col_index].get("name", "")

    # Rename by updating the header cell
    # Table columns are named by their header row
    # Get header range for this column
    header_data = await _graph_get(
        f"{wb_url}/tables/Table1/columns/{col_index}/headerRowRange"
    )
    await _graph_patch(
        f"{wb_url}/tables/Table1/columns/{col_index}/headerRowRange",
        {"values": [[new_name]]},
    )

    return (
        f"✅ Column renamed from **{old_name}** to **{new_name}** "
        f"(index {col_index}).\n"
    )


# ---------------------------------------------------------------------------
#  CORS
# ---------------------------------------------------------------------------


def _compute_widget_renderer_origin(server_domain: str) -> str:
    domain_hash = hashlib.sha256(server_domain.encode()).hexdigest()
    return f"https://{domain_hash}.widget-renderer.usercontent.microsoft.com"


def _build_allowed_origins() -> list[str]:
    origins: list[str] = []
    explicit = os.environ.get("WIDGET_RENDERER_ORIGIN", "").strip()
    if explicit:
        origins.append(explicit)
    origins.append(_compute_widget_renderer_origin(AZURE_HOST))
    origins.append(f"https://{AZURE_HOST}")
    origins.append("http://localhost:5173")
    origins.append("http://localhost:8000")
    return origins


# ---------------------------------------------------------------------------
#  Entrypoint
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import uvicorn
    from starlette.middleware.cors import CORSMiddleware
    from starlette.responses import JSONResponse
    from starlette.routing import Route

    _allowed_origins = _build_allowed_origins()

    async def health(request):
        return JSONResponse({"status": "healthy", "server": "PortfolioMCP"})

    app = mcp.streamable_http_app()
    app.routes.insert(0, Route("/health", health))

    app.add_middleware(
        CORSMiddleware,
        allow_origins=_allowed_origins,
        allow_methods=["GET", "POST", "OPTIONS"],
        allow_headers=["Content-Type", "Accept", "mcp-session-id"],
        expose_headers=["mcp-session-id"],
        max_age=600,
    )

    port = int(os.environ.get("PORT", "8000"))
    print(f"CORS allowed origins: {_allowed_origins}")
    print(f"MCP endpoint: http://0.0.0.0:{port}/mcp")

    uvicorn.run(app, host="0.0.0.0", port=port)
