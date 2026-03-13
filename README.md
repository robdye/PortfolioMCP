# PortfolioMCP

A Model Context Protocol (MCP) server that provides CRUD operations on a portfolio Excel workbook stored in SharePoint. Uses the Microsoft Graph Excel API with client credentials (app-only) authentication.

## Tools

| Tool | Description |
|------|-------------|
| **get_portfolio** | Get all holdings from the SharePoint portfolio Excel file |
| **get_holding** | Get a single holding by ticker symbol |
| **add_holding** | Add a new stock holding to the portfolio |
| **update_holding** | Update shares, cost, sector for an existing holding |
| **remove_holding** | Remove a holding by ticker |

## Prerequisites

- Python 3.13+
- Microsoft Entra app registration with `Files.ReadWrite.All` application permission
- SharePoint site with a portfolio Excel workbook

## Environment Variables

```
GRAPH_CLIENT_ID=<Entra app client ID>
GRAPH_CLIENT_SECRET=<Client secret>
GRAPH_TENANT_ID=<Entra tenant ID>
SHAREPOINT_SITE=<SharePoint hostname, e.g. contoso.sharepoint.com>
PORTFOLIO_FILE_PATH=<File path in document library, e.g. AlphaAnalyzer-Portfolio.xlsx>
```

## Run Locally

```bash
pip install -r requirements.txt
python server.py
```

Server starts at `http://localhost:8000`. MCP endpoint at `http://localhost:8000/mcp`.

## Deploy to Azure

```powershell
az webapp deploy --name portfolio-mcp --resource-group rg-portfolio-mcp --src-path publish.zip --type zip
```

## Architecture

This server is designed to be used as a separate MCP action within a declarative agent alongside AlphaAnalyzerMCP (Finnhub market data). The two servers are independent — this one handles SharePoint Excel CRUD, the other handles real-time market data.
