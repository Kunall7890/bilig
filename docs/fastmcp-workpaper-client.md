---
title: FastMCP client for WorkPaper formula tools
description: Use FastMCP's Python client to call Bilig WorkPaper MCP tools over Streamable HTTP and prove formula readback without spreadsheet UI automation.
published: true
---

# FastMCP Client for WorkPaper Formula Tools

Use this when a Python agent stack already uses FastMCP and needs workbook
formula tools without opening Excel, LibreOffice, Google Sheets, or browser UI
automation.

FastMCP owns the MCP client session. Bilig owns the WorkPaper formula tools:
read sheets, read cells, write one input, verify recalculated readback, and
export a WorkPaper JSON boundary.

## Run The Client

```sh
cd examples/fastmcp-workpaper-client
uv run --python 3.12 --with 'fastmcp-slim[client]' \
  python fastmcp_workpaper_client.py --output .tmp/fastmcp-workpaper-proof.json
```

The script connects to:

```text
https://bilig.proompteng.ai/mcp
```

The client code uses FastMCP directly:

```python
from fastmcp import Client
```

It verifies the hosted Bilig MCP endpoint exposes:

- `list_sheets`
- `read_range`
- `read_cell`
- `set_cell_contents`
- `get_cell_display_value`
- `export_workpaper_document`
- `validate_formula`

Then it reads `Summary!B2` and `Summary!B3`, writes `Inputs!B3 = 0.4`, checks
the request-local restore proof, exports the WorkPaper document, and prints a
compact JSON proof with `verified: true`.

## Proof Shape

```json
{
  "client": "fastmcp",
  "transport": "streamable-http",
  "endpoint": "https://bilig.proompteng.ai/mcp",
  "readback": {
    "expectedCustomersCell": "Summary!B2",
    "expectedCustomersFormula": "=Inputs!B2*Inputs!B3",
    "expectedArrCell": "Summary!B3",
    "expectedArrFormula": "=B2*Inputs!B4",
    "editedCell": "Inputs!B3",
    "previousSerialized": 0.25,
    "newSerialized": 0.4,
    "restoredMatchesAfter": true,
    "persisted": false
  },
  "verified": true
}
```

## Hosted Endpoint Boundary

The hosted endpoint is stateless. It is useful for FastMCP tool discovery,
smoke tests, and docs examples because it does not write user files.

Do not expect a later `read_cell` call against the hosted endpoint to observe a
previous call's edit. The `set_cell_contents` call returns request-local proof
for that call. Use the local file-backed stdio server when the workflow must
persist a private WorkPaper JSON file:

```sh
npm exec --package @bilig/workpaper@latest -- \
  bilig-workpaper-mcp \
  --workpaper ./pricing.workpaper.json \
  --init-demo-workpaper \
  --writable
```

## Why This Fits FastMCP

FastMCP's `Client` gives Python code a deterministic way to connect to local
or remote MCP servers, list tools, and call tools with structured results.
Bilig supplies the workbook formula side of that boundary.

This is a good fit for:

- Python agent tests that need a known MCP server;
- workflow agents that need exact sheet/cell addresses;
- formula-backed quote, payout, budget, or import-validation checks;
- a clean smoke test before wiring a private file-backed WorkPaper server.

Official FastMCP references:

- FastMCP client docs: <https://gofastmcp.com/clients/client>
- FastMCP community showcase: <https://gofastmcp.com/community/showcase>
