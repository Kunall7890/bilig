# FastMCP WorkPaper client

Use this example when a Python agent, test harness, or MCP client wants to call
Bilig workbook tools through FastMCP instead of driving Excel, LibreOffice, or a
browser spreadsheet grid.

FastMCP owns the MCP client session. Bilig owns the WorkPaper formula tools:
read sheets, inspect cells, write one input, verify recalculated readback, and
export the WorkPaper JSON boundary.

## Run

Hosted smoke test:

```sh
uv run --python 3.12 --with 'fastmcp-slim[client]' \
  python fastmcp_workpaper_client.py --output .tmp/fastmcp-workpaper-proof.json
```

Expected output:

```json
{
  "client": "fastmcp",
  "transport": "streamable-http",
  "endpoint": "https://bilig.proompteng.ai/mcp",
  "verified": true
}
```

The full proof file includes the discovered tool names, sheet names,
`Summary!B2` and `Summary!B3` formula readback, the `Inputs!B3` edit proof, and
the exported WorkPaper document size.

Local file-backed check:

```sh
uv run --python 3.12 --with 'fastmcp-slim[client]' \
  python fastmcp_workpaper_client.py \
    --transport stdio \
    --workpaper .tmp/fastmcp-pricing.workpaper.json \
    --output .tmp/fastmcp-workpaper-stdio.json
```

This launches `@bilig/workpaper` with FastMCP's stdio client config, writes
`Inputs!B3 = 0.4`, verifies `Summary!B2 = 8` and `Summary!B3 = 96000`, closes
the client, reopens the same WorkPaper JSON file, and verifies the edit is still
there.

The same server config is available in `mcp.json` for agents and MCP clients
that read standard `mcpServers` configuration.

## Hosted Endpoint Boundary

The default endpoint is intentionally stateless. It is useful for tool
discovery, smoke tests, and docs examples because it does not write user files.
The `set_cell_contents` call still returns request-local proof:

- `editedCell`
- `before`
- `after`
- `restored`
- `checks.restoredMatchesAfter`
- `persistence.serializedBytes`

Do not expect a later `read_cell` call against the hosted endpoint to observe a
previous call's edit. Use the local file-backed stdio server for private or
persisted WorkPaper files:

```sh
npm exec --package @bilig/workpaper@latest -- \
  bilig-workpaper-mcp \
  --workpaper ./pricing.workpaper.json \
  --init-demo-workpaper \
  --writable
```

## Why This Exists

FastMCP's Python client is a good deterministic harness for MCP servers. Bilig's
MCP server gives workbook-shaped business logic a tool boundary that agents can
verify without screenshot automation.

Use this when your Python agent stack needs:

- formula-backed quote, payout, budget, or import-validation logic;
- exact sheet/cell addresses rather than screenshots;
- a proof object that includes write, readback, and export evidence;
- a smoke test for a hosted Streamable HTTP MCP server.

Official FastMCP references:

- <https://gofastmcp.com/clients/client>
- <https://gofastmcp.com/clients/transports>
- <https://gofastmcp.com/community/showcase>
