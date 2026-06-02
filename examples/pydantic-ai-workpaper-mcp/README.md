# Pydantic AI WorkPaper MCP

Use this example when a Pydantic AI app needs workbook formulas through MCP and
wants typed proof before an agent trusts the result.

The default smoke test is keyless. It launches Bilig's published WorkPaper MCP
server over stdio, edits `Inputs!B3`, reads the calculated `Summary!B3` value,
persists the WorkPaper JSON, restores it, and validates the result with a
Pydantic model.

```sh
uv run --python 3.12 --with pydantic-ai --with mcp --with fastmcp \
  python pydantic_ai_workpaper_mcp.py --output .tmp/pydantic-ai-workpaper-proof.json
```

Expected result:

```json
{
  "framework": "pydantic-ai",
  "toolset": "MCPToolset",
  "packageSpec": "@bilig/workpaper@latest",
  "verified": true,
  "beforeExpectedArr": 60000,
  "afterExpectedArr": 96000
}
```

The script uses Pydantic AI's modern `MCPToolset` with `fastmcp`'s
`StdioTransport`. It does not use the deprecated `MCPServerStdio` path.
