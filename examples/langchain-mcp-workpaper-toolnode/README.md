# LangChain MCP WorkPaper ToolNode

This example connects LangChain.js to the Bilig WorkPaper MCP server and runs
the returned MCP tools through a real LangGraph.js `ToolNode`.

It is for agent stacks that already use LangChain or LangGraph and need a
spreadsheet tool that can prove the write, formula readback, JSON persistence,
and restored value without opening Excel or a browser UI.

Official references:

- https://docs.langchain.com/oss/javascript/langchain/mcp
- https://docs.langchain.com/oss/javascript/langchain/tools
- https://reference.langchain.com/javascript/classes/_langchain_langgraph.prebuilt.ToolNode.html

## Local Smoke

```sh
pnpm install --ignore-workspace --lockfile=false
pnpm run typecheck
pnpm run smoke
```

The smoke starts:

```sh
npm exec --yes --package @bilig/workpaper@latest -- \
  bilig-workpaper-mcp \
  --workpaper .tmp/pricing.workpaper.json \
  --init-demo-workpaper \
  --writable
```

Then it uses `@langchain/mcp-adapters` to load the MCP tools and `ToolNode` to
call:

1. `read_cell` for `Summary!B3`.
2. `set_cell_contents` to edit `Inputs!B3` and persist the WorkPaper JSON.
3. `read_cell` again for `Summary!B3`.
4. `get_cell_display_value` for display readback.
5. `export_workpaper_document` for serialized WorkPaper JSON proof.
6. A second read-only MCP client restart against the same file to prove the
   persisted formula result survived process boundaries.

Expected proof shape:

```json
{
  "framework": "langchainjs-mcp-adapters-toolnode",
  "mcpTransport": "stdio",
  "editedCell": "Inputs!B3",
  "dependentCell": "Summary!B3",
  "before": 60000,
  "after": 96000,
  "afterRestart": 96000,
  "displayValue": "96000",
  "checks": {
    "dependentCellChanged": true,
    "persistedToDisk": true,
    "restartReadbackMatchesAfter": true,
    "exportedWorkPaperDocument": true
  },
  "verified": true
}
```

## Boundaries

This proves the LangChain MCP adapter path into Bilig WorkPaper tools. It does
not prove arbitrary XLSX compatibility, macros, pivots, charts, external links,
unsupported formulas, or desktop Excel parity.
