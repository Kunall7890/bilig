# LangGraph.js WorkPaper Tool State

This example runs a real `@langchain/langgraph` `ToolNode` against a Bilig
WorkPaper. It is for agents that need spreadsheet-style business logic inside
graph state without driving Excel, a browser, or cached XLSX formula values.

LangGraph owns graph execution and tool-message state. Bilig owns workbook
mutation, formula recalculation, JSON persistence, restore, and readback proof.

Official LangGraph.js references:

- https://docs.langchain.com/oss/javascript/langchain/tools
- https://langchain-ai.github.io/langgraphjs/reference/classes/langgraph.prebuilt.ToolNode.html
- https://langchain-ai.github.io/langgraphjs/reference/functions/langgraph.prebuilt.toolsCondition.html

## Local Smoke

```sh
pnpm install --ignore-workspace --lockfile=false
pnpm run typecheck
pnpm run smoke
```

The smoke invokes a graph with an `AIMessage` containing two tool calls:

1. `read_workpaper_quote` reads the formula-backed quote summary.
2. `set_workpaper_quantity` edits `Inputs!B2`, recalculates formulas, persists
   WorkPaper JSON, restores it, and returns proof that restored formulas match.

Expected proof shape:

```json
{
  "framework": "langgraphjs-toolnode",
  "toolMessageNames": ["read_workpaper_quote", "set_workpaper_quantity"],
  "proof": {
    "editedCell": "Inputs!B2",
    "before": { "total": 1458 },
    "after": { "total": 2187 },
    "afterRestore": { "total": 2187 },
    "verified": true
  }
}
```

Keep the proof with the graph run, trace, checkpoint metadata, or app audit log
instead of logging only "tool succeeded".

## Boundaries

This is a local ToolNode proof, not an official LangGraph template and not a
desktop Excel oracle. Keep Excel, LibreOffice, Microsoft Graph, or a domain
oracle in the loop for macros, pivots, external links, volatile functions, and
workbooks outside Bilig's formula surface.
