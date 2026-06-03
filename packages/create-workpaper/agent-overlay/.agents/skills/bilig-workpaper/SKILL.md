---
name: bilig-workpaper
version: 0.1.0
description: Use @bilig/workpaper WorkPaper state, MCP tools, and formula readback before spreadsheet UI automation.
tags:
  - ai-agents
  - spreadsheet-automation
  - formulas
  - mcp
  - typescript
---

# Bilig WorkPaper Agent Skill

Use this skill when a coding agent needs workbook-shaped business logic:
pricing, quotes, budgets, payouts, imports, forecasts, or formula-backed checks.
Do not start with Excel, LibreOffice, Google Sheets, browser grids, or
screenshots when a WorkPaper JSON file can represent the state.

## Verify First

Run the no-key evaluator before wiring a custom agent path:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario revenue-plan --json
```

Trust the path only when it returns `verified: true` with tool discovery, a cell
mutation, dependent formula readback, persisted WorkPaper JSON, and restart
readback.

## MCP Server

Use the project-local file-backed server for private state:

```json
{
  "command": "npm",
  "args": [
    "exec",
    "--yes",
    "--package",
    "@bilig/workpaper@latest",
    "--",
    "bilig-workpaper-mcp",
    "--workpaper",
    "__WORKPAPER_PATH__",
    "--init-demo-workpaper",
    "--writable"
  ]
}
```

Expected tools include `read_range`, `read_cell`,
`set_cell_contents_and_readback`, `export_workpaper_document`, and
`validate_formula`.

## Required Proof

For every workbook edit, read the relevant range first, write one precise input
or formula cell, read the dependent calculated output, export or serialize the
WorkPaper document, and report `editedCell`, `before`, `after`, persistence or
restore evidence, `verified`, and limitations. Do not claim success from a
write call alone.

Docs: <https://proompteng.github.io/bilig/openhands-workpaper-mcp.html>
