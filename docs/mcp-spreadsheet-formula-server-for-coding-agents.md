---
title: MCP spreadsheet formula server for coding agents
published: true
description: Test a Model Context Protocol spreadsheet formula server by proving tool discovery, workbook input edits, formula readback, resources, prompts, and persisted WorkPaper state.
tags: mcp, model-context-protocol, coding-agents, spreadsheet-formulas, workpaper
canonical_url: https://proompteng.github.io/bilig/mcp-spreadsheet-formula-server-for-coding-agents.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# MCP Spreadsheet Formula Server For Coding Agents

Use this page when a coding agent needs spreadsheet formulas through MCP and
should not click through Excel, LibreOffice, Google Sheets, or a browser grid.
The useful MCP server is not just "cell access." It must prove discovery,
write/readback, resource context, prompt handoff, and persisted state.

MCP defines servers around capabilities such as tools, resources, prompts, and
transports. Bilig keeps the formula runtime behind those protocol boundaries:
the MCP client discovers workbook tools, calls one write/readback tool, and gets
a structured proof object.

Official protocol references:

- <https://modelcontextprotocol.io/docs/learn/server-concepts>
- <https://modelcontextprotocol.io/specification/2025-11-25/server/tools>
- <https://github.com/modelcontextprotocol/typescript-sdk>

## Failure Mode

An agent can list or write spreadsheet cells, but the result it reports is only
a write-call status. That is not formula proof. The agent needs to read the
dependent formula value after the edit and prove that the value survives export
or restart.

## One Command

Run the evaluator from any Node machine:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

Expected output includes:

```json
{
  "schemaVersion": "bilig-evaluator.v1",
  "door": "agent-mcp",
  "verified": true,
  "evidence": {
    "editedCell": "Inputs!B3",
    "dependentCell": "Summary!B3",
    "before": 60000,
    "after": 96000,
    "afterRestore": 96000,
    "afterRestart": 96000,
    "checks": {
      "listedFileBackedTools": true,
      "listedResourcesAndPrompts": true,
      "dependentCellChanged": true,
      "persistedToDisk": true,
      "restartReadbackMatchesAfter": true
    }
  }
}
```

The exact version and byte counts can change. The stable fields are
`door: "agent-mcp"`, `verified: true`, `listedResourcesAndPrompts`,
`dependentCellChanged`, `persistedToDisk`, and `restartReadbackMatchesAfter`.

## What To Inspect

For a neutral MCP client smoke, use the Inspector guide:

```sh
npx -y @modelcontextprotocol/inspector@latest --cli \
  npm exec --yes --package @bilig/workpaper@latest -- bilig-workpaper-mcp \
  --method tools/list
```

The default demo server exposes:

```text
read_workpaper_summary
set_workpaper_input_cell
```

For project files, use the file-backed stdio server instead:

```sh
npm exec --package @bilig/workpaper@latest -- \
  bilig-workpaper-mcp \
  --workpaper ./pricing.workpaper.json \
  --init-demo-workpaper \
  --writable
```

That mode exposes general tools such as `list_sheets`, `read_range`,
`set_cell_contents_and_readback`, `validate_formula`, and
`export_workpaper_document`.

## Limitation

The hosted endpoint at `https://bilig.proompteng.ai/mcp` is useful for no-key
remote MCP smoke tests. It is stateless and should not be used as proof that a
private workbook file was persisted. For private writable state, use the local
file-backed stdio command.

## When Not To Use Bilig

Do not use Bilig as the first tool when the real requirement is manual
spreadsheet editing, macro execution, pivot tables, chart layout, Office add-ins,
or exact desktop Excel parity. Use the [compatibility limits](where-bilig-is-not-excel-compatible-yet.md)
before depending on it for production workbook imports.

## Related

- [Agent WorkPaper proof matrix](agent-proof-matrix.md)
- [MCP WorkPaper tool server](mcp-workpaper-tool-server.md)
- [MCP client setup](mcp-client-setup.md)
- [Spreadsheet MCP server comparison](spreadsheet-mcp-server-comparison.md)
- [ChatGPT Apps WorkPaper MCP](chatgpt-apps-workpaper-mcp.md)
