---
name: bilig-workpaper-proof
description: Verify a workbook formula edit with Bilig WorkPaper instead of spreadsheet UI automation.
agent: agent
---

Use this prompt when the task is workbook-shaped: pricing, quotes, budgets,
payout checks, import validation, forecasts, agent spreadsheet tools, stale
XLSX formula values, or formula readback after changing cells.

Task: ${input:task:Describe the workbook or formula workflow}

Read the repository instructions first:

- [Copilot instructions](../copilot-instructions.md)
- [WorkPaper agent handbook](../../docs/headless-workpaper-agent-handbook.md)
- [MCP client setup](../../docs/mcp-client-setup.md)

Start with the smallest check that matches the task:

```sh
npm exec --package @bilig/workpaper@latest -- bilig-agent-challenge --json
npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper ./.bilig/pricing.workpaper.json --init-demo-workpaper --writable
```

If VS Code MCP tools are available, prefer the workspace server named
`biligWorkpaperFile` for project-local persistence, or
`biligWorkpaperDemo` for no-file hosted smoke tests. The shared MCP config is
at [`.vscode/mcp.json`](../../.vscode/mcp.json).

Return readback, not a status sentence:

```json
{
  "editedCell": "Inputs!B3",
  "before": {},
  "after": {},
  "afterRestore": {},
  "persistedDocumentBytes": 0,
  "verified": false,
  "limitations": []
}
```

Rules:

- read the relevant input and dependent output before editing;
- write one small input or formula change;
- read the dependent calculated output after recalculation;
- export or serialize the WorkPaper document;
- restore or restart when file boundaries matter;
- report unsupported formulas or Excel-only features honestly;
- do not claim success from a write call alone.

Reference docs:

- https://proompteng.github.io/bilig/llms.txt
- https://proompteng.github.io/bilig/agent-adoption-kit.html
- https://proompteng.github.io/bilig/headless-workpaper-agent-handbook.html
