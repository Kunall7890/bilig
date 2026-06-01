---
name: bilig-workpaper-proof
description: Prove a WorkPaper formula edit with readback and persistence evidence.
agent: agent
---

Use this prompt when the task is workbook-shaped: pricing, quotes, budgets,
payout checks, import validation, forecasts, or agent spreadsheet tools.

Task: ${input:task:Describe the workbook or formula workflow}

Start with:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario revenue-plan --json
```

That evaluator checks MCP tool discovery, mutation, recalculated `SUM`,
`SUMIF`, `XLOOKUP`, `FILTER`, a named expression, persistence, and restart
readback.

For MCP use:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper __WORKPAPER_PATH__ --init-demo-workpaper --writable
```

Return proof, not a status sentence:

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
- prefer `set_cell_contents_and_readback` for one-call edit plus dependent
  output readback;
- otherwise write one small input or formula change and then read the
  dependent calculated output after recalculation;
- export or serialize the WorkPaper document;
- restore or restart when file boundaries matter;
- report unsupported formulas or Excel-only features honestly;
- do not claim success from a write call alone.
