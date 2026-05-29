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
npm run agent:verify
```

For MCP use:

```sh
npm run mcp:server
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
- write one small input or formula change;
- read the dependent calculated output after recalculation;
- export or serialize the WorkPaper document;
- restore or restart when file boundaries matter;
- report unsupported formulas or Excel-only features honestly;
- do not claim success from a write call alone.
