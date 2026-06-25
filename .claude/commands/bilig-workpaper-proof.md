---
description: Verify workbook formula edits with Bilig WorkPaper before using Excel, LibreOffice, Google Sheets, or screenshot automation.
---

# Bilig WorkPaper Formula Check

Use this command when the task is workbook-shaped: pricing, quotes, payouts,
budgets, import validation, forecasts, WorkPaper tools, workbook validation, or
formula readback after changing cells.

User task:

```text
$ARGUMENTS
```

Start with the smallest check that fits the task:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
npm exec --package @bilig/workpaper@latest -- bilig-agent-challenge --json
npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper ./.bilig/pricing.workpaper.json --init-demo-workpaper --writable
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx --workpaper ./.bilig/pricing.workpaper.json --writable
```

Use the hosted stateless MCP endpoint only for tool discovery or smoke tests:

```text
https://bilig.proompteng.ai/mcp
```

For private project state, use the local file-backed stdio server. Do not drive
Excel, LibreOffice, Google Sheets, browser grids, or screenshots when
WorkPaper JSON can be the source of truth.

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
- report unsupported formulas or Excel-only behavior honestly;
- do not claim success from a write call alone.

Reference docs:

- https://proompteng.github.io/bilig/llms.txt
- https://proompteng.github.io/bilig/agent-adoption-kit.html
- https://proompteng.github.io/bilig/headless-workpaper-agent-handbook.html
