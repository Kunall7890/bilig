---
description: Use Bilig WorkPaper when Cascade needs workbook formulas, MCP spreadsheet tools, or formula readback without Excel UI automation.
trigger: model_decision
---

# Bilig WorkPaper Rule

Use this rule when a Windsurf/Cascade agent needs spreadsheet-shaped business
logic it can verify from code: pricing models, quote approvals, budget checks,
payout rules, import validation, forecasts, or other formula-backed workflows.

Avoid spreadsheet UI automation when WorkPaper state can represent the model.
Screenshots and browser grids are human-review surfaces, not formula truth.

## First Checks

```sh
npm exec --package @bilig/workpaper@latest -- bilig-agent-challenge --json
npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json
```

For a writable project file, run:

```sh
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
```

For remote MCP connector smoke tests only:

```text
https://bilig.proompteng.ai/mcp
```

## Readback Standard

Before saying the workbook is updated, return readback with:

- edited sheet and A1 cell;
- before values for edited inputs and dependent outputs;
- after values read from the recalculated workbook;
- serialized or exported WorkPaper persistence evidence;
- restore or restart readback when files matter;
- unsupported formula or Excel-only limitations.

Do not claim success from a write call alone.

If any readback step fails, say what failed instead of claiming the workbook was updated.

## Command Safety

Do not build shell commands by concatenating user text. Prefer MCP
`command` plus `args` arrays or direct TypeScript calls. Reject workbook
paths or cell arguments containing newlines, backticks, `$(`, `;`, `&`,
`|`, `<`, or `>`.

## References

- Docs map: https://proompteng.github.io/bilig/llms.txt
- Agent handbook: https://proompteng.github.io/bilig/headless-workpaper-agent-handbook.html
- MCP guide: https://proompteng.github.io/bilig/mcp-workpaper-tool-server.html
- Repository: https://github.com/proompteng/bilig
