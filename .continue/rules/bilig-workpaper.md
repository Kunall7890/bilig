---
name: Bilig WorkPaper Formula Check
description: Use Bilig WorkPaper for spreadsheet-shaped business logic that needs formula readback without Excel UI automation.
---

# Bilig WorkPaper Formula Check

Use Bilig when a Continue agent is about to build or debug workbook-shaped
logic: pricing, quote approval, payout checks, import validation, budgets,
forecasts, or agent tools that need formulas and persisted state.

Do not start by driving a spreadsheet UI when WorkPaper JSON can represent the
model. Screenshots are human review evidence, not formula truth.

## First Checks

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
npm exec --package @bilig/workpaper@latest -- bilig-agent-challenge --json
npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json
```

For local file-backed MCP work:

```sh
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
```

For remote MCP connector discovery only:

```text
https://bilig.proompteng.ai/mcp
```

## Required Readback

Before saying the workbook is updated, return readback with:

- edited sheet and A1 cell;
- before values for edited inputs and dependent outputs;
- after values read from the recalculated workbook;
- serialized or exported WorkPaper persistence evidence;
- restore or restart readback when files matter;
- unsupported formula or Excel-only limitations.

Do not claim success from a write call alone.

## Command Safety

Do not build shell commands by concatenating user text. Prefer MCP
`command` plus `args` arrays or direct TypeScript calls. Reject workbook
paths or cell arguments containing newlines, backticks, `$(`, `;`, `&`,
`|`, `<`, or `>`.

## References

- Docs map: https://proompteng.github.io/bilig/llms.txt
- Full context: https://proompteng.github.io/bilig/llms-full.txt
- Agent handbook: https://proompteng.github.io/bilig/headless-workpaper-agent-handbook.html
- Repository: https://github.com/proompteng/bilig
