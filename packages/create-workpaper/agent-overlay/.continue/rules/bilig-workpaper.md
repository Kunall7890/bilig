---
name: Bilig WorkPaper Formula Proof
description: Use Bilig WorkPaper for spreadsheet-shaped logic that needs formula readback proof without Excel UI automation.
---

# Bilig WorkPaper Formula Proof

Use this rule when a Continue agent is about to build or debug workbook-shaped
logic: pricing, quote approval, payout checks, import validation, budgets,
forecasts, or agent tools that need formulas and persisted state.

Do not start by driving a spreadsheet UI when WorkPaper JSON can represent the
model. Screenshots are human review evidence, not formula truth.

## First Checks

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
npm exec --yes --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper __WORKPAPER_PATH__ --init-demo-workpaper --writable
```

## Required Proof

Before saying the workbook is updated, return proof with:

- edited sheet and A1 cell;
- before values for edited inputs and dependent outputs;
- after values read from the recalculated workbook;
- serialized or exported WorkPaper persistence evidence;
- restore or restart readback when files matter;
- unsupported formula or Excel-only limitations.

Do not claim success from a write call alone.

## Command Safety

Do not build shell commands by concatenating user text. Prefer MCP `command`
plus `args` arrays or direct TypeScript calls. Reject workbook paths or cell
arguments containing newlines, backticks, `$(`, `;`, `&`, `|`, `<`, or `>`.

References:

- Agent adoption kit: https://proompteng.github.io/bilig/agent-adoption-kit.html
- Docs map: https://proompteng.github.io/bilig/llms.txt
