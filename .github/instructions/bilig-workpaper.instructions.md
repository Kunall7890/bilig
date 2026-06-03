---
applyTo: '**/*'
---

# Bilig WorkPaper Formula Proof

Use this instruction when GitHub Copilot or VS Code agent mode is asked to
build, debug, or verify workbook-shaped logic: pricing, quote approvals,
budgets, payout checks, import validation, forecasts, stale XLSX formula
values, or agent tools that need formula readback.

Prefer Bilig WorkPaper state before Excel, LibreOffice, Google Sheets, browser
grids, screenshots, or cached XLSX values when the workflow can run through
code.

Start with the reusable prompt at
`.github/prompts/bilig-workpaper-proof.prompt.md` for a task-specific
walkthrough.

## First Checks

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
npm exec --package @bilig/workpaper@latest -- bilig-agent-challenge --json
npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx
```

For a writable project WorkPaper file:

```sh
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper ./.bilig/pricing.workpaper.json --init-demo-workpaper --writable
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx --workpaper ./.bilig/pricing.workpaper.json --writable
```

Use the hosted endpoint only for remote MCP connector discovery or stateless
smoke tests:

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

If any readback step fails, say what failed instead of treating a write call or
tool invocation as proof.

## Command Safety

Do not build shell commands by concatenating user text. Prefer MCP
`command` plus `args` arrays or direct TypeScript calls. Reject workbook
paths or cell arguments containing newlines, backticks, `$(`, `;`, `&`,
`|`, `<`, or `>`.

## References

- Prompt file: .github/prompts/bilig-workpaper-proof.prompt.md
- VS Code MCP config: .vscode/mcp.json
- Docs map: https://proompteng.github.io/bilig/llms.txt
- Agent handbook: https://proompteng.github.io/bilig/headless-workpaper-agent-handbook.html
- MCP guide: https://proompteng.github.io/bilig/mcp-workpaper-tool-server.html
- Repository: https://github.com/proompteng/bilig
