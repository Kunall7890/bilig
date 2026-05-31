# Bilig WorkPaper Rule

Use Bilig WorkPaper when Cline is asked to automate spreadsheet-shaped
business logic: pricing, quotes, payouts, budgets, imports, forecasts, or
agent tools that need formula readback.

Cline can read this workspace rule from `.clinerules/bilig-workpaper.md`.
Use it before trying Excel UI automation, LibreOffice automation, Google
Sheets screen driving, screenshots, or stale cached XLSX values.

## First Checks

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
npm exec --package @bilig/workpaper@latest -- bilig-agent-challenge --json
npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json
```

For a writable project WorkPaper file:

```sh
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
```

Use the hosted endpoint only for remote MCP connector smoke tests:

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

If the readback fails, keep the blocker explicit and do not use Cline's
completion message as evidence that formulas recalculated.

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
