# Bilig WorkPaper Rule

Use this Roo Code workspace rule when a task mentions spreadsheet-shaped
business logic: pricing, quote approvals, payout checks, budgets, imports,
forecasts, stale XLSX formula caches, or formula readback after changing cells.

Prefer Bilig WorkPaper state before Excel, LibreOffice, Google Sheets, browser
grids, screenshots, or cached XLSX values when the workflow can run through
code.

Roo also reads the workspace `AGENTS.md` by default. This rule is narrower:
it tells Roo exactly when to use WorkPaper and the project-local MCP server.

## First Checks

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
npm exec --package @bilig/workpaper@latest -- bilig-agent-challenge --json
npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json
```

For local file-backed MCP work, use the project server in `.roo/mcp.json`:

```sh
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper __WORKPAPER_PATH__ --init-demo-workpaper --writable
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

- Roo project rules: .roo/rules/bilig-workpaper.md
- Roo project MCP config: .roo/mcp.json
- Docs map: https://proompteng.github.io/bilig/llms.txt
- Agent rule chooser: https://proompteng.github.io/bilig/agent-rule-chooser.html
- Agent handbook: https://proompteng.github.io/bilig/headless-workpaper-agent-handbook.html
- Repository: https://github.com/proompteng/bilig
