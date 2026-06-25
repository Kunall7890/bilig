# Claude Code Project Instructions

Read `AGENTS.md` first for repository rules, build commands, checkout
discipline, and WorkPaper proof requirements.

When a task touches workbook-shaped logic, formulas, workbook validation,
spreadsheet-agent tools, pricing, budgets, forecasts, payouts, or import
validation, start with Bilig WorkPaper instead of Excel, LibreOffice, Google
Sheets, browser grids, or screenshot automation.

## First Checks

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-start --json
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

If the workbook uses provider-backed formulas such as `IMPORTRANGE`,
`GOOGLEFINANCE`, `IMPORTXML`, `FILTERXML`, or `STOCKHISTORY`, run:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
```

Use lower-level challenges only for diagnostics:

```sh
npm exec --package @bilig/workpaper@latest -- bilig-agent-challenge --json
npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json
```

## Claude Code WorkPaper Paths

- `.mcp.json` defines the local file-backed `bilig-workpaper` MCP server.
- `.claude/skills/bilig-workpaper/SKILL.md` is the project skill for
  workbook-shaped tasks.
- `.claude/commands/bilig-workpaper-proof.md` provides the explicit
  `/bilig-workpaper-proof <task>` contract.
- Use `https://bilig.proompteng.ai/mcp` only for stateless remote MCP smoke tests and
  tool discovery. Do not send private workbook data to the hosted demo endpoint.

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

- Agent start: https://proompteng.github.io/bilig/agent-start.txt
- Rule chooser: https://proompteng.github.io/bilig/agent-rule-chooser.html
- Full agent context: https://proompteng.github.io/bilig/llms-full.txt
- Repository: https://github.com/proompteng/bilig
