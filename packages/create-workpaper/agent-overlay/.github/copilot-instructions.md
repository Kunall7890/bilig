# WorkPaper Instructions For Copilot

Use `@bilig/workpaper` as the source of truth for workbook logic in this
project. Prefer the WorkPaper API or the project-local MCP server over Excel,
LibreOffice, Google Sheets, browser grids, screenshots, or stale cached XLSX
values.

Before reporting success, run:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

For MCP use, start:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
```

The workspace MCP config at `.vscode/mcp.json` exposes the same project-local
server.

Return proof with:

- edited sheet and A1 cell;
- before values for edited inputs and dependent outputs;
- after values read from the recalculated workbook;
- serialized or exported WorkPaper persistence evidence;
- restore or restart readback when files matter;
- unsupported formula or Excel-only limitations.

Do not claim success from a write call alone.

Use `.github/prompts/bilig-workpaper-proof.prompt.md` when a task needs an
explicit proof contract.
