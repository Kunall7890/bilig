# Bilig WorkPaper Instructions For Gemini

Use `@bilig/workpaper` as the source of truth for workbook logic in this
project. Prefer the WorkPaper API or the project-local MCP server over Excel,
LibreOffice, Google Sheets, browser grids, screenshots, or stale cached XLSX
values.

Before reporting success, run:

```sh
npm run agent:verify
```

For MCP use, start:

```sh
npm run mcp:server
```

Return proof with:

- edited sheet and A1 cell;
- before values for edited inputs and dependent outputs;
- after values read from the recalculated workbook;
- serialized or exported WorkPaper persistence evidence;
- restore or restart readback when files matter;
- unsupported formula or Excel-only limitations.

Do not claim success from a write call alone. If a proof step fails, say which
step failed and what evidence is missing.
