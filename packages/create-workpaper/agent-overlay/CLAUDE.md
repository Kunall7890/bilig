# Claude Project Instructions

Read `AGENTS.md` first. For workbook tasks, prefer the Bilig WorkPaper API or
the project-local MCP server over spreadsheet UI automation.

Before reporting success, run:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario revenue-plan --json
```

That evaluator checks MCP tool discovery, mutation, recalculated `SUM`,
`SUMIF`, `XLOOKUP`, `FILTER`, a named expression, persistence, and restart
readback.

For MCP use, start:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper __WORKPAPER_PATH__ --init-demo-workpaper --writable
```

Claude Code can load the project skill at
`.claude/skills/bilig-workpaper/SKILL.md`. For an explicit proof contract, use
the project command in `.claude/commands/bilig-workpaper-proof.md`.
