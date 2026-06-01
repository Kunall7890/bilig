# Claude Project Instructions

Read `AGENTS.md` first. For workbook tasks, prefer the Bilig WorkPaper API or
the project-local MCP server over spreadsheet UI automation.

Before reporting success, run:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

For MCP use, start:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
```

Claude Code can load the project skill at
`.claude/skills/bilig-workpaper/SKILL.md`. For an explicit proof contract, use
the project command in `.claude/commands/bilig-workpaper-proof.md`.
