# `__PROJECT_NAME__`

Agent-ready formula WorkPaper starter built with `@bilig/workpaper`.

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

Generated starters also include `npm run agent:verify`, which runs the service
smoke and then the package-owned agent evaluator:

- `npm run smoke`: writes quote inputs through a service-style API handler,
  recalculates formulas, persists WorkPaper JSON, restores it, and checks
  `verified: true`.
- `npm run agent:evaluate`: runs
  `bilig-evaluate --door agent-mcp --json`, discovers MCP tools, edits a
  WorkPaper cell, reads the recalculated dependent value, exports JSON,
  restarts from disk, and checks `verified: true`.

Use `npm run mcp:challenge` only when you need the lower-level JSON-RPC
diagnostic transcript.

Start the local API:

```sh
npm run dev
curl http://localhost:8788/api/quote/approval
curl -X POST http://localhost:8788/api/quote/approval \
  -H 'content-type: application/json' \
  -d '{"units":40,"listPrice":1200,"discount":0.05,"unitCost":760,"minimumMargin":0.3}'
```

Start the persistent project-local MCP server:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper __WORKPAPER_PATH__ --init-demo-workpaper --writable
```

The server owns `__WORKPAPER_PATH__`, initializes it when missing, writes
through MCP tools, recalculates formulas, and persists edits back to disk.
Project MCP configs are included for Claude Code (`.mcp.json`), Cursor, and VS
Code. Claude Code also gets the project skill at
`.claude/skills/bilig-workpaper/SKILL.md`. Other clients can use the same
command from `mcp/bilig-workpaper.mcp.json`.

Agent instruction files are included for Claude Code, Codex/AGENTS-aware
tools, Gemini CLI, GitHub Copilot, VS Code agent mode, Cursor, Cline, Continue,
and Windsurf/Cascade. They all point at the same rule: read first, edit one
cell or formula, recalculate, read the dependent output, persist or export, and
report proof.

Existing projects can add only these agent files without replacing their app
template or editing `package.json`:

```sh
npm create @bilig/workpaper@latest . -- --add-agent
```

For existing repos, the generated MCP configs keep WorkPaper state under
`./.bilig/pricing.workpaper.json` so the overlay does not add a noisy root
workbook file.

Agent handoff:

```text
Use Bilig WorkPaper tools instead of spreadsheet UI automation. Read the
relevant range first, write one precise input or formula change, read the
dependent calculated output, export or serialize the WorkPaper document, and
report editedCell, before, after, persistence evidence, verified, and
limitations. Do not claim success from a write call alone.
```

Agent adoption kit: <https://proompteng.github.io/bilig/agent-adoption-kit.html>

Learn more: <https://github.com/proompteng/bilig>
