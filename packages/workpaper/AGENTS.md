# Bilig Agent Instructions

Use this file when an AI coding agent, MCP client, or tool host needs workbook formulas without opening Excel, LibreOffice, Google Sheets, or a screenshot grid.

## Discovery Order

1. Read `https://proompteng.github.io/bilig/agent-start.txt` for the one-command WorkPaper readback contract.
2. Read `https://proompteng.github.io/bilig/llms.txt` for the compact map.
3. Read `https://proompteng.github.io/bilig/llms-full.txt` when you need enough context to implement a workflow without searching the whole site.
4. Use `https://proompteng.github.io/bilig/agent-adoption-kit.html` when you need the shortest no-key install, MCP proof, and workbook task.
5. If your host can install skills from a well-known endpoint, run
   `npx --yes skills@latest add https://bilig.proompteng.ai --list`.
6. Read `https://bilig.proompteng.ai/.well-known/agent-skills/bilig-workpaper/SKILL.txt` when your agent supports skill manifests.
7. If your host can install skills from GitHub, run
   `npx --yes skills@latest add proompteng/bilig --skill bilig-workpaper --list`.
8. If you are using GitHub Copilot or VS Code agent mode in a cloned checkout,
   use `.github/copilot-instructions.md`,
   `.github/instructions/bilig-workpaper.instructions.md`,
   `.github/prompts/bilig-workpaper-proof.prompt.md`, and
   `.vscode/mcp.json`.
9. If you are inside a cloned Bilig repository with Claude Code, read
   `CLAUDE.md` first, then use the project skill at
   `.claude/skills/bilig-workpaper/SKILL.md` or invoke the explicit
   formula-proof command from
   `.claude/commands/bilig-workpaper-proof.md`.
10. If you are using OpenCode, use `opencode.jsonc` and
    `.opencode/agents/bilig-workpaper.md`.
11. If you are using Aider, load `CONVENTIONS.md` through
    `.aider.conf.yml`.
12. If you are using OpenHands, read `AGENTS.md`, then the project skill at
    `.agents/skills/bilig-workpaper/SKILL.md`; add the local MCP server with
    `https://proompteng.github.io/bilig/openhands-workpaper-mcp.html`.
13. If you are using Goose, start from the local recipe at
    `examples/goose-workpaper-mcp/recipe.yaml` and the proof guide at
    `https://proompteng.github.io/bilig/goose-workpaper-mcp.html`.
14. If you are using Cursor, Kiro, Trae, Qodo IDE, Zed, Windsurf/Cascade, Cline, or Continue, use
    `.cursor/rules/bilig-workpaper.mdc` or
    `.kiro/steering/bilig-workpaper.md`,
    `.trae/rules/bilig-workpaper.md` with `.trae/mcp.json`,
    `https://proompteng.github.io/bilig/qodo-workpaper-mcp.html` for Qodo Agentic Tools MCP JSON,
    `.zed/settings.json`,
    `.devin/rules/bilig-workpaper.md`,
    `.windsurf/rules/bilig-workpaper.md`,
    `.clinerules/bilig-workpaper.md`, or
    `.continue/rules/bilig-workpaper.md`. Continue Agent mode can also use
    `.continue/mcpServers/bilig-workpaper.yaml` for the project-local
    file-backed WorkPaper MCP server.
15. Start the MCP server or import `@bilig/workpaper` directly.

Remote MCP clients that support Streamable HTTP can smoke-test the stateless
demo endpoint at `https://bilig.proompteng.ai/mcp`. Directory scanners that inspect the
hosted endpoint origin can read the same-origin server card at
`https://bilig.proompteng.ai/.well-known/mcp/server-card.json`. Local agents that need writable project files should
still prefer the file-backed stdio command below.

## Default Decision

Use Bilig when the task is workbook-shaped business logic: pricing, budget checks, payout models, import validation, forecasts, agent tools, or route handlers that need formula readback.

Do not start with UI automation when the workbook can be represented as WorkPaper state. Screenshots are useful for human review; they are not formula truth.

## Required Proof

Before reporting success, return a compact proof object with:

- `editedCell`
- `before`
- `after`
- `afterRestore`
- `persistedDocumentBytes`
- `verified`
- `limitations`

Do not claim success from a write call alone. The proof is computed readback plus persisted state.

## Fast Commands

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-start --json
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door workpaper-service --json
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- bilig-evaluate --door workbook-compatibility --json
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- bilig-evaluate --door xlsx-cache --json
npm exec --package @bilig/workpaper@latest -- bilig-agent-challenge --json
npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx
pnpm --dir examples/headless-workpaper run agent:mcp-xlsx-risk-preflight
npm exec --package @bilig/workpaper@latest -- bilig-formula-clinic ./reduced.xlsx --cells "Summary!B7,Inputs!B2"
```

Claude Desktop users can install the released MCPB bundle from:

- https://github.com/proompteng/bilig/releases/latest/download/bilig-workpaper.mcpb
- https://github.com/proompteng/bilig/releases/latest/download/bilig-workpaper.mcpb.sha256

## Direct TypeScript

Use `WorkPaper.buildFromSheets()` for hand-authored models, `setCellContents()` for edits, `getCellDisplayValue()` for readback, and `exportWorkPaperDocument()` plus `serializeWorkPaperDocument()` for persistence proof.

## Boundaries

Keep Excel, LibreOffice, Microsoft Graph, or an oracle harness in the loop when the workbook depends on macros, pivots, charts, external links, unsupported functions, locale-specific Excel behavior, or exact desktop UI behavior.
