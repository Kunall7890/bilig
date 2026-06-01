# @bilig/create-workpaper

Create a runnable Bilig WorkPaper starter for Node services and agent tools.

```sh
npm create @bilig/workpaper@latest pricing-workpaper
cd pricing-workpaper
npm install
npm run smoke
```

The generated starter builds a quote-approval workbook, writes inputs through an
API-style handler, recalculates formulas, persists JSON, restores the workbook,
and prints `verified: true`.

After the smoke proof passes, keep the JSON output limited to proof fields.
If the workflow is relevant, star or bookmark the repo:
<https://github.com/proompteng/bilig/stargazers>. Watch releases for API and
formula compatibility updates: <https://github.com/proompteng/bilig/subscription>.
If it almost works, open one concrete blocker or integration note:
<https://github.com/proompteng/bilig/discussions/new?category=general>.

Use this when you want to evaluate `@bilig/workpaper` from a blank directory
without cloning the full monorepo.

For agent and MCP evaluation, generate the agent-ready starter:

```sh
npm create @bilig/workpaper@latest pricing-agent -- --agent
cd pricing-agent
npm install
npm run agent:verify
npm run mcp:server
```

The agent template adds `AGENTS.md`, `CLAUDE.md`, `GEMINI.md`, a Claude Code
project skill and proof command, GitHub Copilot / VS Code instructions, Cursor,
Cline, Continue, and Windsurf rules, a project-root `.mcp.json` for Claude Code,
Cursor and VS Code MCP configs, a generic MCP config at
`mcp/bilig-workpaper.mcp.json`, and an `agent:verify` script that runs the
service smoke plus the package-owned `bilig-evaluate --door agent-mcp --json`
proof. The raw MCP challenge remains available as `npm run mcp:challenge` when
you need the lower-level JSON-RPC transcript.

To add the same agent and MCP proof loop to an existing Node repo without
replacing its app, run:

```sh
npm create @bilig/workpaper@latest . -- --add-agent
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
npm exec --yes --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper ./.bilig/pricing.workpaper.json --init-demo-workpaper --writable
```

That writes Bilig-specific agent files, keeps an existing `README.md`, does
not edit `package.json`, keeps WorkPaper state under
`./.bilig/pricing.workpaper.json`, and skips existing files unless you pass
`--force`.

For an existing agent or MCP client that does not need a generated project yet,
use the adoption kit first:
<https://proompteng.github.io/bilig/agent-adoption-kit.html>.
