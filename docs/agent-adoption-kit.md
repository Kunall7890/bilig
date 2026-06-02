---
title: Agent Adoption Kit for WorkPaper formulas
published: true
description: One no-key adoption path for coding agents: install the Bilig WorkPaper skill, run the agent MCP evaluator, paste the workbook task, and return verified formula readback.
tags: ai agents, mcp, workbook formulas, spreadsheet automation, node
canonical_url: https://proompteng.github.io/bilig/agent-adoption-kit.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Agent Adoption Kit

Use this page when you are deciding whether Bilig should be the workbook tool
inside a coding agent, MCP client, or agent framework. The path is deliberately
short: install the instructions, run one no-key evaluator, paste one workbook
task, and require formula readback before calling the job done.

## Install The Agent Instructions

If your agent supports installable skills, start here:

```sh
npx --yes skills@latest add https://bilig.proompteng.ai --list
npx --yes skills@latest add proompteng/bilig --skill bilig-workpaper --list
```

Use the app-host discovery URL first. Keep the GitHub repo skill command as a
fallback for hosts that only support GitHub skill sources.

If the agent is already inside a cloned Bilig checkout, use the project-local
rules instead:

- Claude Code: `.claude/skills/bilig-workpaper/SKILL.md` or
  `/bilig-workpaper-proof` from `.claude/commands/bilig-workpaper-proof.md`
- GitHub Copilot / VS Code agent mode:
  `.github/copilot-instructions.md`,
  `.github/prompts/bilig-workpaper-proof.prompt.md`, and `.vscode/mcp.json`
- Cursor: `.cursor/rules/bilig-workpaper.mdc`
- Windsurf/Cascade: `.windsurf/rules/bilig-workpaper.md`
- Cline: `.clinerules/bilig-workpaper.md`
- Continue: `.continue/rules/bilig-workpaper.md`

For a clean project that already contains those instruction files, create the
starter:

```sh
npm create @bilig/workpaper@latest pricing-agent -- --agent
```

The generated project includes `AGENTS.md`, `CLAUDE.md`, `GEMINI.md`,
Copilot / VS Code instructions, Cursor, Cline, Continue, and Windsurf rules,
plus local MCP configs and `npm run agent:verify`.

For an existing repo, add only the agent/MCP files:

```sh
npm create @bilig/workpaper@latest . -- --add-agent
```

That keeps the app template, existing `README.md`, and `package.json` intact.
The generated MCP configs use direct `npm exec` and store local workbook state
at `./.bilig/pricing.workpaper.json`, so agents can run the WorkPaper server
without needing project scripts or a root-level state file. If an agent policy
file already exists, the CLI leaves it untouched and writes
`BILIG_WORKPAPER_INSTALL.md` with the skipped paths and the short handoff block
to paste into the current policy.

For web fetch, give the agent the compact map first:

```text
https://proompteng.github.io/bilig/llms.txt
```

## Run The No-Key Check

This checks the published package and the file-backed MCP tool path without
cloning the repo or using an API key:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

For a richer workbook check, use the revenue-plan scenario:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario revenue-plan --json
```

That scenario proves `SUM`, `SUMIF`, `XLOOKUP`, `FILTER`, a named expression,
JSON persistence, and restart readback through the same MCP door.

If the workbook includes provider-backed formulas such as `IMPORTRANGE`, run the
adapter-boundary check:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
```

That check should show `#BLOCKED!` and `provider-backed-adapter-missing` before
a local synthetic adapter is installed, then a fresh `96000` readback with
diagnostics cleared after the adapter path runs. It does not call Google Sheets.

A passing run must return `schemaVersion: "bilig-evaluator.v1"`,
`door: "agent-mcp"`, `verified: true`, and these checks:

- tools, resources, and prompts were discovered;
- one input cell changed;
- a dependent formula cell changed after recalculation;
- WorkPaper JSON was exported and persisted;
- restart readback matched the post-edit value.

Use the raw MCP challenge only when you need the lower-level JSON-RPC proof:

```sh
npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json
```

Use the service evaluator when the agent will import `@bilig/workpaper` instead
of using MCP:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door workpaper-service --json
```

## Wire The Local MCP Server

Use file-backed stdio for private project state:

```json
{
  "mcpServers": {
    "bilig-workpaper": {
      "command": "npm",
      "args": [
        "exec",
        "--package",
        "@bilig/workpaper@latest",
        "--",
        "bilig-workpaper-mcp",
        "--workpaper",
        "./.bilig/pricing.workpaper.json",
        "--init-demo-workpaper",
        "--writable"
      ]
    }
  }
}
```

If the repository already has a workbook, create the file-backed WorkPaper from
that XLSX first:

```sh
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx --workpaper ./.bilig/pricing.workpaper.json --writable
```

That command imports the XLSX once, refuses to replace an existing WorkPaper
JSON unless `--overwrite-workpaper` is present, and then exposes the same
`read_cell`, `set_cell_contents_and_readback`, and `export_workpaper_document`
tools.

Use the hosted endpoint only for smoke tests and tool discovery:

```text
https://bilig.proompteng.ai/mcp
```

The hosted endpoint is stateless. It is not where private workbook files live.

## Paste This Task Into An Agent

```text
Use Bilig WorkPaper tools instead of spreadsheet UI automation. Build or load a
small workbook with Inputs!B2 as customers, Inputs!B3 as average revenue, and
Summary!B3 as the revenue formula. First read the relevant input and summary
range. Then set Inputs!B3 to 4800, read Summary!B3 after recalculation, export
or serialize the WorkPaper document, restore it, and return editedCell, before,
after, afterRestore, persistedDocumentBytes, verified, and limitations.

Do not claim success from a write call alone. Success requires computed
readback plus persisted or restored state.
```

## Expected Result

The exact values can change, but the evaluator result should look like this:

```json
{
  "schemaVersion": "bilig-evaluator.v1",
  "door": "agent-mcp",
  "verified": true,
  "packageVersions": {
    "@bilig/workpaper": "0.153.0",
    "xlsx-formula-recalc": "0.153.0"
  },
  "evidence": {
    "editedCell": "Inputs!B3",
    "dependentCell": "Summary!B3",
    "before": 60000,
    "after": 96000,
    "afterRestore": 96000,
    "afterRestart": 96000,
    "checks": {
      "listedFileBackedTools": true,
      "listedResourcesAndPrompts": true,
      "formulaValidationPassed": true,
      "dependentCellChanged": true,
      "persistedToDisk": true,
      "exportContainsWorkPaperDocument": true,
      "restartReadbackMatchesAfter": true
    }
  }
}
```

Reject answers that only say a cell was written. The point of Bilig is that the
agent returns the calculated result and saved state from the workbook itself.

## After The Check

If the check matches your workflow, keep the repo so you can find it again:
<https://github.com/proompteng/bilig/stargazers>.

If you need release notifications for agent or MCP changes, watch releases:
<https://github.com/proompteng/bilig/subscription>.

If it almost works but adoption is blocked, open the concrete blocker:
<https://github.com/proompteng/bilig/discussions/new?category=general>.

## Next Pages

- [Evaluate Bilig as an agent MCP workbook tool](eval-agent-mcp.md)
- [Headless WorkPaper agent handbook](headless-workpaper-agent-handbook.md)
- [MCP client setup](mcp-client-setup.md)
- [Agent workbook challenge](agent-workbook-challenge.md)
- [Workbook tools for agent frameworks](agent-framework-workbook-tools.md)
