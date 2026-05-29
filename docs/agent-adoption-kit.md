---
title: Agent Adoption Kit for WorkPaper formulas
published: true
description: One no-key adoption path for coding agents: install the Bilig WorkPaper skill, run the MCP proof, paste the workbook task, and return verified formula readback.
tags: ai agents, mcp, workbook formulas, spreadsheet automation, node
canonical_url: https://proompteng.github.io/bilig/agent-adoption-kit.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Agent Adoption Kit

Use this page when you are evaluating whether Bilig should be the workbook tool
inside a coding agent, MCP client, or agent framework. It is intentionally one
path: install the agent instructions, run the no-key proof, paste one workbook
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

- Claude Code: `.claude/skills/bilig-workpaper/SKILL.md`
- Cursor: `.cursor/rules/bilig-workpaper.mdc`
- Windsurf/Cascade: `.windsurf/rules/bilig-workpaper.md`
- Cline: `.clinerules/bilig-workpaper.md`
- Continue: `.continue/rules/bilig-workpaper.md`

For web fetch, give the agent the compact map first:

```text
https://proompteng.github.io/bilig/llms.txt
```

## Run The No-Key Proof

This proves the published package and the file-backed MCP tool path without
cloning the repo or using an API key:

```sh
npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json
```

A passing run must return `verified: true` plus these checks:

- tools, resources, and prompts were discovered;
- one input cell changed;
- a dependent formula cell changed after recalculation;
- WorkPaper JSON was exported and persisted;
- restart readback matched the post-edit value.

Use the direct API challenge when the agent will import `@bilig/workpaper`
instead of using MCP:

```sh
npm exec --package @bilig/workpaper@latest -- bilig-agent-challenge --json
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
        "./pricing.workpaper.json",
        "--init-demo-workpaper",
        "--writable"
      ]
    }
  }
}
```

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

The exact values depend on the starter workbook, but the proof shape should look
like this:

```json
{
  "editedCell": "Inputs!B3",
  "dependentCell": "Summary!B3",
  "before": 60000,
  "after": 96000,
  "afterRestore": 96000,
  "persistedDocumentBytes": 1162,
  "verified": true,
  "limitations": []
}
```

Reject answers that only say a cell was written. The point of Bilig is that the
agent can prove the calculated result and the saved state.

## After The Proof

If the proof matches your workflow, keep the repo so you can find it again:
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
