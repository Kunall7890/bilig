---
title: Coding agent rule chooser for Bilig WorkPaper
published: true
description: Pick the Bilig instruction, rule, prompt, or MCP config for Codex, Claude Code, GitHub Copilot, VS Code, Cursor, Windsurf, Cline, Continue, and Gemini CLI.
tags: ai-agents, agent-rules, mcp, workbook formulas, coding agents
canonical_url: https://proompteng.github.io/bilig/agent-rule-chooser.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Coding Agent Rule Chooser

Use this page when a coding agent is in a cloned repo and you need to know
which Bilig file it should read before spreadsheet-shaped work.

Run the no-key agent MCP proof first:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

The result must include `schemaVersion: "bilig-evaluator.v1"`,
`door: "agent-mcp"`, `verified: true`, edited cell evidence, formula readback,
exported or persisted WorkPaper state, and restore or restart readback. A write
call alone is not success.

## Quick Choice

| Agent host | Use this Bilig file | Tool hookup | Proof bar |
| --- | --- | --- | --- |
| Codex | `AGENTS.md` in the repo directory chain. Public handoff: `docs/AGENTS.md`. | Optional `.mcp.json` when the Codex environment supports MCP. | Run `bilig-agent-start --json`, then `bilig-evaluate --door agent-mcp --json`. |
| Claude Code | `CLAUDE.md`, then `.claude/skills/bilig-workpaper/SKILL.md` or `.claude/commands/bilig-workpaper-proof.md`. | `.mcp.json` defines the file-backed `bilig-workpaper` stdio server. | Use `/bilig-workpaper-proof <task>` before Excel, LibreOffice, Sheets, browser grids, or screenshots. |
| GitHub Copilot | `.github/copilot-instructions.md` plus `.github/instructions/bilig-workpaper.instructions.md`. | `.github/prompts/bilig-workpaper-proof.prompt.md` for the task prompt, `.vscode/mcp.json` in VS Code. | Copilot should return WorkPaper readback fields, not spreadsheet UI status. |
| VS Code agent mode | `.github/copilot-instructions.md` and `.github/instructions/bilig-workpaper.instructions.md`. | `.vscode/mcp.json` for `biligWorkpaperDemo` and `biligWorkpaperFile`. | Use the workspace MCP config before copying a generic `mcpServers` manifest. |
| Cursor | `.cursor/rules/bilig-workpaper.mdc`. | `.cursor/mcp.json` for local file-backed WorkPaper tools. | Treat `.cursorrules` as legacy; use the project rule and MCP config here. |
| Windsurf/Cascade | `.devin/rules/bilig-workpaper.md`, with `.windsurf/rules/bilig-workpaper.md` kept as a fallback. | Start with the same `bilig-evaluate --door agent-mcp --json` command, then file-backed MCP if state must persist. | The rule uses `trigger: model_decision`; require computed readback before reporting success. |
| Cline | `.clinerules/bilig-workpaper.md`. | Add MCP separately only if your Cline setup exposes custom MCP servers. | Cline should use the workspace rule when workbook formulas, cells, or MCP spreadsheet tools appear. |
| Continue | `.continue/rules/bilig-workpaper.md`. | Configure MCP separately through Continue if you want direct tool calls. | Local rules are version controlled; use this one for Agent, Chat, and Edit requests that touch workbook logic. |
| Gemini CLI | `gemini-extension.json` plus `gemini-workpaper-context.md`; generated starters also include `GEMINI.md`. | `gemini extensions install https://github.com/proompteng/bilig --ref main`. | The extension starts the `bilig-workpaper` MCP server and injects the WorkPaper proof context. |

## Existing Repo Overlay

For a repo that already has app code, add only the agent and MCP files:

```sh
npm create @bilig/workpaper@latest . -- --add-agent
```

That overlay writes `AGENTS.md`, `CLAUDE.md`, `GEMINI.md`, Copilot and VS Code
instructions, Cursor, Cline, Continue, Cascade/Devin and Windsurf rules, and
MCP configs. It does not overwrite an existing app `README.md` or
`package.json`.

For a blank agent-ready project, use:

```sh
npm create @bilig/workpaper@latest pricing-agent -- --agent
```

## Confusion Guards

- `docs/AGENTS.md` is a public handoff page. Codex reads `AGENTS.md` from the
  cloned repo directory chain.
- Claude Code reads `CLAUDE.md`, not `AGENTS.md`; this repo's project memory
  routes it to the Claude Code skill, slash command, and `.mcp.json`.
- `.vscode/mcp.json` uses the VS Code `servers` shape. `mcp/bilig-workpaper.mcp.json`
  is the reusable `mcpServers` shape for other clients.
- Cascade/Devin docs currently prefer `.devin/rules`; the `.windsurf/rules`
  mirror remains for compatible Windsurf/Cascade installs.
- `GEMINI.md` is the normal Gemini CLI context file, but this repo exposes the
  installable Gemini extension path first.

## Official Host Docs Checked

- [Codex AGENTS.md](https://github.com/openai/codex/blob/main/docs/agents_md.md)
- [Claude Code memory](https://code.claude.com/docs/en/memory)
- [GitHub Copilot response customization](https://docs.github.com/en/copilot/concepts/prompting/response-customization)
- [VS Code MCP configuration](https://code.visualstudio.com/docs/copilot/reference/mcp-configuration)
- [Cursor rules](https://docs.cursor.com/en/context/rules)
- [Windsurf/Cascade memories and rules](https://docs.windsurf.com/windsurf/cascade/memories)
- [Cline rules](https://docs.cline.bot/customization/cline-rules)
- [Continue rules](https://docs.continue.dev/customize/rules)
- [Gemini CLI GEMINI.md context](https://google-gemini.github.io/gemini-cli/docs/cli/gemini-md.html)

## Related

- [Agent Adoption Kit](agent-adoption-kit.md)
- [Agent WorkPaper proof matrix](agent-proof-matrix.md)
- [Agent proof transcripts](agent-proof-transcripts.md)
- [Headless WorkPaper agent handbook](headless-workpaper-agent-handbook.md)
- [Evaluate Bilig as an agent MCP workbook tool](eval-agent-mcp.md)
