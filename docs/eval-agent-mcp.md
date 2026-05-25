---
title: Evaluate Bilig as an agent MCP workbook tool
published: true
description: Copy-paste evaluator for coding agents and MCP clients that need workbook writes, formula readback, JSON persistence, and restart proof.
tags: agents, mcp, workpaper, spreadsheet, evaluator
canonical_url: https://proompteng.github.io/bilig/eval-agent-mcp.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Evaluate Bilig as an agent MCP workbook tool

Use this when an agent is about to drive a spreadsheet UI by screenshots or
clicks. The narrower contract is better: list workbook tools, write one input
cell, read the dependent formula output, export WorkPaper JSON, restart from the
persisted file, and return proof.

## One command

```sh
npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json
```

## Expected proof

The current challenge prints this shape:

```json
{
  "transport": "stdio-json-rpc",
  "serverName": "bilig-headless-workpaper",
  "tools": [
    "list_sheets",
    "read_range",
    "read_cell",
    "set_cell_contents",
    "get_cell_display_value",
    "export_workpaper_document",
    "validate_formula"
  ],
  "editedCell": "Inputs!B3",
  "dependentCell": "Summary!B3",
  "before": 60000,
  "after": 96000,
  "afterRestart": 96000,
  "displayValue": "96000",
  "persistence": {
    "persisted": true,
    "serializedBytes": 1162
  },
  "checks": {
    "listedFileBackedTools": true,
    "listedResourcesAndPrompts": true,
    "formulaValidationPassed": true,
    "dependentCellChanged": true,
    "persistedToDisk": true,
    "exportContainsWorkPaperDocument": true,
    "restartReadbackMatchesAfter": true,
    "displayValueRead": true
  },
  "verified": true,
  "star": "https://github.com/proompteng/bilig/stargazers",
  "watchReleases": "https://github.com/proompteng/bilig/subscription",
  "adoptionBlocker": "https://github.com/proompteng/bilig/discussions/new?category=general"
}
```

The exact byte count can change. The invariants are `dependentCellChanged`,
`persistedToDisk`, `restartReadbackMatchesAfter`, and `verified: true`.

## What this proves

- the published package exposes a file-backed MCP stdio server
- an agent can discover spreadsheet tools and prompts
- an input edit changes a dependent formula result
- the updated WorkPaper document can be exported and persisted
- restart readback matches the calculated value after the edit

## What this does not prove

This does not prove arbitrary workbook compatibility, macros, pivots, charts,
external links, unsupported formulas, or desktop Excel parity. It proves the
agent tool contract: no screenshot truth, no blind write-only success, and no
missing persistence proof.

## After the proof

- Star Bilig if this gives your agent the workbook tool contract it needed:
  <https://github.com/proompteng/bilig/stargazers>
- Watch releases for MCP and agent-tool updates:
  <https://github.com/proompteng/bilig/subscription>
- Report the exact adoption blocker:
  <https://github.com/proompteng/bilig/discussions/new?category=general>

## Related

- [Agent workbook challenge](agent-workbook-challenge.md)
- [MCP WorkPaper tool server](mcp-workpaper-tool-server.md)
- [Headless WorkPaper agent handbook](headless-workpaper-agent-handbook.md)
- [MCP client setup](mcp-client-setup.md)
