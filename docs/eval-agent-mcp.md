---
title: Evaluate Bilig as an agent MCP workbook tool
published: true
description: Evaluate MCP workbook tools for agents that need workbook writes, formula readback, JSON persistence, and restart proof.
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
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

If you are handing this to another coding agent, start from the
[Agent Adoption Kit](agent-adoption-kit.md). It includes the installable skill,
one MCP config, a workbook task, and the pass/fail proof object.

If you need the raw JSON-RPC challenge output, run:

```sh
npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json
```

## Expected proof

The evaluator prints this shape:

```json
{
  "schemaVersion": "bilig-evaluator.v1",
  "door": "agent-mcp",
  "doorName": "Agent MCP proof",
  "verified": true,
  "packageVersions": {
    "@bilig/workpaper": "0.140.0",
    "xlsx-formula-recalc": "0.140.0"
  },
  "evidence": {
    "editedCell": "Inputs!B3",
    "dependentCell": "Summary!B3",
    "before": 60000,
    "after": 96000,
    "afterRestore": 96000,
    "afterRestart": 96000,
    "persistedDocumentBytes": 1162,
    "toolCount": 8,
    "tools": [
      "list_sheets",
      "read_range",
      "read_cell",
      "set_cell_contents",
      "set_cell_contents_and_readback",
      "get_cell_display_value",
      "export_workpaper_document",
      "validate_formula"
    ],
    "checks": {
      "listedFileBackedTools": true,
      "listedResourcesAndPrompts": true,
      "formulaValidationPassed": true,
      "dependentCellChanged": true,
      "persistedToDisk": true,
      "exportContainsWorkPaperDocument": true,
      "restartReadbackMatchesAfter": true,
      "displayValueRead": true
    }
  }
}
```

The exact package versions, byte count, and duration can change. The invariants
are `door: "agent-mcp"`, `dependentCellChanged`, `persistedToDisk`,
`restartReadbackMatchesAfter`, `displayValueRead`, and `verified: true`.

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
