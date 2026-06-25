---
title: Agent proof transcripts
published: true
description: Real Bilig WorkPaper proof transcripts for coding agents that need formula readback, JSON persistence, and restart verification without spreadsheet UI automation.
tags: agents, mcp, workpaper, spreadsheet automation, proof
canonical_url: https://proompteng.github.io/bilig/agent-proof-transcripts.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Agent Proof Transcripts

Use these pages when a reviewer asks what a successful Bilig session actually
looks like. They are terminal transcripts, not screenshots and not UI recordings.
The shared proof comes from a public no-key run:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

Observed package versions in the transcript:

```json
{
  "@bilig/workpaper": "0.157.0",
  "xlsx-formula-recalc": "0.157.0"
}
```

The important proof shape is the same for every agent host:

| Step | Evidence |
| --- | --- |
| Prompt | Ask the agent to prove a workbook-shaped formula edit before driving Excel, LibreOffice, Google Sheets, or a browser grid. |
| Tool call | Run `bilig-evaluate --door agent-mcp --json` or the local MCP equivalent. |
| Result | Return `schemaVersion: "bilig-evaluator.v1"`, `door: "agent-mcp"`, and `verified: true`. |
| Workbook state change | `Inputs!B3` changes the dependent `Summary!B3` value from `60000` to `96000`. |
| Formula readback | `after`, `afterRestore`, and `afterRestart` all read `96000`. |
| JSON export | `persistedDocumentBytes` is `1162`, and the exported WorkPaper document is restored. |
| Restart readback verification | `restartReadbackMatchesAfter: true`. |

Pick the host transcript that matches the agent you are using:

- [Codex](agent-proof-transcripts/codex.md)
- [Claude Code](agent-proof-transcripts/claude-code.md)
- [GitHub Copilot and VS Code agent mode](agent-proof-transcripts/copilot.md)
- [Cursor](agent-proof-transcripts/cursor.md)
- [Continue](agent-proof-transcripts/continue.md)

## Shared Verified Output

This is the compact output every transcript checks:

```json
{
  "schemaVersion": "bilig-evaluator.v1",
  "door": "agent-mcp",
  "packageVersions": {
    "@bilig/workpaper": "0.157.0",
    "xlsx-formula-recalc": "0.157.0"
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
  },
  "verified": true
}
```

## What This Proves

The transcript proves that a coding agent can discover the WorkPaper MCP tool
surface, write a cell, read a dependent formula, export JSON, restore the
document, and verify restart readback.

It does not prove Excel desktop UI automation, macro execution, pivot refresh,
chart layout, private workbook compatibility, or hosted multi-user storage.

## Related

- [Agent WorkPaper proof matrix](agent-proof-matrix.md)
- [Evaluate Bilig as an agent MCP workbook tool](eval-agent-mcp.md)
- [MCP spreadsheet formula server for coding agents](mcp-spreadsheet-formula-server-for-coding-agents.md)
- [WorkPaper agent handbook](headless-workpaper-agent-handbook.md)
