---
title: Cursor WorkPaper proof transcript
published: true
description: A Cursor-oriented Bilig WorkPaper transcript showing rule-file routing, tool call, formula readback, JSON export, and restart verification.
tags: cursor, ai-agents, mcp, workpaper, proof
canonical_url: https://proompteng.github.io/bilig/agent-proof-transcripts/cursor.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Cursor WorkPaper Proof Transcript

## Prompt

```text
Use the Cursor Bilig WorkPaper rule before spreadsheet UI automation. Prove that
the agent can edit a workbook input, read the recalculated formula, export JSON,
restore, and verify the restored result.
```

Cursor should load `.cursor/rules/bilig-workpaper.mdc` and use
`.cursor/mcp.json` when it needs a project-local MCP server.

## Tool Call

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

## Result

```json
{
  "schemaVersion": "bilig-evaluator.v1",
  "door": "agent-mcp",
  "sourceProof": {
    "transport": "stdio-json-rpc",
    "serverName": "bilig-headless-workpaper",
    "editedCell": "Inputs!B3",
    "dependentCell": "Summary!B3",
    "before": 60000,
    "after": 96000,
    "afterRestore": 96000,
    "afterRestart": 96000,
    "persistedDocumentBytes": 1162,
    "displayValue": "96000",
    "checks": {
      "listedFileBackedTools": true,
      "listedResourcesAndPrompts": true,
      "formulaValidationPassed": true,
      "dependentCellChanged": true,
      "persistedToDisk": true,
      "restartReadbackMatchesAfter": true
    }
  },
  "verified": true
}
```

## Workbook State Change

The proof changes `Inputs!B3`, then reads `Summary!B3` as `96000`.

## Formula Readback

Cursor should report the dependent cell readback and the display value, not only
the raw write response. In this transcript, `displayValue` is `"96000"`.

## JSON Export

The MCP proof exported a WorkPaper document and recorded
`persistedDocumentBytes: 1162`.

## Restart Readback Verification

The source proof includes `afterRestart: 96000` and
`restartReadbackMatchesAfter: true`.

## Limitations

This transcript proves the Cursor rule and WorkPaper MCP proof shape. It does
not prove browser-grid automation, a hosted multi-tenant workbook service, or
private file handling through the remote demo endpoint.

## Related

- [Agent proof transcripts](../agent-proof-transcripts.md)
- [Coding agent rule chooser](../agent-rule-chooser.md)
- [MCP spreadsheet formula server for coding agents](../mcp-spreadsheet-formula-server-for-coding-agents.md)
