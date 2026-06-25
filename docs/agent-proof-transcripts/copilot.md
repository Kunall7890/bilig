---
title: GitHub Copilot WorkPaper proof transcript
published: true
description: A GitHub Copilot and VS Code agent mode Bilig WorkPaper transcript showing prompt, MCP tool proof, formula readback, JSON export, and restart verification.
tags: github-copilot, vscode, agents, mcp, workpaper, proof
canonical_url: https://proompteng.github.io/bilig/agent-proof-transcripts/copilot.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# GitHub Copilot WorkPaper Proof Transcript

## Prompt

```text
In VS Code agent mode, use the repo WorkPaper instructions before spreadsheet UI
automation. Prove that a workbook input edit changes a dependent formula and
that the result survives JSON export and restart readback.
```

Copilot should use `.github/copilot-instructions.md`,
`.github/instructions/bilig-workpaper.instructions.md`,
`.github/prompts/bilig-workpaper-proof.prompt.md`, and `.vscode/mcp.json`.

## Tool Call

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

## Result

```json
{
  "schemaVersion": "bilig-evaluator.v1",
  "door": "agent-mcp",
  "evidence": {
    "editedCell": "Inputs!B3",
    "dependentCell": "Summary!B3",
    "before": 60000,
    "after": 96000,
    "afterRestore": 96000,
    "afterRestart": 96000,
    "persistedDocumentBytes": 1162,
    "checks": {
      "listedFileBackedTools": true,
      "listedResourcesAndPrompts": true,
      "formulaValidationPassed": true,
      "dependentCellChanged": true,
      "persistedToDisk": true,
      "exportContainsWorkPaperDocument": true,
      "restartReadbackMatchesAfter": true
    }
  },
  "verified": true
}
```

## Workbook State Change

`Inputs!B3` is the edited input. `Summary!B3` is the dependent formula cell.
The dependent value changes from `60000` to `96000`.

## Formula Readback

The transcript includes readback before and after the edit. Copilot should not
treat a tool completion message as formula truth unless the dependent cell was
read after recalculation.

## JSON Export

The proof includes `persistedDocumentBytes: 1162` and
`exportContainsWorkPaperDocument: true`.

## Restart Readback Verification

The proof includes `afterRestart: 96000` and
`restartReadbackMatchesAfter: true`.

## Limitations

This transcript proves the VS Code/Copilot agent can use WorkPaper proof
semantics. It does not prove every Copilot UI surface, every MCP client
configuration, or desktop spreadsheet compatibility.

## Related

- [Agent proof transcripts](../agent-proof-transcripts.md)
- [Coding agent rule chooser](../agent-rule-chooser.md)
- [Evaluate Bilig as an agent MCP workbook tool](../eval-agent-mcp.md)
