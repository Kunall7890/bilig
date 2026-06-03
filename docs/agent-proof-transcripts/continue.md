---
title: Continue WorkPaper proof transcript
published: true
description: A Continue-oriented Bilig WorkPaper transcript showing rule-file routing, prompt, tool call, formula readback, JSON export, and restart verification.
tags: continue, ai-agents, mcp, workpaper, proof
canonical_url: https://proompteng.github.io/bilig/agent-proof-transcripts/continue.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Continue WorkPaper Proof Transcript

## Prompt

```text
Use the Continue Bilig WorkPaper rule before automating a spreadsheet UI. Prove
the workbook edit with formula readback, JSON export, restore, and restart
verification.
```

Continue should load `.continue/rules/bilig-workpaper.md`. If the workspace
needs local file persistence, use the file-backed MCP command from that rule.

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
      "restartReadbackMatchesAfter": true,
      "displayValueRead": true
    }
  },
  "verified": true,
  "limitations": [
    "This challenge proves the file-backed MCP WorkPaper tool surface, not Excel desktop UI automation.",
    "For XLSX-specific behavior, run bilig-formula-clinic or the XLSX recalculation example with a real workbook fixture."
  ]
}
```

## Workbook State Change

`Inputs!B3` is edited, and the dependent formula cell `Summary!B3` changes from
`60000` to `96000`.

## Formula Readback

The Continue response should include the `before`, `after`, `afterRestore`, and
`afterRestart` values. If one is missing, the transcript is not complete enough
to trust.

## JSON Export

The run records `persistedDocumentBytes: 1162` and
`exportContainsWorkPaperDocument: true`.

## Restart Readback Verification

The run records `restartReadbackMatchesAfter: true`, so the proof is not just an
in-memory edit.

## Limitations

This transcript proves a local WorkPaper proof path for Continue. It does not
prove Excel-only features, spreadsheet UI behavior, or unsupported formulas.

## Related

- [Agent proof transcripts](../agent-proof-transcripts.md)
- [Coding agent rule chooser](../agent-rule-chooser.md)
- [Agent WorkPaper proof matrix](../agent-proof-matrix.md)
