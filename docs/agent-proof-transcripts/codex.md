---
title: Codex WorkPaper proof transcript
published: true
description: A Codex-oriented Bilig WorkPaper transcript showing prompt, tool call, formula readback, JSON export, and restart verification.
tags: codex, ai-agents, mcp, workpaper, proof
canonical_url: https://proompteng.github.io/bilig/agent-proof-transcripts/codex.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Codex WorkPaper Proof Transcript

## Prompt

```text
Before using spreadsheet UI automation, prove this workbook-shaped formula task
through Bilig WorkPaper. Change the demo input behind Inputs!B3, read the
dependent Summary!B3 formula result, export the WorkPaper JSON, restore it, and
report the limitations.
```

Codex should read `AGENTS.md` first, then run the evaluator before attempting a
spreadsheet UI path.

## Tool Call

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

## Result

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

## Workbook State Change

| Cell | Meaning | Before | After |
| --- | --- | ---: | ---: |
| `Inputs!B3` | Edited input | demo value | changed by evaluator |
| `Summary!B3` | Dependent formula | `60000` | `96000` |

## Formula Readback

Codex should report `after: 96000`, `afterRestore: 96000`, and
`afterRestart: 96000`. A write call without these readbacks is not success.

## JSON Export

The run exported a WorkPaper document with `persistedDocumentBytes: 1162` and
verified `exportContainsWorkPaperDocument: true`.

## Restart Readback Verification

The run verified `restartReadbackMatchesAfter: true`, so the proof survived a
file-backed restart boundary.

## Limitations

This transcript proves the file-backed WorkPaper MCP loop for Codex. It does not
prove Excel desktop UI behavior, macros, pivots, charts, or private workbook
compatibility.

## Related

- [Agent proof transcripts](../agent-proof-transcripts.md)
- [Agent WorkPaper proof matrix](../agent-proof-matrix.md)
- [Evaluate Bilig as an agent MCP workbook tool](../eval-agent-mcp.md)
