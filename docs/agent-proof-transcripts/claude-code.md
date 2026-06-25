---
title: Claude Code WorkPaper proof transcript
published: true
description: A Claude Code-oriented Bilig WorkPaper transcript showing project memory, prompt, tool call, formula readback, JSON export, and restart verification.
tags: claude-code, agents, mcp, workpaper, proof
canonical_url: https://proompteng.github.io/bilig/agent-proof-transcripts/claude-code.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Claude Code WorkPaper Proof Transcript

## Prompt

```text
Use the repo WorkPaper instructions. Prove the workbook edit with Bilig before
driving a spreadsheet UI: change the demo input, read the dependent formula,
export JSON, restore, and return the verified proof object.
```

Claude Code should load `CLAUDE.md`, then use
`.claude/skills/bilig-workpaper/SKILL.md` or
`.claude/commands/bilig-workpaper-proof.md` for the proof contract.

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
      "dependentCellChanged": true,
      "persistedToDisk": true,
      "restartReadbackMatchesAfter": true
    }
  },
  "verified": true
}
```

## Workbook State Change

The transcript edits `Inputs!B3` and verifies the dependent `Summary!B3`
formula changes from `60000` to `96000`.

## Formula Readback

Claude Code should include the computed `after`, `afterRestore`, and
`afterRestart` values in its final answer. The proof is incomplete if the agent
only says the MCP write succeeded.

## JSON Export

The evaluator exported and restored a WorkPaper JSON document with
`persistedDocumentBytes: 1162`.

## Restart Readback Verification

The run checked `restartReadbackMatchesAfter: true`, which is the useful part
for a local coding agent: another process can reopen the document and see the
same dependent formula result.

## Limitations

This transcript proves Claude Code can follow a local WorkPaper proof path. It
does not prove that the hosted remote MCP endpoint should receive private
workbook data. Use local stdio/file-backed MCP for private workbooks.

## Related

- [Agent proof transcripts](../agent-proof-transcripts.md)
- [Coding agent rule chooser](../agent-rule-chooser.md)
- [WorkPaper agent handbook](../headless-workpaper-agent-handbook.md)
