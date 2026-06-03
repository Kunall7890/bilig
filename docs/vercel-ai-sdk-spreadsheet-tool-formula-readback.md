---
title: Vercel AI SDK spreadsheet tool: generateText and streamText with formula readback
published: true
description: Wrap Bilig WorkPaper as Vercel AI SDK tools so generateText and streamText can return verified spreadsheet formula readback instead of a write-only result.
tags: vercel-ai-sdk, ai-sdk, spreadsheet-tool, tool-calling, workpaper
canonical_url: https://proompteng.github.io/bilig/vercel-ai-sdk-spreadsheet-tool-formula-readback.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Vercel AI SDK Spreadsheet Tool: generateText And streamText With Formula Readback

Use this page when an AI SDK app needs workbook-shaped calculations behind a
tool call. The tool should return proof, not a vague "updated cell" message.

The AI SDK documents tool calling for `generateText()` and `streamText()`.
Bilig's wrapper keeps the AI SDK boundary thin: the model calls a tool, the
tool edits one WorkPaper input, formulas recalculate in Node, and the tool
returns before/after/restore proof.

Official AI SDK reference:

- <https://ai-sdk.dev/docs/ai-sdk-core/tools-and-tool-calling>

## Failure Mode

An agent tool changes `Inputs!B3`, but the app only records that the write call
completed. The model then explains a stale or unverified value. For workbook
logic, the tool result must include the dependent formula readback.

## One Command

Run the no-provider `generateText()` smoke from a clean checkout:

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig
pnpm --dir examples/headless-workpaper install --ignore-workspace
pnpm --dir examples/headless-workpaper run agent:ai-sdk-generate-text
```

Expected output includes:

```json
{
  "apiShape": "AI SDK generateText -> tool -> execute",
  "modelCallCount": 2,
  "toolNames": ["readWorkPaperSummary", "setWorkPaperInputCell"],
  "writeResult": {
    "editedCell": "Inputs!B3",
    "before": { "expectedArr": 60000, "targetGap": -34000 },
    "after": { "expectedArr": 96000, "targetGap": 5600 },
    "checks": {
      "formulasPersisted": true,
      "restoredMatchesAfter": true,
      "expectedArrChanged": true
    }
  }
}
```

For streaming tools:

```sh
pnpm --dir examples/headless-workpaper run agent:ai-sdk-stream-text
```

Expected streaming output includes:

```json
{
  "apiShape": "AI SDK streamText -> tool -> execute",
  "modelStreamCallCount": 2,
  "streamChunkTypes": ["tool-call", "tool-result", "tool-call", "tool-result", "text-delta", "text-delta"],
  "writeResult": {
    "editedCell": "Inputs!B3",
    "after": { "expectedArr": 96000, "targetGap": 5600 },
    "checks": {
      "restoredMatchesAfter": true,
      "expectedArrChanged": true
    }
  }
}
```

## Minimal Tool Boundary

The `@bilig/workpaper/ai-sdk` helper returns AI SDK `tool()` definitions:

```ts
import { WorkPaper } from '@bilig/workpaper'
import { createAiSdkWorkPaperTools } from '@bilig/workpaper/ai-sdk'

const workpaper = WorkPaper.buildFromSheets({
  Inputs: [
    ['Metric', 'Value'],
    ['Qualified opportunities', 20],
    ['Win rate', 0.25],
    ['Average ARR', 12000],
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Expected customers', '=Inputs!B2*Inputs!B3'],
    ['Expected ARR', '=B2*Inputs!B4'],
  ],
})

const tools = createAiSdkWorkPaperTools({
  workpaper,
  defaultReadRange: 'Summary!A1:B3',
  proofRange: 'Summary!A1:B3',
  writableSheets: ['Inputs'],
})
```

Keep model prompts separate from formula correctness. The deterministic proof
comes from the tool result.

## Limitation

The checked examples use `MockLanguageModelV3` and provider-free streams so the
tool contract is reproducible in CI. They prove the AI SDK tool boundary, not
the quality of a production model response or provider-specific retry behavior.

## When Not To Use Bilig

Do not wrap Bilig as an AI SDK tool for one-off arithmetic, for manual
spreadsheet editing, or for workbook files where Excel is allowed to calculate
later. Use it when the Node process must own the calculated answer before the
agent continues.

## Related

- [Agent WorkPaper proof matrix](agent-proof-matrix.md)
- [Agent framework spreadsheet tools](vercel-ai-sdk-langchain-spreadsheet-tool.md)
- [Workbook tools for agent frameworks](agent-framework-workbook-tools.md)
- [Agent WorkPaper tool-calling recipe](agent-workpaper-tool-calling-recipe.md)
