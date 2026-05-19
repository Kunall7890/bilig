---
title: OpenAI Agents SDK WorkPaper tools
published: true
description: Wrap @bilig/headless workbook reads and verified edits as OpenAI Agents SDK function tools.
tags: openai agents sdk, tool calling, spreadsheet, workbook, typescript
canonical_url: https://proompteng.github.io/bilig/openai-agents-sdk-workpaper-tool.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# OpenAI Agents SDK WorkPaper Tools

Use this path when an OpenAI Agents SDK app needs a workbook tool it can call
from Node without opening Excel, LibreOffice, Google Sheets, or a screenshot UI.
The agent gets two ordinary function tools:

- `read_workpaper_summary` reads computed WorkPaper values and serialized cells.
- `set_workpaper_input_cell` writes one validated input cell and returns
  before/after readback, formula persistence checks, and restored JSON proof.

The maintained smoke script is provider-free by default. It imports
`Agent`, `tool()`, `RunContext`, and `invokeFunctionTool()` from
`@openai/agents`, creates a real SDK agent and function tools, then invokes the
tools locally so the read/write contract can run in CI without an API key:

```sh
pnpm --dir examples/headless-workpaper run agent:openai-agents-sdk
```

The OpenAI Agents SDK documents function tools as local functions wrapped with a
schema through `tool()`, and the same tools can be attached to an `Agent`:
<https://openai.github.io/openai-agents-js/guides/tools/>.

## Minimal Tool Shape

```ts
import { Agent, RunContext, invokeFunctionTool, tool } from '@openai/agents'
import { z } from 'zod'
import { WorkPaper } from '@bilig/headless'

const workbook = WorkPaper.buildFromSheets({
  Inputs: [
    ['Metric', 'Value'],
    ['Qualified opportunities', 20],
    ['Win rate', 0.25],
    ['Average ARR', 12000],
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Expected ARR', '=Inputs!B2*Inputs!B3*Inputs!B4'],
  ],
})

const setWorkPaperInputCell = tool({
  name: 'set_workpaper_input_cell',
  description: 'Set one validated WorkPaper input cell and return formula readback.',
  parameters: z.object({
    sheetName: z.literal('Inputs'),
    address: z.string().regex(/^[A-Z]+[1-9][0-9]*$/),
    value: z.union([z.string(), z.number(), z.boolean(), z.null()]),
  }),
  execute: async ({ sheetName, address, value }) => {
    const sheet = workbook.getSheetId(sheetName)
    const summarySheet = workbook.getSheetId('Summary')
    if (sheet === undefined) {
      throw new Error(`Unknown sheet: ${sheetName}`)
    }
    if (summarySheet === undefined) {
      throw new Error('Summary sheet is missing')
    }
    const cell = workbook.simpleCellAddressFromString(address, sheet)
    const summaryRange = workbook.simpleCellRangeFromString('Summary!A1:B2', summarySheet)
    if (cell === undefined) {
      throw new Error(`Invalid cell: ${sheetName}!${address}`)
    }
    if (summaryRange === undefined) {
      throw new Error('Summary range is invalid')
    }

    const before = workbook.getRangeValues(summaryRange)
    workbook.setCellContents(cell, value)

    return {
      editedCell: `${sheetName}!${address}`,
      before,
      after: workbook.getRangeValues(summaryRange),
    }
  },
})

const agent = new Agent({
  name: 'WorkPaper verification agent',
  instructions: 'Use WorkPaper tools and answer only from computed readback.',
  tools: [setWorkPaperInputCell],
})

const result = await invokeFunctionTool({
  tool: setWorkPaperInputCell,
  runContext: new RunContext(),
  input: JSON.stringify({
    sheetName: 'Inputs',
    address: 'B3',
    value: 0.4,
  }),
})

console.log(agent.name, result)
```

For a production adapter, use the full example instead of this short snippet:
[`examples/headless-workpaper/openai-agents-sdk-tool-smoke.ts`](../examples/headless-workpaper/openai-agents-sdk-tool-smoke.ts).
It also verifies persisted formulas by exporting a WorkPaper document, restoring
it, and comparing the computed readback after restore.

## Expected Proof

The smoke output includes this shape:

```json
{
  "apiShape": "OpenAI Agents SDK Agent -> tool() -> invokeFunctionTool()",
  "package": "@openai/agents",
  "agentName": "WorkPaper verification agent",
  "toolNames": ["read_workpaper_summary", "set_workpaper_input_cell"],
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

Keep the workbook mutation closed-world: validate sheet names and A1 addresses,
write one input at a time, recalculate through WorkPaper, return computed
readback, and persist only after the verification passes.
