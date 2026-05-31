---
title: Mastra WorkPaper spreadsheet tool
published: true
description: 'Use @bilig/workpaper as the workbook logic behind a Mastra createTool: read a range, write one input, and return formula readback.'
tags: mastra, createTool, spreadsheet, workpaper, typescript
canonical_url: https://proompteng.github.io/bilig/mastra-workpaper-spreadsheet-tool.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Mastra WorkPaper Spreadsheet Tool

If a Mastra agent needs spreadsheet math, keep the workbook code in ordinary
TypeScript. Mastra should get small tool wrappers: one tool reads a summary
range, and one tool writes a validated input cell and returns the formula
readback.

That keeps the agent boundary boring. `@bilig/workpaper` owns formulas,
serialization, and restore checks; `createTool` owns the schema and the tool
name the model sees.

## Run the real Mastra smoke

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig
pnpm --dir examples/mastra-workpaper-tool install --ignore-workspace --lockfile=false
pnpm --dir examples/mastra-workpaper-tool run smoke
```

The smoke uses real `@mastra/core` `createTool()` objects. It does not call a
model or require an API key. The local run invokes the tools, writes
`Inputs!B3 = 0.4`, reads recalculated formula values, serializes the WorkPaper,
restores it, and verifies the restored readback.

Passing output includes:

```json
{
  "apiShape": "Mastra createTool -> execute -> WorkPaper readback",
  "toolIds": ["read-workpaper-summary", "set-workpaper-input-cell"],
  "writeResult": {
    "editedCell": "Inputs!B3",
    "before": { "expectedArr": 60000 },
    "after": { "expectedArr": 96000 },
    "checks": {
      "formulasPersisted": true,
      "restoredMatchesAfter": true,
      "expectedArrChanged": true
    }
  }
}
```

Runnable source:
[`examples/mastra-workpaper-tool/src/mastra-workpaper-tool.ts`](../examples/mastra-workpaper-tool/src/mastra-workpaper-tool.ts).

Do not open an upstream Mastra PR unless a maintainer asks for one. This page
is the local proof lane first.

## Run the framework-shape adapter

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig
pnpm --dir examples/headless-workpaper install --ignore-workspace
pnpm --dir examples/headless-workpaper run agent:framework-adapters
```

The adapter lane returns tool IDs and a verified write result without installing
Mastra:

```json
{
  "toolIds": ["read-workpaper-summary", "set-workpaper-input-cell"],
  "writeResult": {
    "editedCell": "Inputs!B3",
    "checks": {
      "formulasPersisted": true,
      "restoredMatchesAfter": true,
      "expectedArrChanged": true
    }
  }
}
```

## Mastra shape

The real smoke and the framework-shape adapter both follow the
`createTool({ id, description, inputSchema, outputSchema, execute })` shape from
the Mastra docs:

```ts
export const setWorkPaperInputCell = createTool({
  id: 'set-workpaper-input-cell',
  description: 'Set one WorkPaper input cell and return formula readback.',
  inputSchema: setInputCellInputSchema,
  outputSchema: workPaperWriteOutputSchema,
  execute: async (input) => setWorkPaperInputCellInWorkbook(input),
})
```

Use a narrow input schema. For the demo, the write tool accepts only the
`Inputs` sheet and an A1-style address. That keeps an agent from treating the
workbook like an arbitrary mutation surface.

## What to copy

- Keep `@bilig/workpaper` WorkPaper construction in your application code.
- Validate tool arguments before writing.
- Return before/after formula readback, not just an "updated" message.
- Serialize and restore the WorkPaper document inside the tool result when the
  workflow depends on persistence.

Official Mastra reference: <https://mastra.ai/reference/tools/create-tool>.

Adapter source:
[`examples/headless-workpaper/agent-framework-adapters.ts`](../examples/headless-workpaper/agent-framework-adapters.ts).
