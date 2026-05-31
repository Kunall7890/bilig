# Mastra WorkPaper Tool

This example uses real `@mastra/core` `createTool()` objects around a Bilig
WorkPaper. It does not call a model or require an API key. The smoke script
invokes the tools locally and proves the workbook recalculates after a tool
write.

Run it from a cloned checkout:

```sh
pnpm --dir examples/mastra-workpaper-tool install --ignore-workspace --lockfile=false
pnpm --dir examples/mastra-workpaper-tool run smoke
```

Expected proof:

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

The useful contract is the tool result, not the agent text. Return the edited
cell, before and after formula values, formula contracts, and restore checks so
a Mastra workflow can trust the calculation without asking the model to do the
math.
