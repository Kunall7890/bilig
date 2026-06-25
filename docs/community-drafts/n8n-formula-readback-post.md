# n8n Formula Readback Notes

Use this note when answering a specific n8n workflow thread about spreadsheet
formula recalculation or readback. Do not paste it as a launch announcement.

## Workflow Boundary

The common n8n spreadsheet boundary is not writing workbook-shaped data. It is
proving the formula result changed after the workflow edits an input.

The maintained example is here:

<https://github.com/proompteng/bilig/tree/main/examples/n8n-workpaper-formula-readback>

It uses a public Bilig WorkPaper endpoint to:

1. send one forecast input edit;
2. recalculate dependent formulas;
3. return before and after computed values;
4. export and restore WorkPaper JSON;
5. fail unless restored output still matches.

Endpoint:

```text
POST https://bilig.proompteng.ai/api/workpaper/n8n/forecast
```

Example request:

```json
{
  "sheetName": "Inputs",
  "address": "B3",
  "value": 0.4
}
```

Proof fields to mention:

```json
{
  "verified": true,
  "editedCell": "Inputs!B3",
  "checks": {
    "formulasPersisted": true,
    "restoredMatchesAfter": true,
    "computedOutputChanged": true
  }
}
```

The point is narrow: backend workflows and tool integrations need direct
formula readback before continuing. This is not a replacement for Excel as a
human spreadsheet UI.

## Posting Rule

Reply only in exact-match threads where the problem is formula recalculation,
cached formula values, or avoiding spreadsheet UI automation. Keep the answer
short and link the runnable workflow.
