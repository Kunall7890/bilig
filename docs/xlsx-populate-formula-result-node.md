---
title: xlsx-populate formula results in Node.js
published: true
description: How to handle calculated formula values when xlsx-populate writes formulas but a Node.js service needs fresh readback.
tags: typescript, node, xlsx-populate, xlsx, formulas, recalculation
canonical_url: https://proompteng.github.io/bilig/xlsx-populate-formula-result-node.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# xlsx-populate formula results in Node.js

`xlsx-populate` can write cells and formulas into XLSX files. It is not a
workbook calculation engine.

That matters when a backend flow needs both:

- formula text in the generated workbook; and
- a fresh calculated value immediately after changing inputs.

Those are different responsibilities. A file writer can serialize the formula.
A calculator has to evaluate the dependency graph.

## The common trap

Setting a formula does not mean the cached formula result changed:

```ts
cell.formula('A1*10')
```

Some XLSX libraries can serialize a `{ formula, result }` pair, but that still
requires you to compute the result somewhere. Setting `fullCalcOnLoad` can ask
Excel or LibreOffice to recalculate later; it does not give a Node API route a
fresh value now.

## Pick the boundary

Use `xlsx-populate` when the job is primarily XLSX file generation or editing.

Use Excel, LibreOffice, or Microsoft Graph when the result must match Excel and
the operational cost of a spreadsheet host is acceptable.

Use `@bilig/headless` when the workbook is service-owned business logic and the
backend needs:

- input writes through an API;
- formula recalculation in-process;
- output readback before returning a response;
- JSON persistence and restore proof;
- optional XLSX import/export at the edge.

## WorkPaper version of the flow

```ts
import { WorkPaper } from '@bilig/headless'

const workbook = new WorkPaper()
const sheet = workbook.addSheet('Quote')

workbook.setCellContents({ sheet, row: 0, col: 0 }, 42)
workbook.setCellContents({ sheet, row: 0, col: 1 }, '=A1*10')

const value = workbook.getCellDisplayValue({ sheet, row: 0, col: 1 })
const snapshot = workbook.exportSnapshot()
const restored = WorkPaper.buildFromSnapshot(snapshot)

try {
  const restoredSheet = restored.getSheetId('Quote')
  if (restoredSheet === undefined) {
    throw new Error('Missing Quote sheet after restore')
  }

  const restoredValue = restored.getCellDisplayValue({
    sheet: restoredSheet,
    row: 0,
    col: 1,
  })

  console.log({ value, restoredValue })
} finally {
  restored.dispose()
  workbook.dispose()
}
```

That shape avoids the "write a formula, then wait for Excel to populate the
cache" loop. The service owns the calculated state before it emits a response.

## Test with a reduced workbook

If an existing `xlsx-populate` pipeline has a real formula case, reduce it to a
public workbook fixture and run:

```sh
curl -fsSLo formula-clinic-report.ts \
  https://proompteng.github.io/bilig/formula-clinic-report.ts
npx tsx formula-clinic-report.ts ./reduced.xlsx \
  --cells "Quote!B1"
```

The report prints package version, imported sheets, formula samples, requested
readback, and a fixture checklist. It runs locally and does not upload workbook
contents.

## Related

- [Fix stale XLSX formula values in Node.js](stale-xlsx-formula-cache-node.md)
- [ExcelJS formula recalculation in Node.js](exceljs-formula-recalculation-node.md)
- [XLSX formula recalculation in Node.js](xlsx-formula-recalculation-node.md)
- [Formula bug clinic](formula-bug-clinic.md)

Source issue:
<https://github.com/dtjohnson/xlsx-populate/issues/265>
