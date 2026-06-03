---
title: ExcelJS formula result not updating after Node edits
published: true
description: Diagnose the ExcelJS failure mode where formula cells keep stale cached results after a Node.js process edits workbook inputs.
tags: exceljs, node, xlsx, formula-recalculation, stale-cache
canonical_url: https://proompteng.github.io/bilig/exceljs-formula-result-not-updating-after-node-edits.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# ExcelJS Formula Result Not Updating After Node Edits

Use this page for the specific ExcelJS failure where a Node service edits input
cells, but formula cells still show old cached results in the same process.

ExcelJS is the right tool for many `.xlsx` file tasks. The boundary is
calculation: a file library can preserve formula records and cached results
without recalculating the dependency graph after your service changes an input.

ExcelJS documents formula cells as objects with `formula` and optional `result`
data. If your backend needs the fresh value before returning a response, add a
formula runtime at that boundary.

Official ExcelJS reference:

- <https://github.com/exceljs/exceljs#formula-value>

## Failure Mode

The workbook has a formula such as `Quote!B2 = Inputs!B2*Inputs!B3`, and ExcelJS
shows a cached formula result. Your service changes `Inputs!B3`, then reads
`Quote!B2` before Excel opens the file. The cached result can still be the old
number.

Setting `workbook.calcProperties.fullCalcOnLoad = true` is not enough for
in-process readback. It asks a spreadsheet app to recalculate later.

## One Command

Run the ExcelJS bridge demo:

```sh
npx --package @bilig/exceljs-formula-recalc exceljs-recalc --demo --json
```

Expected output includes:

```json
{
  "commandSucceeded": true,
  "recalculationCompleted": true,
  "expectedValueMatched": true,
  "reads": {
    "Summary!B2": {
      "value": 72000
    }
  }
}
```

For a source-level reproduction:

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig
npm --prefix examples/recalc-bridge-workflows install
npm --prefix examples/recalc-bridge-workflows run so:exceljs-44199441
```

That script mirrors the stale-result pattern from "Get computed value of Excel
sheet cell in Node.js": an input changes, ExcelJS still has the old cached
formula result, then `@bilig/exceljs-formula-recalc` verifies and patches the
fresh result.

## Minimal Bridge

Use ExcelJS for workbook files and Bilig only at the recalculation boundary:

```ts
import ExcelJS from 'exceljs'
import { recalculateExceljsWorkbook } from '@bilig/exceljs-formula-recalc'

const workbook = new ExcelJS.Workbook()
// build or load sheets here

const result = await recalculateExceljsWorkbook(workbook, {
  edits: [{ target: 'Inputs!B3', value: 0.4 }],
  reads: ['Summary!B2'],
})

console.log({
  readback: result.reads['Summary!B2'],
  workbookMutated: result.workbookMutated,
  warnings: result.warnings,
})
```

## Limitation

This bridge is for fresh formula readback after Node edits. It is not a
replacement for ExcelJS styling, workbook layout, images, tables, comments, or
file-generation features.

## When Not To Use Bilig

Do not use Bilig when Excel, LibreOffice, or another spreadsheet application
will open and calculate the workbook before any service decision depends on the
value. Do use it when the Node process must reject, persist, route, approve, or
answer based on the recalculated formula result.

## Related

- [ExcelJS formula recalculation in Node.js](exceljs-formula-recalculation-node.md)
- [XLSX formula recalculation in Node.js](xlsx-formula-recalculation-node.md)
- [SheetJS formula result not updating in Node.js](sheetjs-formula-result-not-updating-node.md)
- [Agent WorkPaper proof matrix](agent-proof-matrix.md)
