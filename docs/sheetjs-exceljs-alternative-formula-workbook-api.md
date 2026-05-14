---
title: SheetJS and ExcelJS alternative for formula-backed workbook APIs
published: true
description: Decide when SheetJS, ExcelJS, or @bilig/headless is the right fit for XLSX files, formula records, and verified Node.js workbook execution.
tags: typescript, node, spreadsheet, formulas, xlsx, opensource
canonical_url: https://proompteng.github.io/bilig/sheetjs-exceljs-alternative-formula-workbook-api.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# SheetJS and ExcelJS alternative for formula-backed workbook APIs

Status: public comparison guide for developers evaluating spreadsheet
automation libraries.

Research date: 2026-05-14.

If you are searching for a SheetJS alternative or ExcelJS alternative, do not
start with a package name. Start with the job:

- read and write workbook files
- generate an XLSX report for Excel to open later
- keep formula text and cached or supplied formula results in the file
- run a workbook inside a Node.js service, edit inputs, and read the new result

SheetJS and ExcelJS are strong tools for workbook-file workflows. `bilig` is
not trying to replace that whole layer. The useful Bilig slice is narrower:
`@bilig/headless` gives a Node service or coding agent a WorkPaper object it
can build, mutate, evaluate, persist, restore, and verify without opening Excel
or a browser grid.

## Short Version

Use SheetJS when you need broad spreadsheet-file parsing and export across
formats.

Use ExcelJS when you need to create or edit XLSX workbooks with workbook-file
features such as sheets, rows, styles, and formula records.

Use `@bilig/headless` when the service must mutate a formula-backed workbook
and read the recalculated values back in the same process.

That last sentence is the boundary. If a backend only needs a file, stay with a
file library. If the backend needs the answer now, give it a workbook runtime.

For the broader library choice, start with the
[headless spreadsheet engine use-case chooser](headless-spreadsheet-engine-comparison.md#use-case-chooser).

## The Boundary That Matters

SheetJS Community Edition stores cell formulas in the `f` field and cell values
in the `v` field. Its formula docs explain that, when actual results are needed
in JavaScript, SheetJS Pro has a formula calculator component.

ExcelJS can store formulas and supplied results, but its public package docs say
it cannot process a formula to generate a result.

Those are reasonable design choices for file-centric libraries. They become a
problem only when your app needs to change an input, recalculate dependent
cells, and reject a workflow when computed readback does not match.

That is the place to evaluate `@bilig/headless`.

## Comparison Table

| Need | Start with | Reason |
| --- | --- | --- |
| Parse many spreadsheet file formats into JavaScript data | SheetJS | It is built around file-format import/export and a common spreadsheet object model. |
| Generate XLSX reports with workbook structure and styling | ExcelJS | It focuses on reading, manipulating, and writing XLSX workbook files. |
| Store formulas in a workbook file and let Excel calculate later | SheetJS or ExcelJS | Both can represent formula text and cached or supplied values in workbook data. |
| Recalculate formulas inside a Node service after changing inputs | `@bilig/headless` | It exposes a WorkPaper runtime with formula readback after edits. |
| Give a coding agent a spreadsheet tool it can mutate and verify | `@bilig/headless` | The maintained examples prove writeback, dependent formulas, persistence, and restore. |

## TypeScript Evaluation Path

Install the package in a scratch project:

```sh
mkdir bilig-headless-eval
cd bilig-headless-eval
npm init -y
npm pkg set type=module
npm install @bilig/headless
npm install -D tsx typescript @types/node
```

Create `workbook-runtime-check.ts`:

```ts
import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/headless'

type NumericCell = {
  value: number
}

function readNumber(cell: unknown, label: string): number {
  if (typeof cell === 'object' && cell !== null && typeof (cell as NumericCell).value === 'number') {
    return (cell as NumericCell).value
  }

  throw new Error(`Expected ${label} to be numeric, got ${JSON.stringify(cell)}`)
}

const workbook = WorkPaper.buildFromSheets({
  Inputs: [
    ['Metric', 'Value'],
    ['Customers', 32],
    ['ARPA', 1200],
    ['Discount', 0.04],
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Net revenue', '=Inputs!B2*Inputs!B3*(1-Inputs!B4)'],
  ],
})

const inputs = workbook.getSheetId('Inputs')
const summary = workbook.getSheetId('Summary')
if (inputs === undefined || summary === undefined) {
  throw new Error('Expected Inputs and Summary sheets')
}

const revenue = { sheet: summary, row: 1, col: 1 }
const before = readNumber(workbook.getCellValue(revenue), 'before revenue')

workbook.setCellContents({ sheet: inputs, row: 1, col: 1 }, 40)
const after = readNumber(workbook.getCellValue(revenue), 'after revenue')

const saved = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))
const restored = createWorkPaperFromDocument(parseWorkPaperDocument(saved))
const restoredSummary = restored.getSheetId('Summary')
if (restoredSummary === undefined) {
  throw new Error('Expected restored Summary sheet')
}

const afterRestore = readNumber(restored.getCellValue({ sheet: restoredSummary, row: 1, col: 1 }), 'restored revenue')

console.log({
  before,
  after,
  afterRestore,
  verified: before === 36864 && after === 46080 && afterRestore === after,
})
```

Run it:

```sh
npx tsx workbook-runtime-check.ts
```

Expected output:

```json
{ "before": 36864, "after": 46080, "afterRestore": 46080, "verified": true }
```

That check is intentionally small. It proves the part that file libraries do
not try to own: a Node process changed an input, read a dependent formula value,
serialized the workbook document, restored it, and read the same calculated
value again.

The maintained repository example adds more workflows:

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig/examples/headless-workpaper
npm install
npm run agent:tool-call
npm run agent:verify
```

The agent tool-call loop changes input cells, reads dependent formula outputs,
persists the workbook, restores it, and fails if the restored formulas or
values do not match.

## When To Combine The Tools

Use file libraries at the boundary and Bilig for the runtime model:

1. Use SheetJS or ExcelJS where the product is an `.xlsx` file.
2. Use `@bilig/headless` where the product is trusted computed workbook state.
3. Keep compatibility tests around the boundary so import/export and formula
   runtime behavior are not confused.

This is the honest architecture for many services. File libraries are still
useful. Bilig earns its keep when the service needs an auditable workbook-state
transition, not just a generated spreadsheet file.

## When Not To Choose Bilig

Do not choose Bilig first if the main requirement is broad XLSX styling,
images, charts, pivot tables, or complete Excel compatibility.

Do not choose it if a cached formula result is enough and Excel can calculate
later.

Do not choose it if the workload needs a mature commercial spreadsheet-file
support channel today.

## Related Proof

- [`docs/headless-spreadsheet-engine-comparison.md`](headless-spreadsheet-engine-comparison.md)
- [`docs/agent-spreadsheet-tool-call-loop.md`](agent-spreadsheet-tool-call-loop.md)
- [`docs/persisting-formula-backed-workpaper-documents-in-node.md`](persisting-formula-backed-workpaper-documents-in-node.md)
- [`docs/where-bilig-is-not-excel-compatible-yet.md`](where-bilig-is-not-excel-compatible-yet.md)
- [`examples/headless-workpaper`](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper)

## Sources

- SheetJS Cell Objects:
  <https://docs.sheetjs.com/docs/csf/cell/>
- SheetJS Formulae:
  <https://docs.sheetjs.com/docs/csf/features/formulae>
- SheetJS Parse Options:
  <https://docs.sheetjs.com/docs/api/parse-options>
- ExcelJS package docs:
  <https://www.npmjs.com/package/exceljs>
