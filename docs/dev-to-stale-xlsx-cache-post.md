---
title: Your Node XLSX reader probably did not recalculate that formula
published: false
description: A short maintainer note about stale cached XLSX formula values, when recalculation-on-open is enough, and when a backend needs a real formula runtime.
tags: node, excel, xlsx, typescript
canonical_url: https://proompteng.github.io/bilig/stale-xlsx-formula-cache-node.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
---

I do not trust formula values read straight out of an edited `.xlsx` file
anymore.

The failure mode is boring:

1. a Node job opens an Excel file;
2. it changes an input cell;
3. it reads a formula cell;
4. the formula cell still has the old answer;
5. somebody opens the file in Excel, saves it, and suddenly the number is
   correct.

That does not mean the write failed. It usually means you changed the file, but
you did not run a calculation engine.

An `.xlsx` file can store formula text and a cached value from the last program
that saved the workbook. A lot of Node libraries are good at reading and writing
the file format. That is not the same thing as recalculating the workbook.

## The quick rule

If the file is just a report and a human will open it later, mark it to
recalculate on open and treat the cached values as stale.

If your API route, worker, billing job, quote checker, or approval flow needs
the answer now, calculate before you read the output cell.

That usually means one of these:

- Excel, LibreOffice, or Microsoft Graph if exact Excel behavior matters;
- `xlsx-calc` if its formula coverage fits your workbook;
- HyperFormula if you need a mature broad JavaScript formula engine;
- a smaller workbook-state runtime if your service owns the model and can test
  the formulas directly.

I maintain one of those smaller runtimes:
[`@bilig/headless`](https://www.npmjs.com/package/@bilig/headless). The narrow
thing it tries to do well is: keep workbook-shaped formula logic in Node, write
known inputs, read calculated outputs, and save the workbook state as JSON so
tests can restore it.

It is not Excel in Node. It does not run macros or cover every workbook in the
wild. The useful bit is the readback loop.

## What not to do

This shape is risky if you treat the final read as freshly calculated:

```ts
const workbook = XLSX.readFile('quote.xlsx')
const sheet = workbook.Sheets.Inputs

sheet.B2.v = 42_000
XLSX.writeFile(workbook, 'quote-edited.xlsx')

const reread = XLSX.readFile('quote-edited.xlsx')
console.log(reread.Sheets.Quote.B8.v)
```

That `Quote!B8` value can still be the cached result from before the edit.

## The safer shape

For a backend decision path, put a tiny adapter around the workbook:

- write only named input cells;
- read only named output cells;
- fixture-test unsupported formulas;
- persist the workbook state used for the decision;
- export `.xlsx` only when a human needs the artifact.

Here is the small local shape with `@bilig/headless`:

```ts
import { WorkPaper } from '@bilig/headless'

const workbook = new WorkPaper()
const inputs = workbook.addSheet('Inputs')
const quote = workbook.addSheet('Quote')

workbook.setCellContents({ sheet: inputs, row: 1, col: 1 }, 40_000)
workbook.setCellContents({ sheet: inputs, row: 2, col: 1 }, 0.2)
workbook.setCellContents({ sheet: quote, row: 1, col: 1 }, '=Inputs!B2*Inputs!B3')

const before = workbook.getCellValue({ sheet: quote, row: 1, col: 1 })
workbook.setCellContents({ sheet: inputs, row: 1, col: 1 }, 42_000)
const after = workbook.getCellValue({ sheet: quote, row: 1, col: 1 })

const restored = WorkPaper.buildFromSnapshot(workbook.exportSnapshot())
const restoredQuote = restored.getSheetId('Quote')
if (restoredQuote === undefined) throw new Error('Expected Quote sheet')

const afterRestore = restored.getCellValue({ sheet: restoredQuote, row: 1, col: 1 })
console.log({ before, after, afterRestore })
```

The restore check is the part I care about. If spreadsheet formulas are going
to influence a backend decision, I want a test that proves the value after the
state crosses a serialization boundary.

## Where I would still use Excel

Use Excel, LibreOffice, or Graph when the workbook depends on exact Excel
semantics, add-ins, macros, or weird file features.

Use a Node runtime only when the model is small enough and constrained enough
that you can fixture-test the formula surface you actually use.

I wrote the longer version here:
<https://proompteng.github.io/bilig/stale-xlsx-formula-cache-node.html>

If you have shipped this kind of Excel-backed backend workflow, I am interested
in the failure cases. The most useful feedback is not "nice project"; it is:
"this formula/file shape would make me reject it."

Repo:
<https://github.com/proompteng/bilig>
