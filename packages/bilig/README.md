# bilig

Formula WorkPaper runtime for Node.js services: edit cells, recalculate formulas, read results, and persist JSON.

`bilig` is the unscoped npm entrypoint for the published `@bilig/headless` runtime. Use it when package search or install ergonomics matter, and use `@bilig/headless` when you want the scoped package name directly.

## Install

```bash
npm install bilig
```

## Usage

```ts
import { WorkPaper, exportWorkPaperDocument, serializeWorkPaperDocument } from 'bilig'

const paper = WorkPaper.buildFromSheets({
  Quote: [
    ['Metric', 'Value'],
    ['Seats', 25],
    ['ARR', '=B2*1200'],
  ],
})

const sheet = paper.getSheetId('Quote')
if (sheet === undefined) {
  throw new Error('Quote sheet was not created')
}

console.log(paper.getCellDisplayValue({ sheet, row: 2, col: 1 })) // 30000

const json = serializeWorkPaperDocument(exportWorkPaperDocument(paper, { includeConfig: true }))
```

## Subpaths

```ts
import { createWorkPaperMcpServer } from 'bilig/mcp'
import { exportXlsx, importXlsx } from 'bilig/xlsx'
```

## When to use Bilig

Use Bilig when a Node service or agent tool needs to own workbook state locally:

- pricing and quote calculators
- payout, commission, or billing checks
- spreadsheet-backed import validation
- formula-backed JSON documents
- XLSX formula recalculation before returning a result

For the full runtime docs, examples, and benchmark evidence, see the [Bilig docs](https://proompteng.github.io/bilig/).
