# bilig-workpaper

Bilig WorkPaper runtime for Node.js services, tool integrations, and verified formula readback.

This is the unscoped npm entrypoint for the Bilig headless runtime. Use it when business logic is easiest to review as workbook cells and formulas, but the calculation needs to run in a backend service, queue worker, serverless route, test, or tool integration.

## Install

```sh
npm install bilig-workpaper
```

## Use A WorkPaper In Node

```ts
import { buildA1WorkPaper } from 'bilig-workpaper'

const book = buildA1WorkPaper({
  Inputs: [
    ['Metric', 'Value'],
    ['Units', 40],
    ['Price', 1200],
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Revenue', '=Inputs!B2*Inputs!B3'],
  ],
})

const proof = book.editAndReadback('Inputs!B2', 48, {
  readbackRange: 'Summary!B2',
})

console.log({
  editedCell: proof.editedCell,
  after: proof.afterReadback.displayValues,
  afterRestore: proof.restoredReadback.displayValues,
  persistedDocumentBytes: proof.persistedDocumentBytes,
  verified: proof.verified,
})

book.dispose()
```

## XLSX Import And Export

```ts
import { WorkPaper } from 'bilig-workpaper'
import { exportXlsx, importXlsx } from 'bilig-workpaper/xlsx'
```

Use `xlsx-formula-recalc` when you only need to edit and recalculate XLSX files. Use `exceljs-formula-recalc` when you already use ExcelJS and need recalculated formula results after changing inputs.

## Tool Commands And Optional MCP

The npm tarball includes `AGENTS.md`, `SKILL.md`, and the same CLI entrypoints
as `@bilig/headless`, so tool hosts can inspect `node_modules/bilig-workpaper`
without discovering the scoped package first.

```ts
import { createWorkPaperMcpServer } from 'bilig-workpaper/mcp'
```

For package-owned proof commands, use:

```sh
npm exec --package bilig-workpaper -- bilig-evaluate --door workpaper-service --json
npm exec --package bilig-workpaper -- bilig-evaluate --door agent-mcp --json
npm exec --package bilig-workpaper -- bilig-agent-challenge
npm exec --package bilig-workpaper -- bilig-mcp-challenge
npm exec --package bilig-workpaper -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
```

`bilig-evaluate` prints a `bilig-evaluator.v1` object. The challenge commands
remain available for callers that already know the direct WorkPaper or MCP path
they need.

For a runnable starter project, use `npm create @bilig/workpaper`.

## Scope

Bilig is not a desktop Excel clone. It is a formula workbook runtime for service-owned calculations, JSON persistence, XLSX import/export, and verified readback. Unsupported Excel functions, external workbook links, macros, and volatile functions may need review.

Full docs: <https://proompteng.github.io/bilig/>
