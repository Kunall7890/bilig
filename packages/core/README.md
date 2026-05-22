# @bilig/core

Local-first spreadsheet engine core for bilig with calc, selection, and commit semantics.

## Install

```bash
npm install @bilig/core
```

## Usage

```ts
import { SpreadsheetEngine, createSpreadsheetEngineFromDocument, exportSpreadsheetEngineDocument } from '@bilig/core'

const engine = new SpreadsheetEngine({ workbookName: 'Budget' })
await engine.ready()

engine.createSheet('Sheet1')
engine.setCellValue('Sheet1', 'A1', 10)
engine.setCellFormula('Sheet1', 'B1', 'A1*2')

const document = exportSpreadsheetEngineDocument(engine)

const restored = await createSpreadsheetEngineFromDocument(document, {
  replicaId: 'restored-budget',
})
```

## Agent workbook plans

`@bilig/core` can execute generic plans produced by `@bilig/workbook` without
hardcoded business models:

```ts
import { createWorkbookRunAdapter, SpreadsheetEngine } from '@bilig/core'
import { runWorkbookAction } from '@bilig/workbook'
import { model } from './model.js'

const engine = new SpreadsheetEngine({ workbookName: 'Agent workbook' })
await engine.ready()

const result = await runWorkbookAction(model, 'calculate', createWorkbookRunAdapter(engine))
```

`createWorkbookRunAdapter(engine)` materializes generic `plan.commands` into
engine operations, including range and table-column writes, falls back to
explicit `plan.ops` for low-level plans, reads single-cell `valueEquals` and
`formulaEquals` targets, and verifies generic `exists` and `noFormulaErrors`
checks. When undo is captured, the returned `undo.ops` are portable workbook
operations that can be applied by the runtime. Consumer-defined model semantics
stay in the model; the core adapter only handles workbook execution and proof.

## Persistence

`@bilig/core` exposes a canonical persistence surface for standalone engine usage:

- `exportSpreadsheetEngineDocument(engine)`
- `importSpreadsheetEngineDocument(engine, document)`
- `createSpreadsheetEngineFromDocument(document, options?)`
- `serializeSpreadsheetEngineDocument(document)`
- `parseSpreadsheetEngineDocument(json)`

The persisted document format stores:

- `snapshot`: workbook semantic state
- `replica`: optional replica/sync state

This makes `SpreadsheetEngine` usable without Zero or any bilig app runtime.

## Package entrypoints

- ESM: `./dist/index.js`
- Types: `./dist/index.d.ts`

This package is part of the [bilig](https://github.com/proompteng/bilig) monorepo.
