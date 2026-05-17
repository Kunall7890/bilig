# @bilig/workbook

Agent-first workbook model API and transport-neutral workbook operation language for bilig.

Use this package when a consumer needs to describe workbook work without taking a
dependency on the engine, app server, transport, or replica-state implementation.

The public surface stays generic:

- `defineModel`
- `buildWorkbookActionPlan`
- `formula`
- `WorkbookModel`
- `WorkbookAction`
- `WorkbookRunResult`
- `WorkbookCheckResult`

The low-level operation language remains available:

- `WorkbookOp`
- `WorkbookTxn`
- `EngineOp`
- `EngineOpBatch`
- `isEngineOpBatch`

Formula helpers create portable formula expressions with `@bilig/formula`.
Calculation and workbook execution stay in `@bilig/core` and the app runtime.
