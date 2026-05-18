# @bilig/workbook

Agent-first workbook model API and transport-neutral workbook operation language for bilig.

Build `@bilig/workbook` so an agent would love using it: simple, generic,
predictable, inspectable, verifiable, and never dependent on hardcoded business
models or human spreadsheet UI assumptions.

Use this package when a consumer needs to describe workbook work without taking a
dependency on the engine, app server, transport, or replica-state implementation.

The public surface stays generic:

- `defineModel`
- `buildWorkbookActionPlan`
- `planWorkbookAction`
- `inspectModel`
- `formula`
- `WorkbookModel`
- `WorkbookAction`
- `WorkbookActionPlanResult`
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
Use `planWorkbookAction` when an action name comes from an agent or user input;
it returns `planned` or structured `failed` results instead of requiring
exception control flow.

Formula expressions also keep their workbook inputs separate from their formula
text. A planned `writeFormula` command includes both the parseable formula
string and the generic model refs it used, so an agent can inspect what the
action depends on without reverse-parsing placeholder names.
