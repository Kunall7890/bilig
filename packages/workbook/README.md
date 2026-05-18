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
- `collectWorkbookRefs`
- `describeRef`
- `describePlan`
- `describePlanResult`
- `verifyPlan`
- `formula`
- `WorkbookModel`
- `WorkbookAction`
- `WorkbookActionPlanResult`
- `WorkbookRefDescription`
- `WorkbookActionPlanDescription`
- `WorkbookActionPlanResultDescription`
- `WorkbookPlanVerification`
- `WorkbookPlanIssue`
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

Action plans expose `refsUsed`, a flat deduped list of workbook refs found in
the consumer-defined `refs` object. Use `collectWorkbookRefs` directly when an
agent needs to inspect refs from any nested consumer shape.

Use `describeRef` and `describePlan` when an agent needs JSON-safe intent for
logs, comparisons, approvals, or runtime handoff. Descriptions keep the same
generic refs, commands, checks, changes, and ops, but omit consumer-private
`refs` object shape and helper functions such as `table.column()`.
Use `describePlanResult` when the same JSON-safe handoff is needed for either
planned or failed action planning.

Use `verifyPlan` before runtime handoff when an agent needs to prove a planned
action is internally consistent. It checks for unresolved refs, unparsable
formulas, duplicate resolved refs, and missing concrete ops for commands that
already target a known single cell.
