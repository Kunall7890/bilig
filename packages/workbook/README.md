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
- `findTable`, `findColumn`, `findRange`, `findName`, `findRows`
- `find`
- `workbookRefKinds`
- `isWorkbookRefKind`
- `isWorkbookRef`
- `workbookRowOperators`
- `isWorkbookRowOperator`
- `check`
- `describeModel`
- `describeRef`
- `describePlan`
- `describePlanResult`
- `describeRunResult`
- `describeRuntimeRequirements`
- `verifyPlan`
- `verifyModel`
- `runWorkbookPlan`
- `runWorkbookAction`
- `verifyWorkbookReadbacks`
- `normalizeWorkbookActionInputDescription`
- `workbookActionInputDescriptionKinds`
- `isWorkbookActionInputDescriptionKind`
- `isWorkbookActionInputDescription`
- `isWorkbookActionInput`
- `builtInWorkbookCheckKinds`
- `isBuiltInWorkbookCheckKind`
- `workbookRunErrorCodes`
- `isWorkbookRunErrorCode`
- `formula`
- `workbook.addOp(op, { target?, message? })` inside model actions
- `WorkbookModel`
- `WorkbookAction`
- `WorkbookActionConfig`
- `WorkbookActionDefinition`
- `WorkbookActionContext`
- `WorkbookCheckContext`
- `WorkbookFindWorkbook`
- `WorkbookCheckWorkbook`
- `WorkbookActionWorkbook`
- `WorkbookModelWorkbook`
- `WorkbookFindNamespace`
- `WorkbookRef`
- `WorkbookRefKind`
- `WorkbookRangeRef`
- `WorkbookNameRef`
- `WorkbookTableRef`
- `WorkbookColumnRef`
- `WorkbookRowsRef`
- `WorkbookRowOperator`
- `WorkbookActionInput`
- `WorkbookActionInputDescription`
- `WorkbookActionInputDescriptionKind`
- `WorkbookActionInspection`
- `WorkbookAddOpOptions`
- `WorkbookActionPlanResult`
- `WorkbookModelDescription`
- `WorkbookRefDescription`
- `WorkbookActionPlanDescription`
- `WorkbookActionPlanResultDescription`
- `WorkbookRunResultDescription`
- `WorkbookUndoRefDescription`
- `WorkbookRuntimeRequirements`
- `WorkbookRuntimeRequirement`
- `WorkbookRuntimeCapability`
- `WorkbookPlanVerification`
- `WorkbookPlanIssue`
- `WorkbookModelVerification`
- `WorkbookModelActionVerification`
- `WorkbookModelVerificationOptions`
- `WorkbookRunAdapter`
- `WorkbookRunApplyResult`
- `WorkbookRunReadback`
- `WorkbookReadbackVerification`
- `WorkbookReadbackIssue`
- `WorkbookReadbackIssueCode`
- `WorkbookCheckExpectation`
- `WorkbookCheckExpectationDescription`
- `WorkbookBuiltInCheckKind`
- `WorkbookCustomCheckOptions`
- `WorkbookReadbackCheckOptions`
- `WorkbookRawFormulaOptions`
- `WorkbookRunResult`
- `WorkbookRunError`
- `WorkbookRunErrorCode`
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
`defineModel` returns frozen, normalized model metadata. Model and action names
must be non-empty and already trimmed, while descriptions and input metadata are
trimmed into model-owned frozen copies so the manifest an agent inspected cannot
be mutated later by the caller. The original consumer config remains
caller-owned data; `defineModel` does not freeze or rewrite it.
Action-object manifests only read own `run`, `description`, and `input`
properties. Prototype-inherited metadata is ignored, and an inherited `run`
function is rejected, so agent-visible manifests stay plain and explicit.
Actions can also accept a JSON-safe input:

```ts
import { defineModel } from '@bilig/workbook'

export const model = defineModel({
  name: 'custom-writer',

  find(workbook) {
    return {
      output: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
    }
  },

  actions: {
    write({ refs, workbook, input }) {
      if (typeof input !== 'object' || input === null || Array.isArray(input)) {
        throw new Error('input object required')
      }
      const value = input.value
      if (typeof value !== 'number') {
        throw new Error('numeric value required')
      }
      workbook.writeValue(refs.output, value)
    },
  },
})
```

`planWorkbookAction(model, "write", { value: 12 })` clones and canonicalizes
that input into the plan so an agent can inspect exactly what was requested.
Inputs must be plain JSON values: strings, finite numbers, booleans, `null`,
arrays without holes, and plain objects. This package intentionally does not add
schema dependencies; consumers own their own input validation inside actions.
Use `verifyModel(model, { inputs: { write: { value: 12 } } })` when whole-model
verification needs parameters for specific actions.

Actions can also be plain action objects when an agent needs a richer manifest
without running workbook code:

```ts
actions: {
  write: {
    description: 'Write a consumer-provided value',
    input: {
      kind: 'object',
      fields: {
        value: { kind: 'number', required: true },
      },
    },
    run({ refs, workbook, input }) {
      if (typeof input !== 'object' || input === null || Array.isArray(input)) {
        throw new Error('input object required')
      }
      const value = input.value
      if (typeof value !== 'number') {
        throw new Error('numeric value required')
      }
      workbook.writeValue(refs.output, value)
    },
  },
}
```

Input descriptions are intentionally descriptive metadata, not a schema engine.
They support boring JSON kinds: `json`, `object`, `array`, `string`, `number`,
`boolean`, and `null`. `object` descriptions may list sorted `fields`; `array`
descriptions may list `items`. `normalizeWorkbookActionInputDescription` trims
text, rejects malformed metadata, returns frozen data, and keeps the package
free of `zod`, `effect`, or model-specific validators.
`workbookActionInputDescriptionKinds`, `isWorkbookActionInputDescriptionKind`,
`isWorkbookActionInputDescription`, and `isWorkbookActionInput` expose the same
contract as stable data so generic tool builders can validate model metadata and
JSON-safe payloads without ad hoc string matching.

Formula expressions also keep their workbook inputs separate from their formula
text. A planned `writeFormula` command includes both the parseable formula
string and the generic model refs it used, so an agent can inspect what the
action depends on without reverse-parsing placeholder names.
For formulas outside the small helper set, use
`formula.raw(source, { inputs })`; the source stays parseable while the
declared refs remain inspectable and verifiable. These inputs are a declared
dependency contract for agents, not parser-discovered proof that every formula
reference has been mapped to a model ref.
Formula operands intentionally do not accept bare strings. Use `formula.raw`
for formula source and `formula.text` for spreadsheet string literals, so an
agent cannot confuse a label, cell address, named range, or user-provided string
with executable formula text.

Known single-cell `workbook.format(ref, { numberFormat })` actions compile to
concrete `setCellFormat` ops. Use `numberFormat: null` to plan an explicit
format clear. Style patches remain high-level intent until the runtime resolves
style ids.
Use `workbook.addOp(op, { target?, message? })` inside model actions when a
consumer needs the existing low-level workbook operation language directly. The
op is runtime-guarded with `isWorkbookOp`, cloned into `plan.ops`, and kept in
the command log so agents can inspect it without depending on `@bilig/core`.
When a `target` is supplied for an address or range op, `verifyPlan` checks that
the op touches the same range. For op kinds without an inferable range, `target`
is still useful for logs and approvals but cannot prove the affected cells by
itself.

Action plans expose `refsUsed`, a flat deduped list of workbook refs found in
the consumer-defined `refs` object. Use `collectWorkbookRefs` directly when an
agent needs to inspect refs from any nested consumer shape.
Use `findTable`, `findColumn`, `findRange`, `findName`, and `findRows` directly
when an agent or test needs the same generic refs outside a model callback. The
same helpers are also available as a frozen `find` namespace with short aliases
such as `find.table(...)`, `find.range(...)`, and `find.rows(...)`.
Use the frozen `workbookRefKinds` and `workbookRowOperators` lists, plus
`isWorkbookRefKind`, `isWorkbookRef`, and `isWorkbookRowOperator`, when a
generic agent tool needs to validate model refs or row predicates without
copying string unions or depending on a schema package.
Selector helpers trim text, canonicalize cell addresses, and reject empty table
names, column names, named ranges, headers, invalid range addresses, invalid row
operators, and non-finite row predicate values before a plan reaches runtime.
`findRows` refs include the predicate value in their stable id so two
consumer-defined row selectors do not collapse during dedupe, while labels stay
human-readable for agent logs.
Refs are frozen data objects. Ergonomic helpers like `table.column()` and
`rows.column()` remain available, but they are non-enumerable so JSON
inspection, object keys, and plan descriptions stay data-first.
For table-backed rows, use `rows.column("Amount")` to target only that column in
the matching rows. Runtime adapters can materialize row-filtered columns into
the exact cells to read, write, format, clear, or use as row-wise formula inputs
without hardcoding a business model.
Use `check.exists(ref)` and `check.noFormulaErrors(ref)` directly when an agent
or test needs the same generic planned checks outside a model callback.
Use `check.valueEquals(ref, value)` and `check.formulaEquals(ref, formula)` when
an action should carry machine-readable readback expectations for runtime
verification. `formulaEquals` stores normalized formula text plus explicit model
refs used by that formula, so an agent can inspect the post-action proof target
without depending on a rendered spreadsheet UI.
Use `check.custom({ kind, message, target, refs })` for consumer-defined
invariants; the package does not need to know what the model means. `target`
names the main ref, and `refs` names any supporting refs the invariant depends
on so agents can describe and verify the full check contract.
Custom check kinds cannot reuse built-in names. Use the frozen
`builtInWorkbookCheckKinds` list or `isBuiltInWorkbookCheckKind` guard when a
consumer-facing tool needs to validate check kinds before planning.

Model callback phases are deliberately scoped. `find(workbook)` receives only
the find API; `checks({ workbook })` receives find helpers plus `workbook.check`;
actions receive find helpers, checks, and mutation planning methods. That keeps
discovery and proof declaration separate from workbook mutation intent.

Use `describeModel` when an agent needs a JSON-safe manifest of model name,
model description, sorted action names, per-action descriptions, optional input
descriptions, and whether model-level checks exist without running `find`,
checks, or actions.
Use `describeRef` and `describePlan` when an agent needs JSON-safe intent for
logs, comparisons, approvals, or runtime handoff. Descriptions keep the same
generic action input, refs, commands, checks, changes, and ops, but omit
consumer-private `refs` object shape and helper functions such as
`table.column()`.
Plans are frozen handoff objects: action input, refs used, commands, concrete
ops, changed summaries, and checks cannot be rewritten after planning. That
lets an agent inspect a plan once and pass the same intent to an adapter without
caller-side metadata drift.
Use `describePlanResult` when the same JSON-safe handoff is needed for either
planned or failed action planning.
Use `describeRunResult` after execution when an agent needs the same JSON-safe
shape for `done` or `failed` run results. It preserves changed summaries,
checks, errors, and undo ops, but describes workbook refs without helper
functions such as `table.column()` or `rows.column()`.
Run errors use a stable `WorkbookRunErrorCode` union rather than arbitrary
strings. Use the frozen `workbookRunErrorCodes` list and `isWorkbookRunErrorCode`
guard when an adapter, logger, or approval layer needs to branch on known
failure classes. Runtime adapters should use `apply_failed` for apply
exceptions and `runtime_rejected` for intentional runtime refusal with a
specific message instead of inventing new public codes.
Use `describeRuntimeRequirements(plan)` before runtime handoff when an agent
needs to inspect what the adapter must do. It returns a JSON-safe list of
generic `apply`, `read`, and `verify` requirements with boring capabilities
such as `writeFormula`, `writeValue`, `format`, `clear`, `applyOp`, `read`, and
`verifyCheck`. Command-derived concrete single-cell ops are not repeated as
extra `applyOp` requirements, while explicit or manually assembled ops still
appear as `applyOp`. It does not execute anything and it does not import the
engine.

Use `verifyPlan` before runtime handoff when an agent needs to prove a planned
action is internally consistent. It checks for non-JSON-safe action input,
unresolved refs, unparsable formulas, duplicate resolved refs, and missing
concrete ops for write, clear, and number-format commands that already target a
known single cell. Custom check targets and supporting refs must also resolve
through the model's `refsUsed` contract. Formula readback expectation inputs
must also resolve through `refsUsed`, and expectation formulas must be parseable.
Checks must start as `planned`; consumer code cannot mark a check passed or
failed before runtime proof.
Low-level `addOp` commands must contain valid `WorkbookOp` values, must still
appear in `plan.ops`, and must match their declared `target` when the op exposes
a concrete address or range.
Use `verifyModel` to plan and verify every action in a consumer-defined model
with one JSON-safe result. Pass `inputs` when specific actions require
parameters. Each successfully planned action also includes its runtime
requirements, so an agent can inspect the action manifest, planned intent,
static verification result, and adapter handoff checklist from the same object.

Use `runWorkbookPlan(plan, adapter)` or
`runWorkbookAction(model, actionName, adapter, input)` when an agent needs a
generic apply-and-prove loop. The adapter owns runtime execution and semantic
readback:

```ts
const result = await runWorkbookAction(model, 'write', {
  apply(plan) {
    return runtime.apply(plan.ops)
  },
  read(targets) {
    return runtime.read(targets)
  },
})
```

`runWorkbookAction` plans the action, runs `verifyPlan`, calls
`adapter.apply(plan)`, then evaluates `valueEquals` and `formulaEquals` checks
against `adapter.read(targets, plan)`. It never imports the engine, headless
runtime, app server, or UI. If static verification fails, the apply adapter is
not called. If a readback expectation is missing or mismatched, the returned
`WorkbookRunResult` is `failed` with deterministic error codes such as
`readback_missing`, `value_mismatch`, or `formula_mismatch`. `adapter.read`
must return exactly the requested targets; extra readbacks fail with
`readback_unexpected` because an agent should not accept proof for cells it did
not ask the runtime to inspect.
Formula readbacks are exact: adapters should return formula text in the same
normalized no-leading-`=` form produced by `formula.source`.
`adapter.apply` only applies the plan and may return an undo ref; it cannot
drop, replace, or prove checks. An apply result with `status: "applied"` and
non-empty `errors` is rejected as `runtime_rejected`.
Adapters provide `verifyChecks(checks, plan)` to prove non-readback checks such
as `exists`, `noFormulaErrors`, and consumer-defined `custom` invariants. The
verifier must return the same checks in the same order and may only change
`status` or add JSON-safe `proof`. Changing `kind`, `target`, `refs`,
`expectation`, or `message` fails the run as `invalid_check_verification`.
Unsupported verifier fields are stripped from accepted results. If any verified
check is `failed`, the run returns `failed` with `check_failed`. If any check
remains `planned` after readback and adapter verification, the run returns
`failed` with `check_not_verified`; `status: "done"` means every planned check
has proof.

When running against Bilig's engine, use `createWorkbookRunAdapter(engine)` from
`@bilig/core`. The adapter materializes generic `plan.commands` into engine
operations, falls back to explicit `plan.ops` for low-level plans, reads
single-cell expectation targets, and verifies generic `exists` and
`noFormulaErrors` checks without adding workbook-specific business models to
this package. If the engine captures undo, the run result includes `undo.ops` in
the same portable operation language.
