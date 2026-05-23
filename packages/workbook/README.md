# @bilig/workbook

Generic workbook intent for agents and runtimes.

Build `@bilig/workbook` so an agent would love using it: simple, generic,
predictable, inspectable, verifiable, and never dependent on hardcoded business
models or human spreadsheet UI assumptions.

Use this package when a consumer wants to define their own workbook model and
hand a runtime a portable plan. The package does not import the engine, start a
server, calculate formulas, ship revenue/quote/forecast models, or depend on
`zod`, `effect`, `@bilig/core`, `@bilig/headless`, or `@bilig/agent-api`.

```sh
pnpm add @bilig/workbook
```

## The Shape

```ts
import {
  defineModel,
  describePlan,
  describeRunResult,
  describeRuntimeRequirements,
  formula,
  planWorkbookAction,
  runWorkbookPlan,
  toPlanData,
  verifyPlan,
  verifyPlanData,
} from '@bilig/workbook'

export const model = defineModel({
  name: 'generic-row-calculator',

  find(workbook) {
    const table = workbook.findTable({
      headers: ['Item', 'Quantity', 'Rate', 'Status', 'Total'],
    })
    const rows = workbook.findRows({
      table,
      where: { column: 'Status', op: 'eq', value: 'ready' },
    })

    return {
      table,
      rows,
      quantity: rows.column('Quantity'),
      rate: rows.column('Rate'),
      total: rows.column('Total'),
    }
  },

  checks({ refs, workbook }) {
    return [workbook.check.exists(refs.table), workbook.check.noFormulaErrors(refs.total)]
  },

  actions: {
    recompute({ refs, workbook }) {
      const expected = formula.multiply(refs.quantity, refs.rate)
      workbook.writeFormula(refs.total, expected)
      workbook.check.formulaEquals(refs.total, expected)
    },
  },
})

const planned = planWorkbookAction(model, 'recompute')
if (planned.status === 'failed') throw new Error(planned.errors[0]?.message)

const staticProof = verifyPlan(planned.plan)
const requirements = describeRuntimeRequirements(planned.plan)
const planForLogs = describePlan(planned.plan)
const transportedPlan = JSON.parse(JSON.stringify(toPlanData(planned.plan)))
const transportProof = verifyPlanData(transportedPlan)

const result = await runWorkbookPlan(transportedPlan, adapter)
const resultForLogs = describeRunResult(result)
```

That is the core flow:

1. `defineModel` freezes a consumer-defined model.
2. `find` returns generic refs.
3. `checks` declares facts the runtime must prove.
4. An action builds workbook intent.
5. `verifyPlan` checks the plan without running an engine.
6. `describeRuntimeRequirements` tells an adapter what it must apply, read, and prove.
7. `toPlanData` makes the plan JSON-safe for handoff.
8. `runWorkbookPlan` applies either the in-memory plan or transported plan data through a runtime-owned adapter and returns a boring result with check proof and apply proof when the adapter provides it.

## Public Contract

The main API is intentionally small:

- model: `defineModel`, `inspectModel`, `planWorkbookAction`, `buildWorkbookActionPlan`
- selectors: `findTable`, `findColumn`, `findRange`, `findName`, `findRows`, `find`
- checks: `check.exists`, `check.noFormulaErrors`, `check.valueEquals`, `check.formulaEquals`, `check.custom`
- formulas: `formula.add`, `formula.subtract`, `formula.multiply`, `formula.divide`, `formula.sum`, `formula.call`, `formula.raw`, `formula.text`, `formula.labels`
- input: `checkInput`, `normalizeWorkbookActionInputDescription`
- proof: `verifyPlan`, `verifyModel`, `verifyWorkbookReadbacks`
- descriptions: `describeModel`, `describeRef`, `describePlan`, `describePlanResult`, `describeRuntimeRequirements`, `checkRuntimeRequirements`, `checkRuntimeAdapter`, `describeRunResult`
- transport data: `isWorkbookRefData`, `toWorkbookRefData`, `collectWorkbookRefData`, `hydrateWorkbookRef`, `hydrateWorkbookRefs`, `toPlanData`, `isPlanData`, `checkPlanData`, `hydratePlanData`, `verifyPlanData`
- runtime handoff: `runWorkbookPlan`, `runWorkbookAction`, `WorkbookRunAdapter`
- feature handoff: `defineWorkbookFeaturePlugin`, `checkWorkbookFeaturePlugin`, `checkWorkbookCommandRequest`, `normalizeWorkbookCommandRequest`, `checkWorkbookCommandReceipt`, `normalizeWorkbookCommandReceipt`, `workbookCommandReceiptOpsMatch`
- low-level language: `WorkbookOp`, `WorkbookTxn`, `EngineOp`, `EngineOpBatch`, `isEngineOpBatch`

Stable data helpers are exported for generic tool builders:

- `workbookRefKinds`, `isWorkbookRefKind`, `isWorkbookRef`
- `isWorkbookRefData`, `toWorkbookRefData`, `collectWorkbookRefData`, `hydrateWorkbookRef`, `hydrateWorkbookRefs`
- `isPlanData`, `checkPlanData`
- `workbookRowOperators`, `workbookRowOperatorValueTypes`, `isWorkbookRowOperator`, `isWorkbookRowValueCompatible`
- `builtInWorkbookCheckKinds`, `isBuiltInWorkbookCheckKind`
- `workbookActionInputDescriptionKinds`, `isWorkbookActionInputDescriptionKind`, `isWorkbookActionInputDescription`, `isWorkbookActionInput`, `checkInput`
- `workbookRuntimeRequirementKinds`, `isWorkbookRuntimeRequirementKind`, `workbookRuntimeCapabilities`, `isWorkbookRuntimeCapability`, `checkRuntimeRequirements`
- `workbookCommandCategories`, `isWorkbookCommandCategory`, `workbookCommandExecutionModes`, `isWorkbookCommandExecutionMode`, `workbookCommandReceiptStatuses`, `isWorkbookCommandReceiptStatus`
- `workbookProjectionInterceptorPoints`, `isWorkbookProjectionInterceptorPoint`, `workbookUiContributionSlots`, `isWorkbookUiContributionSlot`, `checkWorkbookCommandRequest`
- `workbookRunErrorCodes`, `isWorkbookRunErrorCode`

Model action manifests are frozen null-prototype maps. Consumers can use normal
business-agnostic action names, including names such as `toString` or
`constructor`, and `planWorkbookAction` only runs own actions from the manifest.
Prototype-inherited actions are ignored, so an agent can treat the action list as
the full executable surface.
Model config and action objects are read as data too: `defineModel` requires
own data properties for `actions` entries and for action-object `run`,
`description`, and `input`. Accessor-backed model metadata is rejected before any
getter can run.
`inspectModel` and `describeModel` use the same manifest boundary, so model
names, descriptions, action maps, and action metadata can be inspected without
triggering hidden getters.
`planWorkbookAction` also validates that boundary before reading action metadata
or running model code. Invalid manifests return a structured `invalid_model`
failure instead of making the agent catch an accessor side effect.

## Selectors

Selectors are not a human spreadsheet UI. They are stable intent for runtimes and
agents.

- `findTable({ headers })` means "find a table with all these headers." Header
  order is normalized, duplicate headers are rejected, and matching is
  case-sensitive after trimming.
- `findRows({ table, where })` means "find rows in this table matching this
  predicate." `eq` and `neq` accept any JSON literal; `contains` and
  `startsWith` accept strings; ordered comparisons accept numbers or strings.
- `findRange` is the escape hatch for an explicit range when the consumer really
  has one. It validates and canonicalizes addresses before runtime handoff.

Refs are frozen data. Helpers such as `table.column("Total")` and
`rows.column("Total")` are non-enumerable, so JSON descriptions stay data-first.
Use `toWorkbookRefData` or `describeRef` when a ref must cross a JSON boundary.
Use `hydrateWorkbookRef` or `hydrateWorkbookRefs` after transport to regain the
local helpers. `verifyPlanData(describePlan(plan))` checks transported plan data
without requiring the consumer's private `refs` object shape.
Ref collection and ref hydration only inspect enumerable own data properties.
Accessors are ignored instead of invoked, so hidden consumer getters cannot run
while an agent is planning, verifying, logging, or hydrating workbook intent.
Array entries follow the same rule, and ref cloning copies only known ref fields
instead of spreading extra enumerable properties.
Selector creation follows the same data boundary. `findTable`, `findColumn`,
`findRange`, and `findRows` read option objects, row predicates, and header
arrays through own data properties, rejecting accessor-backed fields before any
getter can run.

For full action handoff, use `toPlanData(plan)` before JSON transport. A runtime
can call `checkPlanData(data)` to get structured path-based issues before
hydration, call `hydratePlanData(data)` to regain frozen refs and helper
methods, or pass the data directly to `describeRuntimeRequirements(data)` and
`runWorkbookPlan(data, adapter)`. Invalid transported action input and check
proof keep nested JSON paths such as `input.rows[1]` and
`checks[0].proof.when`, so an agent can repair the exact payload field before
hydration. Plan-data guards only trust own payload fields; inherited
prototype fields never satisfy the transport contract. Transported plan arrays
must contain own enumerable data entries too; holes, non-enumerable entries, or
accessor-backed entries are rejected without running getters. The hydrated plan
exposes `refs: { refsUsed }` instead of the consumer's private model-shaped
`refs` object, so transported execution stays generic.

## Action Input

Action input is JSON-safe data, not a schema-framework object. Action metadata
can describe generic input with `json`, `object`, `array`, `string`, `number`,
`boolean`, and `null` kinds. `checkInput(description, value)` returns a plain
`{ status, input, issues }` result so an agent can reject malformed tool payloads
before running workbook model code. Omitted input is valid unless the top-level
description sets `required: true`, so agents can distinguish an optional payload
from a malformed payload. `planWorkbookAction` uses the same check when an action
declares input metadata and preserves each failed input issue as a run error
with `path` and `issueCode`, so agents can branch without parsing messages.
JSON-safety failures keep the nested offending path too, such as
`input.items[2].amount`. Normalized payloads preserve consumer-owned JSON keys
as data, including names like `__proto__` and `constructor`, instead of letting
them affect object prototypes.
Action input payloads and input-description metadata must be enumerable own data
properties. Accessors are rejected without invoking them, so tool payload
validation cannot run hidden consumer code while an agent is planning.

## Formulas

`@bilig/workbook` creates formula expressions. `@bilig/formula` parses and
normalizes formula text. `@bilig/core` or an app runtime calculates it.

Formula helpers keep formula text, workbook dependencies, and formula labels
separate. A planned formula write includes the formula string, the refs used to
build it, and a `labels` array mapping each formula token to the workbook ref it
represents. Runtime adapters use those labels to materialize table columns,
filtered rows, names, and ranges without reverse-engineering hidden JS helpers.
For custom formula text, use `formula.raw(source, { inputs })`; pass
`labels: [{ name, ref }]` when the raw formula uses custom tokens. For
spreadsheet string literals, use `formula.text(value)`. Bare strings are not
formula operands because agents should not guess whether a string is code, a
label, a named range, or user text.

## Runtime Adapter

`@bilig/workbook` does not execute plans. A runtime owns that:

```ts
const adapter = {
  apply(plan) {
    const ops = materializeForThisRuntime(plan)
    return {
      status: 'applied',
      previewOps: ops,
      appliedOps: ops,
      proof: { source: 'runtime', opCount: ops.length },
      undo: { id: 'undo-1' },
    }
  },
  read(targets, plan) {
    return targets.map((target) => ({ target, value: 12 }))
  },
  verifyChecks(checks, plan) {
    return checks.map((entry) => ({ ...entry, status: 'passed' }))
  },
}
```

`runWorkbookPlan` accepts either a live plan or transported plan data and
refuses to call `apply` if static plan verification fails or if the adapter is
missing a required method. Use `checkRuntimeRequirements(data)` when runtime
requirements crossed a JSON boundary and an agent needs path-based diagnostics
before trusting the handoff. Runtime requirement arrays and nested ref arrays
must be own enumerable data entries; holes, non-enumerable entries, or
accessor-backed entries are rejected without running getters. Use
`checkRuntimeAdapter(planOrRequirements, adapter)` when an agent wants to check
`apply`, `read`, and `verifyChecks` coverage before calling the runtime.
Check-only plans do not require `apply`; when runtime requirements
contain only `read` or `verifyCheck`, `runWorkbookPlan` skips mutation and
verifies the declared checks directly.
If an adapter returns both `previewOps` and `appliedOps`, the result reports
whether they matched. If the adapter returns neither, the run records an
unverified apply fact. Use `runWorkbookPlan(plan, adapter, { requireApplyProof:
true })` when an agent must fail closed instead of accepting an unproved apply.
Runtime apply results, undo refs, apply errors, and check verifier output are
validated from own fields only; prototype-inherited fields are ignored before
they can become run proof. Adapter-returned ops and verifier proof must be data
properties, including non-enumerable guard fields such as `kind`. Runtime
evidence arrays must contain own enumerable data entries; holes,
non-enumerable entries, and accessors are rejected before any getter can run
during validation, cloning, or preview/apply comparison.
Readback checks attach proof to passed checks, such as
`{ source: "readback", value: 12 }` or
`{ source: "readback", formula: "(Table[Quantity])*(Table[Rate])" }`.
Generic check verifiers may only change `status` or add JSON-safe `proof`; they
cannot rewrite the check contract.
Consumer `checks()` return values are treated as model-output data too: returned
check arrays must contain own enumerable data entries, and returned check fields
must be own data properties. Accessor-backed or sparse returned checks fail
planning without running hidden getters.
If runtime apply succeeds but readback or check proof fails, the failed result
still carries `changed` and `undo` when the adapter returned applied ops or undo
metadata. A failed result before apply, or a failed apply that reports
`appliedOps: []` without undo metadata, uses `changed: []`.

The result is deliberately plain:

```ts
type WorkbookRunResult =
  | {
      status: 'done'
      apply?: WorkbookRunApplySummary
      changed: WorkbookChangeSummary[]
      checks: WorkbookCheckResult[]
      undo?: WorkbookUndoRef
      unverified?: WorkbookRunUnverified[]
    }
  | {
      status: 'failed'
      errors: WorkbookRunError[]
      apply?: WorkbookRunApplySummary
      changed: WorkbookChangeSummary[]
      checks: WorkbookCheckResult[]
      undo?: WorkbookUndoRef
      unverified?: WorkbookRunUnverified[]
    }
```

## Feature Handoff

Feature command requests are plain data for runtimes that expose workbook
features to agents. Use `checkWorkbookFeaturePlugin(data)` before registering
consumer-provided feature metadata. It returns stable path issues for commands,
projection interceptors, UI contributions, dependencies, lifecycle hooks, and
nested command input-description or UI metadata fields.

Use `checkWorkbookCommandRequest(data)` before dispatching a transported
request. It returns stable path issues such as `featureId`, `commandId`,
`category`, `mode`, and nested input paths like `input.rows[1]`, and
`normalizeWorkbookCommandRequest` returns the frozen request data for the
runtime. The exported command category, execution-mode, receipt-status,
projection-point, and UI-slot lists let tool builders present and validate
command contracts without importing a schema framework.

Use `checkWorkbookCommandReceipt(data)` before trusting runtime command evidence.
It returns the same boring `{ status, issues }` shape for receipt fields such as
`status`, `featureId`, `commandId`, `previewOps`, `appliedOps`, `undo`,
`changedRanges`, `proof`, `metadata`, and `errors`. Feature manifests, command
requests, and command receipts are validated from own payload fields only;
prototype-inherited fields are ignored. Receipt ops are frozen after
normalization, changed ranges must be own-field data, and manifest or receipt
arrays must contain own enumerable data entries. Holes, non-enumerable entries,
and accessor-backed ops, undo ops, ranges, or errors are rejected before any
getter can run.
`workbookCommandReceiptOpsMatch` uses canonical op equality instead of object
property order and refuses accessor-backed proof data.

## Low-Level Ops

Most models should use the small action API: `writeFormula`, `writeValue`,
`format`, `clear`, and checks. If a consumer needs the existing workbook
operation language directly, call `workbook.addOp(op, { target, message })`
inside an action.

The op is guarded with `isWorkbookOp`, cloned into `plan.ops`, and kept in the
command log. If a target is supplied and the op exposes a concrete range,
`verifyPlan` checks that the target and op agree.
Low-level op guards accept plain own-field payloads only. Prototype-inherited
op fields, nested ranges, and batch clocks are ignored so transported ops cannot
smuggle proof through object prototypes. Accessor-backed required fields,
nested fields, and op-array entries are rejected from descriptors without
running getters.

## Example

See [examples/workbook-agent-model](../../examples/workbook-agent-model) for a
generic model that plans, verifies, describes, runs, and prints proof without
depending on a hardcoded business model.
