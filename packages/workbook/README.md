# @bilig/workbook

Agent-first workbook models and portable workbook intent for Bilig.

The bar is simple: an agent must love this library. It should be simple,
generic, predictable, inspectable, verifiable, and never depend on hardcoded
business models or human spreadsheet UI assumptions.

Use `@bilig/workbook` when you want a consumer-defined workbook model that an
agent can inspect, plan, verify, and hand to a runtime. The package does not
calculate formulas, open a grid, or ship built-in business templates. It gives
you a small public language for saying:

- what workbook structure to find;
- what action to take;
- what formulas or values to write;
- what checks prove the action worked;
- what portable operations a runtime should execute.

The consumer owns the model. Bilig owns the contracts.

## Install

```sh
pnpm add @bilig/workbook
```

Runtime dependencies are intentionally small:

- `@bilig/protocol`
- `@bilig/formula`

This package does not depend on `@bilig/core`, `@bilig/headless`,
`@bilig/agent-api`, `zod`, or `effect`.

## What It Is

`@bilig/workbook` is a public API package for building workbook intent. A model
is just a named contract with three parts:

- `find`: create refs to workbook structure.
- `checks`: declare proof the runtime must return.
- `actions`: build workbook commands from refs and optional JSON input.

The package returns frozen, JSON-safe plans. A runtime adapter executes those
plans with `@bilig/core`, `apps/bilig`, or consumer infrastructure.

```text
consumer model -> @bilig/workbook plan -> runtime adapter -> proof result
```

## What It Is Not

`@bilig/workbook` is not a workbook engine. It does not own workbook state,
calculate formulas, render cells, or infer business meaning from sheet names.

Do not add built-in revenue models, prepaid models, financial models, report
templates, or any other domain assumptions here. If a consumer needs those, they
define them in their own package using this generic API.

## Quick Start

```ts
import { defineModel, formula } from '@bilig/workbook'

export const model = defineModel({
  name: 'custom-calculation',
  description: 'Writes Result from Base and Rate.',

  find(workbook) {
    const table = workbook.findTable({
      headers: ['Base', 'Rate', 'Result'],
    })

    return {
      table,
      base: table.column('Base'),
      rate: table.column('Rate'),
      result: table.column('Result'),
    }
  },

  checks({ refs, workbook }) {
    return [workbook.check.exists(refs.table), workbook.check.noFormulaErrors(refs.result)]
  },

  actions: {
    calculate({ refs, workbook }) {
      workbook.writeFormula(refs.result, formula.multiply(refs.base, refs.rate))
      workbook.check.noFormulaErrors(refs.result)
    },
  },
})
```

This model does not know what the table represents. It only knows how to find
columns, write a formula, and ask the runtime to prove the result has no formula
errors.

## Agent Flow

Agents should use the package in this order:

1. Inspect the model with `describeModel(model)`.
2. Plan an action with `planWorkbookAction(model, actionName, input)`.
3. Build a command bundle with `buildWorkbookCommandBundle(plan, options)`, or
   use `planWorkbookCommand(model, actionName, input, options)` directly.
4. Read `describeCommandBundle(command)` for the runtime handoff contract.
5. Run `verifyWorkbookCommandBundle(command)` before execution.
6. Hand the command to a runtime adapter with `runWorkbookCommandBundle`.
7. Inspect `describeRunResult(result)`.
8. Inspect `result.receipt` when present for revisions, rendered proof, undo proof, and warnings.
9. Treat `status: "done"` as success only when checks are returned as proof.

The important rule: planned checks are promises, not proof. Runtime proof must
turn checks into `passed` or `failed`.

Runnable proof lives in
[`examples/workbook-agent-model`](../../examples/workbook-agent-model). It uses
`@bilig/workbook` for the public model contract and `@bilig/core` only as one
runtime adapter choice.

## Humans Should Know

Use this package when you are designing a stable workbook tool surface for
agents, services, or tests. The model is the documentation. Good models have:

- a clear `name`;
- action names that are normal verbs;
- `description` text when the action is not obvious;
- refs returned from `find` with business meaning owned by the consumer;
- checks that describe the exact proof expected after runtime execution.

Avoid hiding intent in comments, UI screenshots, or sheet-position assumptions.
Put the contract in the model.

## Refs

Refs are handles to workbook structure. They are stable enough to pass through a
plan, describe in logs, and map in a runtime adapter.

```ts
const input = workbook.findRange({ sheetName: 'Inputs', address: 'B2' })
const total = workbook.findName('Total')
const table = workbook.findTable({ name: 'Items' })
const amount = table.column('Amount')
```

Ref ids are opaque. Do not parse them. Use `label` for display and
`describeRef(ref)` for structured logging.

## Find API

Use the direct methods inside `find(workbook)`:

```ts
workbook.findTable({ name: 'Items' })
workbook.findTable({ sheetName: 'Data', headers: ['Kind', 'Value'] })
workbook.findColumn({ table, name: 'Value' })
workbook.findRange({ sheetName: 'Inputs', address: 'B2' })
workbook.findRange({ sheetName: 'Inputs', startAddress: 'B2', endAddress: 'D8' })
workbook.findName('Total')
workbook.findRows({
  table,
  where: { column: 'Status', op: 'eq', value: 'Active' },
})
```

Convenience methods keep common code short:

```ts
const table = workbook.findTable({ name: 'Items' })
const rows = workbook.findRows({
  table,
  where: { column: 'Status', op: 'eq', value: 'Active' },
})

return {
  rows,
  amount: rows.column('Amount'),
  total: table.column('Total'),
}
```

Selectors validate their inputs before runtime handoff. Empty names, empty
headers, unsupported row operators, invalid ranges, non-finite values, and
malformed objects fail while planning.

## Formulas

Use `formula.*` helpers so formula text and formula inputs stay separate.

```ts
formula.add(refs.left, refs.right)
formula.subtract(refs.total, refs.discount)
formula.multiply(refs.base, refs.rate)
formula.divide(refs.amount, 12)
formula.sum(refs.amount, refs.tax, refs.fee)
formula.call('ROUND', [refs.amount, 2])
formula.text('ready')
formula.raw('SUM(Items[Amount])', { inputs: [refs.amount] })
```

The rule is:

```text
@bilig/workbook creates formula expressions
@bilig/formula parses and normalizes formula language
@bilig/core calculates formulas
```

Formula expressions expose:

- `source`: parseable formula text without the leading `=`;
- `inputs`: the refs that a runtime adapter must resolve.

Formula helpers validate operands and declared raw-formula inputs at runtime.
Malformed refs fail before planning, so agents do not hand opaque bad
dependencies to a runtime adapter.

Runtime adapters materialize formula inputs by replacing whole tokens only. Text
inside a quoted string or a larger identifier is left alone.

## Actions

Actions build intent. They do not execute anything.

```ts
actions: {
  calculate({ refs, workbook }) {
    workbook.writeFormula(refs.result, formula.multiply(refs.base, refs.rate));
    workbook.check.noFormulaErrors(refs.result);
  },

  reset({ refs, workbook }) {
    workbook.clear(refs.result);
  },
}
```

Actions can also declare JSON input for agent tool calls:

```ts
actions: {
  setValue: {
    description: "Writes a provided value into the input ref.",
    input: {
      kind: "object",
      fields: {
        value: { kind: "number", required: true },
      },
    },
    run({ refs, workbook, input }) {
      if (typeof input !== "object" || input === null || Array.isArray(input)) {
        throw new Error("Input object is required");
      }
      const value = input.value;
      if (typeof value !== "number") {
        throw new Error("Input value must be a number");
      }
      workbook.writeValue(refs.input, value);
      workbook.check.valueEquals(refs.input, value);
    },
  },
}
```

When an object input declares `fields`, those fields are the whole accepted
shape. Unknown keys are rejected before model code runs, which catches typoed
agent tool arguments instead of silently carrying them into an action. Use
`{ kind: "object" }` without `fields`, or `{ kind: "json" }`, when the consumer
really wants an open payload.

Supported action commands:

- `workbook.writeFormula(ref, formulaExpression)`
- `workbook.writeValue(ref, jsonLiteral)`
- `workbook.format(ref, { style?, numberFormat? })`
- `workbook.clear(ref)`
- `workbook.addOp(op, { target?, message? })`

`addOp` is the escape hatch for the transport-neutral operation language. Prefer
the simple commands when they express the action.

## Checks

Checks are how a model states what must be proven.

```ts
workbook.check.exists(refs.table)
workbook.check.noFormulaErrors(refs.result)
workbook.check.valueEquals(refs.output, 42)
workbook.check.valuesEqual(refs.range, [
  [1, 2],
  [3, 4],
])
workbook.check.formulaEquals(refs.output, formula.add(refs.left, refs.right))
workbook.check.formulasEqual(refs.range, [['SUM(Items[Amount])', null]])
workbook.check.custom({
  kind: 'consumerInvariant',
  target: refs.table,
  refs: [refs.amount, refs.result],
  message: 'Consumer-defined invariant passed',
})
```

The runtime adapter either verifies checks itself or returns readbacks that
`@bilig/workbook` can compare against the expected values and formulas.
When a readback check is evaluated, the returned check keeps a simple `proof`
field with the value, values, formula, or formulas that were actually read. A
passed check is therefore inspectable evidence, not just a status flag.

For non-readback checks such as `exists`, `noFormulaErrors`, and custom
consumer invariants, `verifyChecks` may attach runtime evidence:

```ts
verifyChecks(checks) {
  return checks.map((check) => ({
    ...check,
    status: 'passed',
    proof: {
      kind: 'runtime',
      message: 'Runtime verifier confirmed the check',
      data: { check: check.kind },
    },
  }))
}
```

The verifier must return the same checks in the same order. It may change
`status` and add runtime proof, but it cannot change the target, message, refs,
expectation, or remove readback proof.

Readbacks can be scalar, matrix-shaped, or cell-level. Cell-level readbacks are
often easiest for agents to inspect because they keep the target, cell address,
value, and formula together:

```ts
read(targets) {
  return [
    {
      target: targets[0],
      cells: [
        { sheetName: 'Sheet1', address: 'B2', value: 12, formula: 'A2*B2' },
      ],
    },
  ]
}
```

For a range target, `@bilig/workbook` derives the expected value or formula
matrix from complete in-range cell readbacks.

## Planning

Use `planWorkbookAction` when an action name or input may come from an agent,
tool call, or user request.

```ts
import { describePlanResult, planWorkbookAction, verifyPlan } from '@bilig/workbook'
import { model } from './model.js'

const planned = planWorkbookAction(model, 'calculate')

console.log(describePlanResult(planned))

if (planned.status === 'planned') {
  const verification = verifyPlan(planned.plan)
  if (verification.status === 'invalid') {
    throw new Error(JSON.stringify(verification.issues))
  }
}
```

A successful plan contains:

- `modelName`
- `actionName`
- `input`
- `refs`
- `refsUsed`
- `commands`
- `ops`
- `changed`
- `checks`

Models and plans are frozen. They are meant to be inspected, passed to a
runtime, and verified, not mutated after creation.

## Command Bundles

Plans describe intent. Command bundles are the executable handoff object for an
agent or service runtime.

```ts
import { describeCommandBundle, planWorkbookCommand, runWorkbookCommandBundle, verifyWorkbookCommandBundle } from '@bilig/workbook'

const planned = planWorkbookCommand(model, 'calculate', undefined, {
  baseRevision: 'rev-42',
  idempotencyKey: 'agent-turn-123',
})

if (planned.status === 'planned') {
  console.log(describeCommandBundle(planned.command))

  const verification = verifyWorkbookCommandBundle(planned.command)
  if (verification.status === 'invalid') {
    throw new Error(JSON.stringify(verification.issues))
  }

  const result = await runWorkbookCommandBundle(planned.command, adapter)
}
```

A command bundle includes:

- `commandId`: deterministic id for the exact plan, optional base revision, and optional idempotency key.
- `idempotencyKey`: optional retry key supplied by the caller.
- `baseRevision`: optional runtime precondition supplied by the caller.
- `plan`: the frozen workbook action plan.
- `requirements`: the apply/read/verify checklist for the adapter.
- `verification`: static plan verification captured before runtime execution.

`verifyWorkbookCommandBundle` proves the bundle still matches its embedded
plan, requirements, verification, input, model name, action name, and command
id. If someone mutates the bundle between approval and execution, the command
runner fails before `adapter.apply` is called.

Adapters receive the command as the optional second or third argument to
`preview`, `apply`, `read`, and `verifyChecks`, so runtimes can enforce
revision checks, idempotency, locks, and audit logging without changing the
consumer model.

## Runtime Handoff

The runtime adapter owns execution. `@bilig/workbook` owns the contract.

```ts
import { runWorkbookAction } from '@bilig/workbook'

const result = await runWorkbookAction(model, 'calculate', {
  preview(plan) {
    return runtime.preview(plan)
  },
  apply(plan) {
    return runtime.apply(plan)
  },
  read(targets, plan) {
    return runtime.read(targets, plan)
  },
  verifyChecks(checks, plan) {
    return runtime.verifyChecks(checks, plan)
  },
})
```

Adapter methods:

- `preview(plan, command?)`: optional materialization step for inspection and approval.
- `apply(plan, command?)`: required execution step; may return runtime receipt proof.
- `read(targets, plan, command?)`: optional readback step for value and formula checks.
- `verifyChecks(checks, plan, command?)`: optional runtime-owned proof step.

`@bilig/core` provides the canonical engine adapter:

```ts
import { createWorkbookRunAdapter } from '@bilig/core'
import { runWorkbookAction } from '@bilig/workbook'

const result = await runWorkbookAction(model, 'calculate', createWorkbookRunAdapter(engine))
```

The core adapter returns generic runtime receipt proof for apply, synchronous
mutation propagation, and undo capture. It does not invent app revisions or
rendered proof; `apps/bilig` or a consumer runtime can add those when it owns
the lock, persistence, and rendered readback path.

`@bilig/workbook` does not import `@bilig/core`; consumers choose their runtime.

## Results

The public result shape is deliberately boring.

```ts
type WorkbookRunResult =
  | {
      status: 'done'
      changed: readonly WorkbookChangeSummary[]
      checks: readonly WorkbookCheckResult[]
      undo?: WorkbookUndoRef
      applied?: WorkbookAppliedSummary
      receipt?: WorkbookRunReceipt
    }
  | {
      status: 'failed'
      errors: readonly WorkbookRunError[]
      checks: readonly WorkbookCheckResult[]
      undo?: WorkbookUndoRef
      receipt?: WorkbookRunReceipt
    }
```

If apply succeeds and proof later fails, the failed result preserves `undo` when
the adapter supplied it.

Command-bundle runs and adapters that return receipt proof also include a
`receipt`. Receipts are for agents and audit logs: they tie the run back to the
command id, idempotency key, optional base revision, runtime-applied revision,
calculated revision, rendered revision, rendered diffs, proof entries, warnings,
and undo metadata.

```ts
type WorkbookRunReceipt = {
  commandId?: string
  idempotencyKey?: string
  modelName: string
  actionName: string
  baseRevision?: WorkbookRevision
  appliedRevision?: WorkbookRevision
  calculatedRevision?: WorkbookRevision
  renderedRevision?: WorkbookRevision
  previewed: boolean
  applied: boolean
  verified: boolean
  checkCount: number
  passedCheckCount: number
  failedCheckCount: number
  unverifiedCheckCount: number
  proof: readonly WorkbookReceiptProof[]
  warnings?: readonly string[]
  undo?: WorkbookUndoRef
}
```

Receipt proof kinds are deliberately generic: `preview`, `apply`,
`authoritativeReadback`, `renderedReadback`, `semanticReadback`,
`recalculation`, `undo`, `check`, and `custom`. The runtime can attach rendered
diffs and proof entries without importing `@bilig/core` into this package.

Readback checks include runtime evidence on the check itself:

```ts
{
  status: 'passed',
  kind: 'valueEquals',
  message: 'Sheet1!B2 equals 12',
  expectation: { kind: 'valueEquals', value: 12 },
  proof: { kind: 'value', value: 12 },
}
```

Runtime-owned checks can use generic runtime proof:

```ts
{
  status: 'passed',
  kind: 'consumerInvariant',
  message: 'Consumer invariant holds',
  proof: {
    kind: 'runtime',
    message: 'Runtime verifier confirmed the invariant',
    data: { verifier: 'core' },
  },
}
```

Model code cannot pre-fill `proof`; `verifyPlan` rejects planned checks that try
to carry runtime evidence before the adapter has run.

## Low-Level Ops

The transport-neutral operation language remains public:

- `WorkbookOp`
- `WorkbookTxn`
- `EngineOp`
- `WorkbookOpBatch`
- `EngineOpBatch`
- `isWorkbookOp`
- `isEngineOp`
- `isEngineOps`
- `isEngineOpBatch`

Use low-level ops when a consumer needs exact workbook mutation intent:

```ts
workbook.addOp(
  {
    kind: 'setCellValue',
    sheetName: 'Sheet1',
    address: 'B2',
    value: 42,
  },
  {
    target: refs.output,
    message: 'Seed output value',
  },
)
```

The guards validate literal values, formula text, cell addresses, ranges,
identifiers, enum values, and batch shape.

## Public API Map

Authoring:

- `defineModel`
- `formula`
- `find`, `findTable`, `findColumn`, `findRange`, `findName`, `findRows`
- `check`

Planning and inspection:

- `planWorkbookAction`
- `planWorkbookCommand`
- `buildWorkbookActionPlan`
- `buildWorkbookCommandBundle`
- `inspectModel`
- `describeModel`
- `describeRef`
- `describePlan`
- `describePlanResult`
- `describeCommandBundle`
- `describeRuntimeRequirements`
- `collectWorkbookRefs`

Verification and execution:

- `verifyModel`
- `verifyPlan`
- `verifyWorkbookCommandBundle`
- `runWorkbookAction`
- `runWorkbookPlan`
- `runWorkbookCommandBundle`
- `verifyWorkbookReadbacks`
- `describeRunResult`

Stable code lists and guards:

- `workbookPlanIssueCodes`
- `isWorkbookPlanIssueCode`
- `workbookReadbackIssueCodes`
- `isWorkbookReadbackIssueCode`
- `workbookRunErrorCodes`
- `isWorkbookRunErrorCode`
- `workbookCommandBundleIssueCodes`
- `isWorkbookCommandBundleIssueCode`

Primary types:

- `WorkbookModel`
- `WorkbookAction`
- `WorkbookActionPlan`
- `WorkbookActionPlanResult`
- `WorkbookCommandBundle`
- `WorkbookCommandBundleResult`
- `WorkbookRunAdapter`
- `WorkbookRunResult`
- `WorkbookCheckResult`
- `WorkbookRunError`
- `WorkbookRef`
- `WorkbookCheckProof`
- `WorkbookOp`
- `EngineOpBatch`

## Rules For This Package

- Keep it generic.
- Keep names normal.
- Keep dependencies small.
- Keep refs opaque.
- Keep checks explicit.
- Keep runtime execution outside this package.
- Keep formula calculation in `@bilig/core`.
- Keep parse and normalization in `@bilig/formula`.
- Keep consumer business meaning in consumer models.

## Development

```sh
pnpm --filter @bilig/workbook test
pnpm --filter @bilig/workbook build
pnpm typecheck
pnpm lint
```
