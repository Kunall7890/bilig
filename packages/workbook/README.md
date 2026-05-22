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
3. Read `describePlanResult(planned)` for a compact explanation.
4. Run `verifyPlan(plan)` before execution.
5. Hand the plan to a runtime adapter with `runWorkbookPlan` or
   `runWorkbookAction`.
6. Inspect `describeRunResult(result)`.
7. Treat `status: "done"` as success only when checks are returned as proof.

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

- `preview(plan)`: optional materialization step for inspection and approval.
- `apply(plan)`: required execution step.
- `read(targets, plan)`: optional readback step for value and formula checks.
- `verifyChecks(checks, plan)`: optional runtime-owned proof step.

`@bilig/core` provides the canonical engine adapter:

```ts
import { createWorkbookRunAdapter } from '@bilig/core'
import { runWorkbookAction } from '@bilig/workbook'

const result = await runWorkbookAction(model, 'calculate', createWorkbookRunAdapter(engine))
```

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
    }
  | {
      status: 'failed'
      errors: readonly WorkbookRunError[]
      checks: readonly WorkbookCheckResult[]
      undo?: WorkbookUndoRef
    }
```

If apply succeeds and proof later fails, the failed result preserves `undo` when
the adapter supplied it.

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
- `buildWorkbookActionPlan`
- `inspectModel`
- `describeModel`
- `describeRef`
- `describePlan`
- `describePlanResult`
- `describeRuntimeRequirements`
- `collectWorkbookRefs`

Verification and execution:

- `verifyModel`
- `verifyPlan`
- `runWorkbookAction`
- `runWorkbookPlan`
- `verifyWorkbookReadbacks`
- `describeRunResult`

Stable code lists and guards:

- `workbookPlanIssueCodes`
- `isWorkbookPlanIssueCode`
- `workbookReadbackIssueCodes`
- `isWorkbookReadbackIssueCode`
- `workbookRunErrorCodes`
- `isWorkbookRunErrorCode`

Primary types:

- `WorkbookModel`
- `WorkbookAction`
- `WorkbookActionPlan`
- `WorkbookActionPlanResult`
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
