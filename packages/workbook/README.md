# @bilig/workbook

Generic workbook intent for agents and runtimes.

Build `@bilig/workbook` so an agent would love using it: simple, generic,
predictable, inspectable, verifiable, and never dependent on hardcoded business
models or human spreadsheet UI assumptions.

Use this package when a consumer wants to define their own workbook model and
hand a runtime a portable plan. Bilig supplies the generic model API, selectors,
formula helpers, checks, JSON-safe transport data, validators, and run-result
proof shapes. It does not import an engine, start a server, calculate formulas,
ship business templates, or depend on `@bilig/core`, `@bilig/headless`,
`@bilig/agent-api`, `zod`, or `effect`.

```sh
pnpm add @bilig/workbook
```

## Use These First

Most consumers should start with only these names:

- `defineModel`
- `formula`
- `prepareWorkbookAction`
- `runWorkbookPlan`
- `describeModel`, `describePlan`, `describeRunResult`

That path lets an agent define intent, inspect it before execution, transport it
as plain data, run it through a runtime-owned adapter, and verify the returned
proof without knowing anything about a rendered spreadsheet UI.

## The Shape

```ts
import { defineModel, describeRunResult, formula, prepareWorkbookAction, runWorkbookPlan } from '@bilig/workbook'

export const model = defineModel({
  name: 'named-range-formula',

  find(workbook) {
    return {
      input: workbook.findName('input'),
      factor: workbook.findName('factor'),
      result: workbook.findName('result'),
    }
  },

  checks({ refs, workbook }) {
    return [workbook.check.exists(refs.result), workbook.check.noFormulaErrors(refs.result)]
  },

  actions: {
    calculate({ refs, workbook }) {
      const expected = formula.multiply(refs.input, refs.factor)
      workbook.writeFormula(refs.result, expected)
      workbook.check.formulaEquals(refs.result, expected)
    },
  },
})

const prepared = prepareWorkbookAction(model, 'calculate')
if (prepared.status === 'failed') throw new Error(prepared.errors[0]?.message)

const result = await runWorkbookPlan(prepared.planData, adapter, { strict: true })
const resultForLogs = describeRunResult(result)
```

The core flow is deliberately boring:

1. `defineModel` freezes a consumer-defined model.
2. `find` returns generic refs.
3. `checks` declares facts the runtime must prove.
4. An action builds workbook intent.
5. `prepareWorkbookAction` verifies the plan, computes requirements, emits
   JSON-safe `planData`, and gives the exact plan a stable id.
6. `runWorkbookPlan(..., { strict: true })` fails closed unless the adapter
   returns plan-bound apply proof, revision proof, resolved refs, command
   receipts, check proof, and no unverified apply facts.

## Which Package

| Package            | Choose when                                                                                   | Do not use for                                     |
| ------------------ | --------------------------------------------------------------------------------------------- | -------------------------------------------------- |
| `@bilig/workbook`  | Defining generic agent intent, refs, formulas, checks, plan data, schemas, and proof handoff. | Calculating formulas or owning workbook state.     |
| `@bilig/workpaper` | Running workbook tools, MCP, or product workflows around persisted WorkPaper state.           | Designing a reusable model API for other runtimes. |
| `@bilig/headless`  | Owning workbook state inside Node with formula recalculation and import/export.               | Publishing generic agent intent contracts.         |
| `@bilig/core`      | Implementing calculation or mutation internals.                                               | Consumer-facing agent model definitions.           |

The root export keeps the ordinary agent path: models, refs, checks, formulas,
plans, runtime proof, command results, schemas, and low-level ops. Subpaths are available when an agent wants a smaller import map:
`@bilig/workbook/model`,
`@bilig/workbook/prepare`, `@bilig/workbook/find`, `@bilig/workbook/check`, `@bilig/workbook/formula`,
`@bilig/workbook/verify`, `@bilig/workbook/runtime`,
`@bilig/workbook/command`, `@bilig/workbook/features`,
`@bilig/workbook/testing`, and `@bilig/workbook/schema`.

## Mental Model

Consumers define models. Bilig does not ship hardcoded business models in this
package.

Models are plain:

- `find(workbook)` binds the workbook parts the model needs.
- `checks({ refs, workbook })` declares proof the runtime must provide.
- `actions` publish constrained input metadata and write workbook intent.
- `prepareWorkbookAction(model, action)` is the canonical preflight for agents.

Refs are generic:

- `findName(name)` binds a named workbook ref.
- `findTable({ name, sheetName, headers })` binds a table by stable traits.
- `findColumn({ table, name })` and `table.column(name)` bind columns.
- `findRows({ table, where })` binds filtered rows.
- `findRange(input)` exists for explicit ranges when a consumer truly has one.

Formulas stay symbolic until a runtime materializes them:

- `formula.multiply(refs.input, refs.factor)` builds formula intent.
- `formula.raw(source, { inputs, labels })` accepts custom formula text.
- `formula.text(value)` creates a spreadsheet string literal.
- `@bilig/formula` parses and normalizes the formula language.
- `@bilig/core` or an app runtime calculates formulas.

Checks are part of the plan, not comments:

- `check.exists(ref)` proves the ref resolved.
- `check.noFormulaErrors(ref)` proves a formula target is clean.
- `check.valueEquals(ref, value)` proves a runtime value.
- `check.formulaEquals(ref, formula)` proves the runtime formula matches intent.
- `check.custom(options)` carries a runtime-owned proof contract.

## Agent-Safe Runtime

`@bilig/workbook` never mutates a workbook by itself. A runtime provides an
adapter:

```ts
const adapter = {
  apply(plan) {
    const ops = materializeForThisRuntime(plan)
    return {
      status: 'applied',
      planId: workbookPlanId(plan),
      baseRevision: currentRevision,
      revision: currentRevision + 1,
      previewOps: ops,
      appliedOps: ops,
      commandReceipts: receiptsFor(plan, ops),
      undo: { id: 'undo-1' },
    }
  },
  read(targets, plan) {
    return readTargetsFromRuntime(targets, plan)
  },
  verifyChecks(checks, plan) {
    return proveChecksFromRuntime(checks, plan)
  },
}
```

Use `runWorkbookPlan(planOrData, adapter, { strict: true })` when an agent needs
production proof. Strict mode requires:

- a valid plan before mutation
- at least one planned check before mutating actions
- adapter capabilities for the planned work
- plan id proof
- base and applied revision proof
- apply proof with no unverified apply facts
- concrete applied ops for every planned command
- command receipts bound to planned digests and concrete `resolvedRefs`
- proof on every passed check

Runtime authors can run the same plain-object, known-key, own-data-option
contract with the `@bilig/workbook/testing` adapter helpers.

The returned `WorkbookRunResult` is intentionally plain:

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

## Data Boundaries

Everything that crosses an agent/runtime boundary is inspectable data:

- `describeModel`, `describePlan`, `describePlanResult`, and `describeRunResult` return JSON-safe descriptions.
- `toPlanData`, `checkPlanData`, and `hydratePlanData` transport and restore executable plan data.
- `verifyPlan`, `verifyPlanData`, `verifyModel`, `checkInput`,
  `checkWorkbookModelDescription`, and `checkWorkbookReadbackProof` return frozen validation verdicts.
- `workbookJsonSchemas`, `workbookJsonSchemaHashes`, and `fixtures/` publish
  checked model, plan, runtime-requirements, command, run-result, and readback artifacts.
- Schemas cover transport shape and stay in parity for shape-enforceable
  constraints such as row predicates, destructive confirmation, and command
  receipt proof. Workbook-math limits such as `scope.maxTouchedCells` are
  enforced by `checkWorkbookCommandBundle`.

Public validators read own data properties and reject malformed, sparse,
accessor-backed, or custom-prototype payloads before hidden consumer code can
run. Public results are frozen before they cross the package boundary.

## Feature Commands

Runtimes can expose workbook extensions with the same data-first contract:

- `checkWorkbookCommandRequest`
- `checkWorkbookCommandBundle`
- `workbookCommandResultForReceipts`
- `checkWorkbookCommandResult`
- `checkWorkbookCommandResultForBundle`
- `checkWorkbookCommandReceipt`

Generic command request, bundle, result, and receipt validators are available on
the root path because agents may need to inspect runtime handoff proof. Runtime
plugin registration, projection interceptors, and UI contribution metadata live
only under `@bilig/workbook/features`. Ordinary models should prefer
`writeFormula`, `writeValue`, `format`, `clear`, and checks.
Format receipts use the same semantic proof path for single cells and ranges:
each requested style or number-format component must cover every resolved cell.

## Low-Level Ops

The existing workbook operation language remains public:

- `WorkbookOp`
- `WorkbookTxn`
- `EngineOp`
- `EngineOpBatch`
- guards such as `isEngineOpBatch`

Most models should not start there. Use `workbook.addOp(op, { target, message })`
inside an action only when the generic action helpers cannot express the
required workbook intent.

## Example

See [examples/workbook-agent-model](../../examples/workbook-agent-model) for a
generic model that plans, verifies, describes, transports, runs, and prints proof
without depending on a hardcoded business model.
