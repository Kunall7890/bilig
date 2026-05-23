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
  verifyPlan,
} from "@bilig/workbook";

export const model = defineModel({
  name: "generic-row-calculator",

  find(workbook) {
    const table = workbook.findTable({
      headers: ["Item", "Quantity", "Rate", "Status", "Total"],
    });
    const rows = workbook.findRows({
      table,
      where: { column: "Status", op: "eq", value: "ready" },
    });

    return {
      table,
      rows,
      quantity: rows.column("Quantity"),
      rate: rows.column("Rate"),
      total: rows.column("Total"),
    };
  },

  checks({ refs, workbook }) {
    return [workbook.check.exists(refs.table), workbook.check.noFormulaErrors(refs.total)];
  },

  actions: {
    recompute({ refs, workbook }) {
      const expected = formula.multiply(refs.quantity, refs.rate);
      workbook.writeFormula(refs.total, expected);
      workbook.check.formulaEquals(refs.total, expected);
    },
  },
});

const planned = planWorkbookAction(model, "recompute");
if (planned.status === "failed") throw new Error(planned.errors[0]?.message);

const staticProof = verifyPlan(planned.plan);
const requirements = describeRuntimeRequirements(planned.plan);
const planForLogs = describePlan(planned.plan);

const result = await runWorkbookPlan(planned.plan, adapter);
const resultForLogs = describeRunResult(result);
```

That is the core flow:

1. `defineModel` freezes a consumer-defined model.
2. `find` returns generic refs.
3. `checks` declares facts the runtime must prove.
4. An action builds workbook intent.
5. `verifyPlan` checks the plan without running an engine.
6. `describeRuntimeRequirements` tells an adapter what it must apply, read, and prove.
7. `runWorkbookPlan` applies the plan through a runtime-owned adapter and returns a boring result with check proof and apply proof when the adapter provides it.

## Public Contract

The main API is intentionally small:

- model: `defineModel`, `inspectModel`, `planWorkbookAction`, `buildWorkbookActionPlan`
- selectors: `findTable`, `findColumn`, `findRange`, `findName`, `findRows`, `find`
- checks: `check.exists`, `check.noFormulaErrors`, `check.valueEquals`, `check.formulaEquals`, `check.custom`
- formulas: `formula.add`, `formula.subtract`, `formula.multiply`, `formula.divide`, `formula.sum`, `formula.call`, `formula.raw`, `formula.text`
- proof: `verifyPlan`, `verifyModel`, `verifyWorkbookReadbacks`
- descriptions: `describeModel`, `describeRef`, `describePlan`, `describePlanResult`, `describeRuntimeRequirements`, `describeRunResult`
- runtime handoff: `runWorkbookPlan`, `runWorkbookAction`, `WorkbookRunAdapter`
- low-level language: `WorkbookOp`, `WorkbookTxn`, `EngineOp`, `EngineOpBatch`, `isEngineOpBatch`

Stable data helpers are exported for generic tool builders:

- `workbookRefKinds`, `isWorkbookRefKind`, `isWorkbookRef`
- `workbookRowOperators`, `workbookRowOperatorValueTypes`, `isWorkbookRowOperator`, `isWorkbookRowValueCompatible`
- `builtInWorkbookCheckKinds`, `isBuiltInWorkbookCheckKind`
- `workbookActionInputDescriptionKinds`, `isWorkbookActionInputDescriptionKind`, `isWorkbookActionInputDescription`, `isWorkbookActionInput`
- `workbookRunErrorCodes`, `isWorkbookRunErrorCode`

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

## Formulas

`@bilig/workbook` creates formula expressions. `@bilig/formula` parses and
normalizes formula text. `@bilig/core` or an app runtime calculates it.

Formula helpers keep formula text and workbook dependencies separate. A planned
formula write includes both the formula string and the refs used to build it.
For custom formula text, use `formula.raw(source, { inputs })`. For spreadsheet
string literals, use `formula.text(value)`. Bare strings are not formula
operands because agents should not guess whether a string is code, a label, a
named range, or user text.

## Runtime Adapter

`@bilig/workbook` does not execute plans. A runtime owns that:

```ts
const adapter = {
  apply(plan) {
    const ops = materializeForThisRuntime(plan);
    return {
      status: "applied",
      previewOps: ops,
      appliedOps: ops,
      proof: { source: "runtime", opCount: ops.length },
      undo: { id: "undo-1" },
    };
  },
  read(targets, plan) {
    return targets.map((target) => ({ target, value: 12 }));
  },
  verifyChecks(checks, plan) {
    return checks.map((entry) => ({ ...entry, status: "passed" }));
  },
};
```

`runWorkbookPlan` refuses to call `apply` if static plan verification fails.
If an adapter returns both `previewOps` and `appliedOps`, the result reports
whether they matched. If the adapter returns neither, the run records an
unverified apply fact. Use `runWorkbookPlan(plan, adapter, { requireApplyProof:
true })` when an agent must fail closed instead of accepting an unproved apply.
Readback checks attach proof to passed checks, such as
`{ source: "readback", value: 12 }` or
`{ source: "readback", formula: "(Table[Quantity])*(Table[Rate])" }`.
Generic check verifiers may only change `status` or add JSON-safe `proof`; they
cannot rewrite the check contract.

The result is deliberately plain:

```ts
type WorkbookRunResult =
  | {
      status: "done";
      apply?: WorkbookRunApplySummary;
      changed: WorkbookChangeSummary[];
      checks: WorkbookCheckResult[];
      undo?: WorkbookUndoRef;
      unverified?: WorkbookRunUnverified[];
    }
  | {
      status: "failed";
      errors: WorkbookRunError[];
      apply?: WorkbookRunApplySummary;
      checks: WorkbookCheckResult[];
      unverified?: WorkbookRunUnverified[];
    };
```

## Low-Level Ops

Most models should use the small action API: `writeFormula`, `writeValue`,
`format`, `clear`, and checks. If a consumer needs the existing workbook
operation language directly, call `workbook.addOp(op, { target, message })`
inside an action.

The op is guarded with `isWorkbookOp`, cloned into `plan.ops`, and kept in the
command log. If a target is supplied and the op exposes a concrete range,
`verifyPlan` checks that the target and op agree.

## Example

See [examples/workbook-agent-model](../../examples/workbook-agent-model) for a
generic model that plans, verifies, describes, runs, and prints proof without
depending on a hardcoded business model.
