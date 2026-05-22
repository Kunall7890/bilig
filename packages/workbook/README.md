# @bilig/workbook

Agent-first workbook models and portable workbook intent for Bilig.

Build `@bilig/workbook` so an agent would love using it: simple, generic,
predictable, inspectable, verifiable, and never dependent on hardcoded business
models or human spreadsheet UI assumptions.

This package is for consumers who want to define their own workbook models. It
does not ship revenue models, prepaid models, reporting templates, or any other
business-specific assumptions. It gives you a small generic language for finding
workbook structure, planning changes, declaring checks, and handing the plan to
a runtime that can execute it.

## Mental Model

`@bilig/workbook` does three things:

1. Define a model with consumer-owned refs, checks, and actions.
2. Turn an action into a frozen, JSON-safe, inspectable plan.
3. Verify the plan and compare runtime readbacks against declared checks.

It does not calculate formulas or own workbook state. That stays in
`@bilig/core`, `apps/bilig`, or a consumer-provided runtime adapter.

```text
consumer model -> @bilig/workbook plan -> runtime adapter -> proof result
```

Refs are the public handle an agent should reason about. Ref ids are stable and
collision-resistant, but they are opaque. Use `label` or `describeRef(ref)` for
logs and explanations. Do not parse ids.

## Install

```sh
pnpm add @bilig/workbook
```

Runtime dependencies are intentionally small:

- `@bilig/protocol`
- `@bilig/formula`

This package does not depend on `@bilig/core`, `@bilig/headless`,
`@bilig/agent-api`, `zod`, or `effect`.

## Quick Start

```ts
import { defineModel, formula } from "@bilig/workbook";

export const model = defineModel({
  name: "custom-calculation",

  find(workbook) {
    const table = workbook.findTable({
      headers: ["Base", "Rate", "Result"],
    });

    return {
      table,
      base: table.column("Base"),
      rate: table.column("Rate"),
      result: table.column("Result"),
    };
  },

  checks({ refs, workbook }) {
    return [
      workbook.check.exists(refs.table),
      workbook.check.noFormulaErrors(refs.result),
    ];
  },

  actions: {
    calculate({ refs, workbook }) {
      workbook.writeFormula(refs.result, formula.multiply(refs.base, refs.rate));
      workbook.check.noFormulaErrors(refs.result);
    },
  },
});
```

The model is generic. The consumer decides what the table means. Bilig only sees
workbook refs, formula intent, checks, and portable operations.

## Planning

Use `planWorkbookAction` when the action name or input may come from an agent,
tool call, or user request.

```ts
import {
  describePlanResult,
  planWorkbookAction,
  verifyPlan,
} from "@bilig/workbook";
import { model } from "./model.js";

const planned = planWorkbookAction(model, "calculate");

console.log(describePlanResult(planned));

if (planned.status === "planned") {
  console.log(verifyPlan(planned.plan));
}
```

A planned action contains:

- `refs`: the consumer-defined refs shape returned by `find`.
- `refsUsed`: a flat deduped list of refs found inside `refs`.
- `commands`: high-level workbook intent such as `writeFormula`, `writeValue`,
  `format`, `clear`, and `op`.
- `ops`: concrete portable workbook ops when they are known without runtime
  resolution.
- `checks`: planned proof requirements.
- `changed`: human-readable summaries for logs and approvals.

Plans are frozen. They are meant to be inspected, passed to a runtime, and
verified, not mutated after creation.

## Finding Workbook Structure

Use the find helpers to create refs:

- `findTable({ name?, sheetName?, headers? })`
- `findColumn({ table, rows?, name })`
- `findRange({ sheetName, address })`
- `findRange({ sheetName, startAddress, endAddress })`
- `findName(name)`
- `findRows({ table, where })`

`findRows` is table-backed. Create a table ref first, then create row-scoped
refs from it:

```ts
const table = workbook.findTable({ name: "Inputs" });
const rows = workbook.findRows({
  table,
  where: { column: "Status", op: "eq", value: "Active" },
});

return {
  rows,
  amount: rows.column("Amount"),
};
```

Selector helpers trim names, normalize cell addresses, reject invalid ranges,
reject empty headers, and reject non-finite row predicate values before runtime
handoff.

## Formulas

Use `formula.*` helpers when possible:

```ts
formula.multiply(refs.base, refs.rate);
formula.sum(refs.amount, refs.tax);
formula.raw("SUM(Inputs[Amount])", { inputs: [refs.amount] });
```

Formula expressions keep formula text and workbook inputs separate. That gives
agents a direct dependency list instead of forcing them to reverse-engineer
formula text.

The rule is:

```text
@bilig/workbook creates formula expressions
@bilig/formula parses and normalizes formula language
@bilig/core calculates formulas
```

Runtime adapters materialize declared formula inputs by replacing whole formula
tokens only. A token inside a string literal or inside a larger identifier is
left alone.

## Checks

Checks are planned proof, not decorations.

```ts
workbook.check.exists(refs.table);
workbook.check.noFormulaErrors(refs.result);
workbook.check.valueEquals(refs.total, 42);
workbook.check.valuesEqual(refs.result, [[6], [20]]);
workbook.check.formulaEquals(refs.output, formula.add(refs.left, refs.right));
workbook.check.custom({
  kind: "consumerInvariant",
  target: refs.table,
  refs: [refs.amount, refs.result],
  message: "Consumer-defined invariant passed",
});
```

Runtime proof must move checks from `planned` to `passed` or `failed`.
`status: "done"` means every check has proof.

## Running Plans

`runWorkbookPlan` and `runWorkbookAction` provide the generic apply-and-prove
loop. The adapter owns execution.

```ts
import { runWorkbookAction } from "@bilig/workbook";

const result = await runWorkbookAction(model, "calculate", {
  preview(plan) {
    return runtime.preview(plan);
  },
  apply(plan) {
    return runtime.apply(plan);
  },
  read(targets, plan) {
    return runtime.read(targets, plan);
  },
  verifyChecks(checks, plan) {
    return runtime.verifyChecks(checks, plan);
  },
});
```

The public result shape is deliberately boring:

```ts
type WorkbookRunResult =
  | {
      status: "done";
      changed: WorkbookChangeSummary[];
      checks: WorkbookCheckResult[];
      undo?: WorkbookUndoRef;
    }
  | {
      status: "failed";
      errors: WorkbookRunError[];
      checks: WorkbookCheckResult[];
      undo?: WorkbookUndoRef;
    };
```

If apply succeeds and proof later fails, the failed result preserves the undo
ref when the adapter supplied one.

## Runtime Handoff

`@bilig/core` provides the canonical engine adapter:

```ts
import { createWorkbookRunAdapter } from "@bilig/core";
import { runWorkbookAction } from "@bilig/workbook";

const result = await runWorkbookAction(
  model,
  "calculate",
  createWorkbookRunAdapter(engine),
);
```

The core adapter:

- materializes generic commands into concrete engine ops;
- aligns table-column formula inputs row by row;
- applies additional low-level ops that are not already represented by
  materialized commands;
- reads semantic values and formulas for readback checks;
- verifies generic `exists` and `noFormulaErrors` checks;
- returns portable undo ops when the engine captures undo.

## Low-Level Ops

The existing transport-neutral operation language remains public:

- `WorkbookOp`
- `WorkbookTxn`
- `EngineOp`
- `EngineOpBatch`
- `isEngineOp`
- `isEngineOps`
- `isEngineOpBatch`

Use `workbook.addOp(op, { target?, message? })` when a consumer needs the
low-level language directly inside a model action.

```ts
workbook.addOp(
  {
    kind: "setCellValue",
    sheetName: "Sheet1",
    address: "B2",
    value: 42,
  },
  {
    target: refs.output,
    message: "Seed output value",
  },
);
```

The op guards validate finite literal values, parseable formula text, valid cell
addresses, ordered ranges, non-empty identifiers, and known enum values.

## Agent Workflow

For an agent, the safe flow is:

1. Inspect the model with `describeModel(model)`.
2. Plan the action with `planWorkbookAction(model, actionName, input)`.
3. Show or log `describePlanResult(planned)`.
4. Reject or repair the request if `verifyPlan(plan)` is invalid.
5. Run the plan through a runtime adapter.
6. Inspect `describeRunResult(result)`.
7. If the result failed after apply and includes `undo`, offer or run rollback
   through the runtime that owns execution.

Do not infer business meaning from Bilig. Put business meaning in the consumer
model, checks, action names, descriptions, and input metadata.

## Public API Map

Primary authoring APIs:

- `defineModel`
- `formula`
- `find`, `findTable`, `findColumn`, `findRange`, `findName`, `findRows`
- `check`
- `workbook.writeFormula`, `workbook.writeValue`, `workbook.format`,
  `workbook.clear`, `workbook.addOp` inside actions

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

- `verifyPlan`
- `verifyModel`
- `runWorkbookPlan`
- `runWorkbookAction`
- `verifyWorkbookReadbacks`
- `describeRunResult`

Stable code guards:

- `workbookPlanIssueCodes`
- `isWorkbookPlanIssueCode`
- `workbookReadbackIssueCodes`
- `isWorkbookReadbackIssueCode`
- `workbookRunErrorCodes`
- `isWorkbookRunErrorCode`

Types are exported for the same surfaces, including `WorkbookModel`,
`WorkbookAction`, `WorkbookActionPlan`, `WorkbookRunResult`,
`WorkbookCheckResult`, `WorkbookRunAdapter`, `WorkbookRunReadback`,
`WorkbookPlanIssue`, `WorkbookReadbackIssue`, and `WorkbookRunError`.

## What Not To Do

- Do not hardcode business models in this package.
- Do not parse ref ids for meaning.
- Do not treat a planned check as proof.
- Do not execute plans in this package.
- Do not add runtime dependencies on `@bilig/core`, `@bilig/headless`,
  `@bilig/agent-api`, `zod`, or `effect`.

## Development

```sh
pnpm --filter @bilig/workbook test
pnpm --filter @bilig/workbook build
pnpm typecheck
pnpm lint
```
