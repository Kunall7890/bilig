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
- `check`
- `describeModel`
- `describeRef`
- `describePlan`
- `describePlanResult`
- `verifyPlan`
- `verifyModel`
- `formula`
- `workbook.addOp(op, { target?, message? })` inside model actions
- `WorkbookModel`
- `WorkbookAction`
- `WorkbookActionInput`
- `WorkbookAddOpOptions`
- `WorkbookActionPlanResult`
- `WorkbookModelDescription`
- `WorkbookRefDescription`
- `WorkbookActionPlanDescription`
- `WorkbookActionPlanResultDescription`
- `WorkbookPlanVerification`
- `WorkbookPlanIssue`
- `WorkbookModelVerification`
- `WorkbookModelActionVerification`
- `WorkbookModelVerificationOptions`
- `WorkbookCustomCheckOptions`
- `WorkbookRawFormulaOptions`
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
Actions can also accept a JSON-safe input:

```ts
import { defineModel } from "@bilig/workbook";

export const model = defineModel({
  name: "custom-writer",

  find(workbook) {
    return {
      output: workbook.findRange({ sheetName: "Sheet1", address: "B2" }),
    };
  },

  actions: {
    write({ refs, workbook, input }) {
      if (typeof input !== "object" || input === null || Array.isArray(input)) {
        throw new Error("input object required");
      }
      const value = input.value;
      if (typeof value !== "number") {
        throw new Error("numeric value required");
      }
      workbook.writeValue(refs.output, value);
    },
  },
});
```

`planWorkbookAction(model, "write", { value: 12 })` clones and canonicalizes
that input into the plan so an agent can inspect exactly what was requested.
Inputs must be plain JSON values: strings, finite numbers, booleans, `null`,
arrays without holes, and plain objects. This package intentionally does not add
schema dependencies; consumers own their own input validation inside actions.
Use `verifyModel(model, { inputs: { write: { value: 12 } } })` when whole-model
verification needs parameters for specific actions.

Formula expressions also keep their workbook inputs separate from their formula
text. A planned `writeFormula` command includes both the parseable formula
string and the generic model refs it used, so an agent can inspect what the
action depends on without reverse-parsing placeholder names.
For formulas outside the small helper set, use
`formula.raw(source, { inputs })`; the source stays parseable while the
declared refs remain inspectable and verifiable. These inputs are a declared
dependency contract for agents, not parser-discovered proof that every formula
reference has been mapped to a model ref.

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
when an agent or test needs the same generic refs outside a model callback.
`findRows` refs include the predicate value in their stable id and label so two
consumer-defined row selectors do not collapse during dedupe.
Use `check.exists(ref)` and `check.noFormulaErrors(ref)` directly when an agent
or test needs the same generic planned checks outside a model callback.
Use `check.custom({ kind, message, target, refs })` for consumer-defined
invariants; the package does not need to know what the model means. `target`
names the main ref, and `refs` names any supporting refs the invariant depends
on so agents can describe and verify the full check contract.

Use `describeModel` when an agent needs a JSON-safe manifest of model name,
action names, and whether model-level checks exist without running `find`,
checks, or actions.
Use `describeRef` and `describePlan` when an agent needs JSON-safe intent for
logs, comparisons, approvals, or runtime handoff. Descriptions keep the same
generic action input, refs, commands, checks, changes, and ops, but omit
consumer-private `refs` object shape and helper functions such as
`table.column()`.
Use `describePlanResult` when the same JSON-safe handoff is needed for either
planned or failed action planning.

Use `verifyPlan` before runtime handoff when an agent needs to prove a planned
action is internally consistent. It checks for non-JSON-safe action input,
unresolved refs, unparsable formulas, duplicate resolved refs, and missing
concrete ops for write, clear, and number-format commands that already target a
known single cell. Custom check targets and supporting refs must also resolve
through the model's `refsUsed` contract. Low-level `addOp` commands must contain
valid `WorkbookOp` values, must still appear in `plan.ops`, and must match their
declared `target` when the op exposes a concrete address or range.
Use `verifyModel` to plan and verify every action in a consumer-defined model
with one JSON-safe result. Pass `inputs` when specific actions require
parameters.
