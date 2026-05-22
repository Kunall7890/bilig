# @bilig/workbook Agent Model Example

This is a runnable consumer model for agents. It shows the public flow:

1. Define a generic model with `defineModel`.
2. Inspect it with `describeModel`.
3. Plan an action with `planWorkbookAction`.
4. Verify the plan with `verifyPlan`.
5. Inspect runtime handoff needs with `describeRuntimeRequirements`.
6. Execute through a runtime adapter with `runWorkbookAction`.
7. Read a JSON-safe result with `describeRunResult`.

The model is intentionally consumer-owned. Bilig does not ship the table shape as
a built-in business model; the example only proves how a consumer can define one.

## Run

```sh
pnpm --dir examples/workbook-agent-model install --ignore-workspace
pnpm --dir examples/workbook-agent-model run start
```

The example depends on `@bilig/workbook` for the public model contract and
`@bilig/core` only as the chosen runtime adapter. `@bilig/workbook` itself does
not depend on `@bilig/core`.

