# @bilig/workbook Agent Model Example

This is a runnable consumer model for agents. It shows the public flow:

1. Define a generic model with `defineModel`.
2. Inspect it with `describeModel`.
3. Build a command bundle with `planWorkbookCommand`.
4. Verify the embedded plan with `verifyPlan`.
5. Inspect runtime handoff needs with `describeRuntimeRequirements`.
6. Inspect the command with `describeCommandBundle`.
7. Preview exact runtime materialization with `previewWorkbookCommandBundle`.
8. Execute the approved command with `runWorkbookCommandBundle`.
9. Read a JSON-safe result and receipt with `describeRunResult`.

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

The core adapter receipt proves the applied op count and kinds, synchronous
mutation propagation, and undo availability. App revisions and rendered readback
proof are intentionally left to the app runtime that owns persistence and UI
observation.
