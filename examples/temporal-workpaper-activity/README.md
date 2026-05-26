# Temporal WorkPaper Activity

Use this when a Temporal TypeScript Workflow needs formula-backed decisions, but
the workbook engine must not run inside Workflow replay code.

The example keeps Temporal Workflows deterministic: `src/workflows.ts` imports
only `@temporalio/workflow` plus local types and calls a proxied Activity. The
Activity runs as normal Node.js code, owns `@bilig/workpaper`, writes an input
cell, recalculates formulas, persists WorkPaper JSON, restores it, and returns a
compact serializable patch plus proof object.

## Run Locally

```sh
cd examples/temporal-workpaper-activity
pnpm install --ignore-workspace --lockfile=false
pnpm run check
pnpm run smoke
```

The smoke uses Temporal's `MockActivityEnvironment`, edits `Inputs!B2`,
recalculates quote formulas, persists a WorkPaper JSON document, restores it,
and verifies restored readback.

## Files

- `src/workflows.ts` defines `quoteApprovalWorkflow` and calls
  `calculateWorkPaperQuoteActivity` through `proxyActivities`.
- `src/activities.ts` owns `@bilig/workpaper`, formula recalculation, JSON
  restore, and proof output.
- `src/smoke.ts` runs the Activity through `MockActivityEnvironment`.
- `scripts/check-temporal-boundary.ts` verifies the Workflow never imports
  `@bilig/workpaper` or file I/O.

## Workflow Boundary

Workflow code imports only `@temporalio/workflow`:

```ts
import { proxyActivities } from '@temporalio/workflow'
import type { TemporalWorkPaperActivities } from './types'

const { calculateWorkPaperQuoteActivity } = proxyActivities<TemporalWorkPaperActivities>({
  startToCloseTimeout: '30 seconds',
  retry: {
    maximumAttempts: 3,
  },
})
```

The Activity imports `@bilig/workpaper` and returns a serializable result:

```ts
const result = await calculateWorkPaperQuoteActivity({
  previousQuantity: 12,
  quantity: 18,
  unitPrice: 125,
  discountRate: 0.1,
  taxRate: 0.08,
  unitCost: 52,
  output: '.tmp/workpaper-proof.json',
})
```

## Proof Shape

The smoke prints and writes:

```json
{
  "patch": {
    "subtotal": 2250,
    "discount_amount": 225,
    "taxable_amount": 2025,
    "tax_amount": 162,
    "total": 2187,
    "margin_amount": 1089
  },
  "proof": {
    "editedCell": "Inputs!B2",
    "before": {
      "total": 1458
    },
    "after": {
      "total": 2187
    },
    "afterRestore": {
      "total": 2187
    },
    "verified": true
  },
  "temporalBoundary": {
    "workflowImportsWorkPaper": false,
    "activityOwnsWorkPaper": true,
    "payloadShape": "serializable-patch-and-proof"
  }
}
```

## Production Gate

Before using this pattern for customer-critical Workflows:

1. Run a real Worker against a local Temporal dev server or Temporal Cloud.
2. Keep `@bilig/workpaper`, Excel import/export, file I/O, network calls, and
   object-store writes in Activities.
3. Keep Activity arguments and results compact; put full WorkPaper JSON in an
   artifact path, database, or object store when it grows.
4. Use Temporal `WorkflowReplayer` against captured Workflow histories whenever
   Workflow code changes.
5. Keep Excel, LibreOffice, Microsoft Graph, or a domain oracle in the loop for
   macros, pivots, external links, and exact desktop Excel behavior.

## Outreach Note

If this is shared with Temporal users, lead with the replay boundary:

> Temporal owns durable orchestration and replay. Bilig owns formula workbook
> recalculation inside an Activity and returns a compact proof.

Do not post it as a generic spreadsheet-engine pitch. Link the runnable example
and smoke output, then ask whether a formula-backed Activity boundary would be
useful to TypeScript SDK users.
