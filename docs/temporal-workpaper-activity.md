---
title: Temporal WorkPaper Activity
published: true
description: Run Bilig WorkPaper from a Temporal TypeScript Activity so Workflows stay deterministic while formula readback and JSON restore proof stay in normal Node.js code.
tags: temporal, durable-execution, workflow-automation, workbook-api, spreadsheet-automation
canonical_url: https://proompteng.github.io/bilig/temporal-workpaper-activity.html
image: /assets/github-social-preview.png
---

# Temporal WorkPaper Activity

Use this when a Temporal TypeScript Workflow needs formula-backed decisions, but
the calculation should run through an Activity instead of Workflow replay code.

Temporal's TypeScript guide separates deterministic Workflow code from
Activities. Activities are the right boundary for WorkPaper state, file I/O,
Excel/XLSX import, object storage, and formula recalculation. This example keeps
that boundary: Temporal owns durable orchestration, retries, workflow history,
and replay; Bilig owns formula workbook state, recalculation, JSON
serialization, restore, and readback proof.

Official Temporal references:

- <https://docs.temporal.io/develop/typescript/>
- <https://docs.temporal.io/develop/typescript/best-practices/testing-suite>
- <https://github.com/temporalio/samples-typescript#external-apps--libraries>
- <https://github.com/temporalio/samples-typescript#contributing>

## Example Artifact

The runnable source lives in:

```text
examples/temporal-workpaper-activity
```

It contains:

- `src/workflows.ts` for the deterministic Workflow boundary
- `src/activities.ts` for the WorkPaper formula Activity
- `src/smoke.ts` for a local Activity smoke through `MockActivityEnvironment`
- `scripts/check-temporal-boundary.ts` for import-boundary checks

Run the proof locally:

```sh
cd examples/temporal-workpaper-activity
pnpm install --ignore-workspace --lockfile=false
pnpm run check
pnpm run smoke
```

The smoke edits `Inputs!B2`, recalculates quote formulas, serializes the
WorkPaper document, restores it, verifies restored calculated values match,
prints JSON, and writes `.tmp/workpaper-proof.json`.

## Workflow Shape

The checked-in Workflow never imports Bilig:

```ts
import { proxyActivities } from '@temporalio/workflow'
import type { TemporalWorkPaperActivities } from './types'

const { calculateWorkPaperQuoteActivity } = proxyActivities<TemporalWorkPaperActivities>({
  startToCloseTimeout: '30 seconds',
  retry: {
    maximumAttempts: 3,
  },
})

export async function quoteApprovalWorkflow(input) {
  return await calculateWorkPaperQuoteActivity(input)
}
```

The Activity imports `@bilig/workpaper`, writes one input, recalculates formulas,
exports WorkPaper JSON, restores the document, and returns a compact
serializable proof.

## Activity Output

The proof contains:

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

Keep the full WorkPaper document in an artifact path, database, or object store
when it grows beyond a compact Activity result.

## Production Gate

Before using this pattern for customer-critical Workflows:

1. Run a real Worker against a local Temporal dev server or Temporal Cloud.
2. Keep `@bilig/workpaper`, XLSX import/export, file I/O, network calls, and
   object-store writes in Activities.
3. Use Activity retry/idempotency design for external writes and proof-artifact
   paths.
4. Use `WorkflowReplayer` against captured histories when Workflow code changes.
5. Keep Excel, LibreOffice, Microsoft Graph, or a domain oracle in the loop for
   macros, pivots, external links, and exact spreadsheet behavior.

## When This Fits

Use it for quote approvals, pricing rules, payout checks, import validation,
order-review gates, and durable workflows where spreadsheet formulas are the
most reviewable representation of business logic.

Do not use it to pretend Bilig is desktop Excel. Use Temporal for orchestration
and replay, Activities for WorkPaper formula work, and external oracles for
Excel-specific behavior.

## Outreach Note

If this is shared in a Temporal issue, forum thread, or sample request, lead with
the concrete boundary it solves:

> Temporal owns durable orchestration and replay. Bilig owns formula workbook
> recalculation inside an Activity and returns both calculated values and
> readback proof.

Do not post it as a generic spreadsheet-engine pitch. Link the runnable example
and smoke output, then ask whether a formula-backed Activity boundary would be
useful to TypeScript SDK users.
