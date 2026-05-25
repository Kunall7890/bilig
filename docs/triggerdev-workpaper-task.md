---
title: Trigger.dev WorkPaper task
published: true
description: Run Bilig WorkPaper inside a Trigger.dev TypeScript task to calculate durable workflow fields with formula readback, JSON persistence, and restore proof.
tags: triggerdev, background-jobs, workflow-automation, workbook-api, spreadsheet-automation
canonical_url: https://proompteng.github.io/bilig/triggerdev-workpaper-task.html
image: /assets/github-social-preview.png
---

# Trigger.dev WorkPaper task

Use this when a Trigger.dev task needs spreadsheet-shaped business logic, but
the formula state should be edited and verified through a TypeScript API instead
of Excel UI automation, browser grid clicks, or stale cached XLSX formula
values.

Trigger.dev's current task API defines resilient TypeScript functions with
`task({ id, run })`. The docs describe tasks as long-running functions with
retry settings, dashboard-visible run state, and JSON-serializable return
values.

Official Trigger.dev references:

- <https://trigger.dev/docs/tasks/overview>
- <https://trigger.dev/docs/triggering>
- <https://trigger.dev/docs/runs>
- <https://trigger.dev/docs/llms.txt>

## Example Artifact

The runnable source lives in:

```text
examples/triggerdev-workpaper-task
```

It contains:

- `src/workpaper-quote.ts` for the account-free WorkPaper calculation helper
- `src/trigger-workpaper-task.ts` for the Trigger.dev `task` wrapper
- `src/smoke.ts` for a no-Trigger local proof

Run it locally:

```sh
cd examples/triggerdev-workpaper-task
pnpm install --ignore-workspace --lockfile=false
pnpm run typecheck
pnpm run smoke
```

The smoke test edits `Inputs!B2`, recalculates quote formulas, serializes the
WorkPaper document, restores it, and verifies that the restored calculated
values match.

## Trigger.dev Task

Put this file in your Trigger.dev `trigger/` directory:

```ts
import { task } from "@trigger.dev/sdk";

import { calculateWorkPaperQuote } from "./workpaper-quote";

export const calculateWorkPaperQuoteTask = task({
  id: "bilig-workpaper-quote",
  retry: {
    maxAttempts: 3,
    minTimeoutInMs: 500,
    maxTimeoutInMs: 30_000,
    factor: 1.8,
    randomize: true,
  },
  run: async (payload) => calculateWorkPaperQuote(payload),
});
```

The real example keeps the payload typed and imports from
`src/workpaper-quote.ts`.

## Task Input

```json
{
  "previousQuantity": 12,
  "quantity": 18,
  "unitPrice": 125,
  "discountRate": 0.1,
  "taxRate": 0.08,
  "unitCost": 52
}
```

## Task Output

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
    "persistedDocumentBytes": 1208,
    "verified": true
  }
}
```

The live smoke output prints all calculated summary fields.

## Workflow Shape

1. Trigger `bilig-workpaper-quote` from your app, webhook handler, schedule, or
   another task.
2. Pass record values such as `quantity`, `unitPrice`, `discountRate`,
   `taxRate`, and `unitCost`.
3. Feed `patch` into the next task or application writeback step.
4. Store `proof` in run metadata, logs, or an audit table when the workflow
   needs readback evidence.

That keeps Trigger.dev in charge of durable execution, retries, observability,
and run history. Bilig owns the formula workbook, recalculation, JSON
serialization, restore, and readback proof.

## When This Fits

Use it for quote approvals, pricing rules, discount calculations, payout checks,
import validation, and durable AI workflow steps that should be reviewable as
formulas but executed inside a Trigger.dev task.

Do not use it to pretend Bilig is desktop Excel. Keep Excel, LibreOffice,
Microsoft Graph, or a domain oracle in the loop for macros, pivots, charts,
external links, and exact spreadsheet UI behavior.

## Outreach Note

If this is shared in a Trigger.dev issue, discussion, or examples PR, lead with
the concrete boundary it solves:

> Trigger.dev owns durable execution. Bilig owns the formula workbook and
> returns both calculated field values and readback proof.

Do not post it as a generic spreadsheet-engine pitch. Link the runnable example
and smoke output, then ask whether a formula-backed durable task would be useful
as a Trigger.dev examples repo contribution.
