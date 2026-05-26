---
title: Inngest WorkPaper Step
published: true
description: Run Bilig WorkPaper inside an Inngest step.run boundary so durable workflows can calculate formula-backed fields and return verified readback.
tags: inngest, durable-workflows, step-functions, workbook-api, spreadsheet-automation
canonical_url: https://proompteng.github.io/bilig/inngest-workpaper-step.html
image: /assets/github-social-preview.png
---

# Inngest WorkPaper Step

Use this when an Inngest function needs formula-backed business logic, but the
calculation should be one durable `step.run()` boundary instead of spreadsheet
UI automation or ad hoc formula code.

Inngest owns event delivery, durable step execution, retries, run history, and
observability. Bilig owns formula workbook state, recalculation, JSON
serialization, restore, and readback proof.

Official Inngest references:

- <https://www.inngest.com/docs/reference/typescript/functions/create>
- <https://www.inngest.com/docs/reference/typescript/v3/functions/step-run>
- <https://www.inngest.com/docs/learn/how-functions-are-executed>
- <https://github.com/inngest/inngest>

## Example Artifact

The runnable source lives in:

```text
examples/inngest-workpaper-step
```

It contains:

- `src/inngest-workpaper-function.ts` for the `Inngest` function wrapper
- `src/workpaper-quote.ts` for the pure WorkPaper formula calculation
- `src/smoke.ts` for local no-account proof
- `scripts/check-inngest-recipe.ts` for static recipe checks

Run the proof locally:

```sh
cd examples/inngest-workpaper-step
pnpm install --ignore-workspace --lockfile=false
pnpm run check
pnpm run typecheck
pnpm run smoke
```

The smoke edits `Inputs!B2`, recalculates quote formulas, serializes the
WorkPaper document, restores it, verifies restored calculated values match, and
prints JSON.

## Function Shape

The checked-in function wrapper uses `step.run()` as the durable formula
boundary:

```ts
import { Inngest } from 'inngest'

import { calculateWorkPaperQuote } from './workpaper-quote.js'

export const inngest = new Inngest({ id: 'bilig-workpaper-example' })

export const calculateWorkPaperQuoteFunction = inngest.createFunction(
  {
    id: 'bilig-workpaper-quote',
    retries: 3,
    triggers: [{ event: 'bilig/quote.requested' }],
  },
  async ({ event, step }) => {
    const result = await step.run('calculate-workpaper-quote', async () => calculateWorkPaperQuote(event.data))

    if (!result.proof.verified) {
      throw new Error('WorkPaper proof failed')
    }

    return result
  },
)
```

The helper imports `@bilig/workpaper`, writes one input, recalculates formulas,
exports WorkPaper JSON, restores the document, and returns a compact
serializable patch plus proof.

## Step Output

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
  }
}
```

Use `patch` for app writeback. Keep `proof` in logs, run metadata, an audit
table, or object storage when the business workflow needs evidence.

## Production Gate

Before using this pattern for customer-critical workflows:

1. Keep `@bilig/workpaper`, XLSX import/export, file I/O, and object-store
   writes inside `step.run()` handlers or service helpers they call.
2. Keep step return values compact and JSON serializable; put large WorkPaper
   JSON documents in a database or object store.
3. Make external writes idempotent because Inngest can retry failed steps.
4. Pin or lock runtime package versions in deployed workers.
5. Keep Excel, LibreOffice, Microsoft Graph, or a domain oracle in the loop for
   macros, pivots, external links, and exact spreadsheet behavior.

## When This Fits

Use it for quote approvals, pricing rules, payout checks, import validation,
order-review gates, and durable AI workflow steps where spreadsheet formulas are
the most reviewable representation of business logic.

Do not use it to pretend Bilig is desktop Excel. Use Inngest for durable
execution, `step.run()` for the retriable formula boundary, and external oracles
for Excel-specific behavior.

## Outreach Boundary

If this is shared in an Inngest issue, discussion, or docs proposal, lead with
the workflow boundary:

> Inngest owns durable steps and retries. Bilig owns formula workbook state and
> returns a compact patch plus readback proof from one `step.run()`.

Do not post it as a generic spreadsheet-engine pitch. Link the runnable example
and smoke output, then ask whether a formula-backed durable step recipe would be
useful to Inngest users.
