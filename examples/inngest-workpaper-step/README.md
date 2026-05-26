# Inngest WorkPaper Step

Use this when an Inngest function needs durable formula-backed fields, but the
spreadsheet calculation should be one retriable `step.run()` boundary instead of
ad hoc spreadsheet code in every workflow.

Inngest owns event delivery, step memoization, retries, run history, and
observability. Bilig owns workbook formula state, recalculation, JSON
persistence, and readback proof.

Official Inngest references:

- <https://www.inngest.com/docs/reference/typescript/functions/create>
- <https://www.inngest.com/docs/reference/typescript/v3/functions/step-run>
- <https://www.inngest.com/docs/learn/how-functions-are-executed>

## Local Smoke

```sh
cd examples/inngest-workpaper-step
pnpm install --ignore-workspace --lockfile=false
pnpm run check
pnpm run typecheck
pnpm run smoke
```

The smoke test calls the same pure WorkPaper helper used by the Inngest function
wrapper, so no Inngest account, event key, dev server, or cloud app is needed to
verify the workbook boundary.

Expected output:

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

## Inngest Shape

`src/inngest-workpaper-function.ts` defines the real function wrapper:

```ts
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

Use the returned `patch` for app writeback and keep `proof` in run logs,
metadata, an audit table, or an object store record.

## Boundary

Use this pattern when:

- the workflow starts from an event such as `bilig/quote.requested`;
- formula-backed fields should be retried independently from notification,
  persistence, or fulfillment steps;
- downstream code needs a compact JSON patch;
- reviewers need before/after/restore proof for the formula result.

Do not use it to pretend Bilig is desktop Excel. Keep Excel, LibreOffice,
Microsoft Graph, or a domain oracle in the loop for macros, pivots, external
links, and exact spreadsheet behavior.

## Outreach Note

If this is shared with Inngest users, lead with the step boundary:

> Inngest owns durable steps and retries. Bilig owns formula workbook state and
> returns a compact patch plus readback proof from one `step.run()`.

Do not post it as a generic spreadsheet-engine pitch. Link the runnable example
and smoke output, then ask whether a formula-backed durable step recipe would be
useful to Inngest users.
