# Trigger.dev WorkPaper Task

This example is a Trigger.dev task that computes formula-backed quote fields
with `@bilig/workpaper`.

Use it when Trigger.dev owns durable task execution, retries, observability, and
workflow triggering, but the pricing, payout, or import-validation calculation
should stay reviewable as workbook cells and formulas.

Trigger.dev documents tasks as long-running resilient functions defined with
`task({ id, run })`; returned task output must be JSON serializable:

- https://trigger.dev/docs/tasks/overview
- https://trigger.dev/docs/triggering
- https://trigger.dev/docs/runs

## Local Smoke

```sh
pnpm install --ignore-workspace --lockfile=false
pnpm run typecheck
pnpm run smoke
```

The smoke test prints a verified result:

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

The live smoke output includes every calculated summary field. It calls the pure
WorkPaper helper so no Trigger.dev account or secret key is needed for local
verification.

## Trigger.dev Shape

1. Add `src/trigger-workpaper-task.ts` to the `trigger/` directory in your
   Trigger.dev project.
2. Keep `src/workpaper-quote.ts` next to it or move the helper into your app's
   normal service layer.
3. Deploy the task with the Trigger.dev CLI.
4. Trigger `bilig-workpaper-quote` with record values such as `quantity`,
   `unitPrice`, `discountRate`, `taxRate`, and `unitCost`.
5. Feed `patch` into your app's writeback step and keep `proof` in logs,
   metadata, or audit storage.

The task edits `Inputs!B2`, recalculates formulas, serializes WorkPaper JSON,
restores it, and verifies that restored calculated values match.

## Boundaries

This is a source example, not an official Trigger.dev template. Use it as the
starting point for a project task or a contribution to the Trigger.dev examples
repo after the proof is already public.

Keep Excel or another oracle in the loop for macros, pivots, external links,
locale-sensitive desktop behavior, or workbook features outside Bilig's formula
surface.
