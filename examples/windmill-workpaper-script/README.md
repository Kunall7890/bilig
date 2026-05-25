# Windmill WorkPaper Script

This example is a Windmill TypeScript script that computes formula-backed
quote fields with `@bilig/workpaper`.

Use it when Windmill owns schedules, webhooks, approvals, or workflow routing,
but the pricing, payout, or import-validation calculation should stay
reviewable as workbook cells and formulas.

Windmill's TypeScript runtime can resolve npm dependencies directly from script
imports and requires a `main` entrypoint for runnable scripts:

- https://www.windmill.dev/docs/advanced/dependencies_in_typescript
- https://www.windmill.dev/docs/getting_started/scripts_quickstart/typescript
- https://www.windmill.dev/docs/script_editor

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

The live smoke output includes every calculated summary field.

## Windmill Shape

1. Create a TypeScript script in Windmill.
2. Paste `src/workpaper-script.ts`.
3. Keep the import as `@bilig/workpaper`; Windmill will compute a script
   lockfile from the import on deployment.
4. Run the script with record values such as `quantity`, `unitPrice`,
   `discountRate`, `taxRate`, and `unitCost`.
5. Feed `patch` into the next Windmill step and keep `proof` in logs or audit
   storage.

The script edits `Inputs!B2`, recalculates formulas, serializes WorkPaper JSON,
restores it, and verifies that restored calculated values match.

## Boundaries

This is a copy-paste script example, not a Windmill Hub package. Use it as the
starting point for a workspace script, flow module, or workflow-as-code task.

Keep Excel or another oracle in the loop for macros, pivots, external links,
locale-sensitive desktop behavior, or workbook features outside Bilig's formula
surface.
