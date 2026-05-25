# Directus WorkPaper Flow Operation

Use this when a Directus Flow needs persisted calculated fields, but the
calculation is easier to review as workbook formulas than as a pile of
JavaScript assignments.

Directus Flow operations are the actions inside a Flow.
Run Script executes in an isolated sandbox without access to npm modules.
For a third-party package, Directus documents a custom operation extension with
an API entrypoint and an app entrypoint:

- <https://directus.io/docs/guides/automate/operations>
- <https://directus.io/docs/tutorials/extensions/use-npm-packages-in-custom-operations>

Bilig fits that boundary: Directus owns the event and persistence workflow,
while `@bilig/workpaper` owns the formula workbook, recalculation, JSON
serialization, restore, and readback proof.

## Example Artifact

The runnable source lives in:

```text
examples/directus-workpaper-flow-operation
```

It contains:

- `src/app.ts` for the Directus operation UI metadata
- `src/api.ts` for the Directus operation API handler
- `src/workpaper-calculated-fields.ts` for the WorkPaper formula calculation
- `src/smoke.ts` for a no-Directus local proof

Run it locally:

```sh
cd examples/directus-workpaper-flow-operation
npm install
npm run typecheck
npm run build
npm run smoke
```

The smoke test edits `Inputs!B2`, recalculates quote formulas, serializes the
WorkPaper document, restores it, and verifies that the restored calculated
values match.

## Operation Input

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

## Operation Output

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

## Directus Flow Shape

1. Trigger on item create/update or run the Flow manually.
2. Run the **Bilig WorkPaper Calculated Fields** operation with the relevant
   record values.
3. Feed `patch` into a Directus **Update Data** operation to persist fields such
   as `subtotal`, `discount_amount`, `taxable_amount`, `tax_amount`, `total`,
   and `margin_amount`.
4. Keep `proof` in logs or an audit field when the business workflow needs a
   readback trail.

That keeps Directus in charge of records, permissions, and Flow orchestration.
The formula model stays in code where it can be tested, versioned, and reviewed.

## When This Fits

Use it for quote approvals, pricing rules, discount calculations, payout checks,
import validation, and operational fields that must be stored back on a Directus
record after formula readback passes.

Do not use it to pretend Bilig is desktop Excel. Keep Excel, LibreOffice,
Microsoft Graph, or a domain oracle in the loop for macros, pivots, charts,
external links, and exact spreadsheet UI behavior. The Directus host must run a
Node version supported by `@bilig/workpaper`.

## Outreach Note

If this is shared in a Directus issue or discussion, lead with the constraint it
solves:

> Run Script cannot import npm packages, so this uses the documented custom
> operation boundary. The operation returns both a persisted field patch and the
> workbook readback proof.

Do not post it as a generic spreadsheet-engine pitch. Link the runnable example
and the smoke output, then ask whether a formula-backed persisted-field
operation matches the Directus extension need.
