# Directus WorkPaper Flow Operation

This example is a Directus custom Flow operation that computes formula-backed
fields with `@bilig/workpaper`.

Use it when Directus owns the record workflow, but a pricing, quote, payout, or
validation calculation is still easiest to review as cells and formulas. The
operation returns a `patch` object for a following Directus **Update Data**
operation plus a readback proof object.

Directus Run Script operations cannot import arbitrary npm packages. Directus
documents custom operation extensions as the path for using npm packages in
Flows:

- https://directus.io/docs/guides/automate/operations
- https://directus.io/docs/tutorials/extensions/use-npm-packages-in-custom-operations

## Local Smoke

```sh
npm install
npm run typecheck
npm run build
npm run smoke
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

The live smoke output includes the full before/after summary fields.

## Flow Shape

1. Trigger from an item create/update event or manual Flow.
2. Run **Bilig WorkPaper Calculated Fields** with record values such as
   `quantity`, `unitPrice`, `discountRate`, `taxRate`, and `unitCost`.
3. Feed the returned `patch` into Directus **Update Data** for fields such as
   `subtotal`, `discount_amount`, `taxable_amount`, `tax_amount`, `total`, and
   `margin_amount`.
4. Store or log the returned `proof` when the workflow needs audit evidence.

The operation proves that the workbook changed one input, recalculated formulas,
serialized WorkPaper JSON, restored it, and read the same calculated values
after restore.

## Boundaries

This is a source example, not a Marketplace package. Use it as the starting
point for a project-owned Directus extension or a submission to a Directus
extension repository.

Keep Excel or another oracle in the loop for macros, pivots, external links,
locale-sensitive desktop behavior, or workbook features outside Bilig's formula
surface. The Directus host must run a Node version supported by
`@bilig/workpaper`.
