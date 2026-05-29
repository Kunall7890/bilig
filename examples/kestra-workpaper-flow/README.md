# Kestra WorkPaper Flow

This example is a Kestra Node Commands flow that calculates quote fields with
`@bilig/workpaper`, writes a JSON proof file, and leaves orchestration to
Kestra.

`blueprint.yaml` is the self-contained Kestra Community Blueprint candidate.
It embeds the Node script, installs `@bilig/workpaper`, writes one input cell,
recalculates formula outputs, restores serialized JSON, and persists
`workpaper-proof.json` with `verified: true`.

Use it when Kestra owns scheduling, retries, task history, and downstream
workflow steps, but the calculation should stay reviewable as workbook cells and
formulas.

Official Kestra references:

- https://kestra.io/docs/how-to-guides/javascript
- https://kestra.io/plugins/plugin-script-node
- https://kestra.io/docs/concepts/blueprints

## Local Smoke

```sh
pnpm install --ignore-workspace --lockfile=false
pnpm run check
pnpm run smoke
```

The smoke test prints a verified result and writes
`.tmp/workpaper-proof.json`:

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

## Kestra Shape

1. For the self-contained path, import `blueprint.yaml`.
2. For the source-file path, upload or sync `kestra-workpaper-flow.ts` as a
   namespace file and import `flow.yml`.
3. Run the flow with the default inputs or pass quote fields from an upstream
   task.
4. Use `outputs.calculate_quote.outputFiles['workpaper-proof.json']` as the
   proof artifact for downstream review or audit storage.

The Commands task installs `@bilig/workpaper`, runs the Node script, edits
`Inputs!B2`, recalculates formula-backed summary fields, serializes WorkPaper
JSON, restores it, and verifies that restored calculated values match.

## Boundaries

This is a Bilig-owned Blueprint candidate, not an accepted official Kestra
Blueprint yet. Submit or update exactly one upstream Blueprint proposal after
checking for duplicates.

Keep Excel or another oracle in the loop for macros, pivots, external links,
locale-sensitive desktop behavior, or workbook features outside Bilig's formula
surface.
