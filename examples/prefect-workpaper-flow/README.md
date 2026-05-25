# Prefect WorkPaper Flow

Use this when a Prefect flow needs formula-backed workflow fields but the
formula model should live in a Node package with readback proof instead of an
Excel UI session or stale cached XLSX values.

The example keeps Prefect in charge of flow/task orchestration and calls a
small TypeScript WorkPaper step with `npx tsx`.

## Run Locally

```sh
pnpm install --ignore-workspace --lockfile=false
pnpm run check
pnpm run smoke
```

The TypeScript smoke edits `Inputs!B2`, recalculates quote formulas, persists a
WorkPaper JSON document, restores it, and verifies restored readback.

To run the Prefect wrapper:

```sh
python3 -m venv .venv
. .venv/bin/activate
pip install -r requirements.txt
python flow.py
```

Run `pnpm install --ignore-workspace --lockfile=false` first so the Prefect
process can call the local `@bilig/workpaper` and `tsx` dependencies.

## Files

- `flow.py` defines `bilig_workpaper_quote_flow` and a retrying Prefect task.
- `workpaper-quote.ts` owns the WorkPaper model, formula recalculation, JSON
  restore, and proof output.
- `scripts/check-flow.ts` verifies the example wiring stays intact.

## Proof Shape

The output contains:

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

Keep Excel, LibreOffice, Microsoft Graph, or a domain oracle in the loop for
macros, pivots, charts, external links, and exact desktop Excel behavior.

