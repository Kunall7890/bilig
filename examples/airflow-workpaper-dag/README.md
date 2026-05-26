# Airflow WorkPaper DAG

Use this when an Apache Airflow DAG needs formula-backed task outputs but the
formula model should live in a Node package with readback proof instead of an
Excel UI session or stale cached XLSX values.

The example keeps Airflow in charge of the DAG, TaskFlow tasks, retries, XCom
summary, and run history. The TypeScript step owns WorkPaper formula
recalculation, JSON restore, and the full proof file.

## Run Locally

```sh
pnpm install --ignore-workspace --lockfile=false
pnpm run check
pnpm run smoke
```

The TypeScript smoke edits `Inputs!B2`, recalculates quote formulas, persists a
WorkPaper JSON document, restores it, and verifies restored readback.

To use the DAG in Airflow, copy or mount this directory so `dags/` is visible to
Airflow and make sure the worker environment can run:

```sh
npx --no-install tsx workpaper-quote.ts --output .tmp/workpaper-proof.json
```

Run `pnpm install --ignore-workspace --lockfile=false` first so the Airflow
worker can call the local `@bilig/workpaper` and `tsx` dependencies.

## Files

- `dags/bilig_workpaper_quote_dag.py` defines the Airflow TaskFlow DAG.
- `workpaper-quote.ts` owns the WorkPaper model, formula recalculation, JSON
  restore, and proof output.
- `scripts/check-dag.ts` verifies the example wiring stays intact.

## Proof Shape

The TypeScript proof file contains:

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

The DAG returns a compact XCom summary so Airflow metadata stays small, while
the full proof stays in `.tmp/workpaper-proof.json` or whatever artifact path
your deployment uses.

Keep Excel, LibreOffice, Microsoft Graph, or a domain oracle in the loop for
macros, pivots, charts, external links, and exact desktop Excel behavior.
