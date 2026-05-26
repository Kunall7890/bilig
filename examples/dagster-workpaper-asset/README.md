# Dagster WorkPaper Asset

Use this when a Dagster asset needs formula-backed materialization metadata, but
the formula model should live in a Node WorkPaper process instead of Python
spreadsheet UI automation or stale XLSX cached values.

The example keeps Dagster in charge of the asset graph, run history, resources,
and materialization metadata. The TypeScript subprocess owns WorkPaper formula
recalculation, JSON restore, Dagster Pipes materialization metadata, and the full
proof file.

## Run Locally

```sh
pnpm install --ignore-workspace --lockfile=false
pnpm run check
pnpm run smoke
```

The smoke edits `Inputs!B2`, recalculates quote formulas, persists a WorkPaper
JSON document, restores it, and verifies restored readback.

To use the asset in Dagster, install the Python dependency and make sure the
asset worker can run the local Node dependencies:

```sh
python -m pip install -r requirements.txt
npx --no-install tsx workpaper-asset.ts --output .tmp/workpaper-proof.json
```

## Files

- `defs/bilig_workpaper_asset.py` defines the Dagster asset and
  `PipesSubprocessClient` resource.
- `workpaper-asset.ts` owns the WorkPaper model, formula recalculation, JSON
  restore, Dagster Pipes metadata message, and proof output.
- `scripts/check-asset.ts` verifies the example wiring stays intact.

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

When the script runs under Dagster Pipes, it emits `report_asset_materialization`
metadata for the calculated patch, proof, proof path, edited cell, and total. The
full proof stays in `.tmp/workpaper-proof.json` or the artifact path your
deployment uses.

Keep Excel, LibreOffice, Microsoft Graph, or a domain oracle in the loop for
macros, pivots, charts, external links, and exact desktop Excel behavior.
