---
title: Dagster WorkPaper Asset
published: true
description: Materialize Bilig WorkPaper formula output from a Dagster asset with Pipes subprocess metadata, readback proof, and JSON restore verification.
tags: dagster, data-orchestration, assets, dagster-pipes, workbook-api, spreadsheet-automation
canonical_url: https://proompteng.github.io/bilig/dagster-workpaper-asset.html
image: /assets/github-social-preview.png
---

# Dagster WorkPaper Asset

Use this when a Dagster asset needs formula-backed materialization metadata, but
the calculation should run through a Node WorkPaper subprocess with proof instead
of Excel UI automation, browser grid clicks, or stale cached XLSX formula values.

Dagster's JavaScript pipeline docs recommend `PipesSubprocessClient` for running
Node processes and mention the `@dagster-io/dagster-pipes` npm package for
production TypeScript processes. This example keeps that boundary: Dagster owns
the asset graph, resources, run history, and materialization metadata, while
Bilig owns formula recalculation, JSON serialization, restore, and readback
proof.

Official Dagster references:

- <https://docs.dagster.io/integrations/external-pipelines/javascript-pipeline>
- <https://docs.dagster.io/integrations/external-pipelines/using-dagster-pipes>
- <https://docs.dagster.io/api/dagster/metadata>

## Example Artifact

The runnable source lives in:

```text
examples/dagster-workpaper-asset
```

It contains:

- `defs/bilig_workpaper_asset.py` for the Dagster asset and
  `PipesSubprocessClient` resource
- `workpaper-asset.ts` for the TypeScript WorkPaper proof subprocess
- `scripts/check-asset.ts` for the local wiring check

Run the TypeScript proof locally:

```sh
cd examples/dagster-workpaper-asset
pnpm install --ignore-workspace --lockfile=false
pnpm run check
pnpm run smoke
```

The smoke test edits `Inputs!B2`, recalculates quote formulas, serializes the
WorkPaper document, restores it, verifies restored calculated values match,
prints JSON, and writes `.tmp/workpaper-proof.json`.

## Dagster Asset Shape

The checked-in asset uses Dagster Pipes:

```python
import dagster as dg


@dg.asset(compute_kind="javascript")
def bilig_workpaper_quote_asset(
    context: dg.AssetExecutionContext,
    pipes_subprocess_client: dg.PipesSubprocessClient,
) -> dg.MaterializeResult:
    return pipes_subprocess_client.run(
        command=[
            "npx",
            "--no-install",
            "tsx",
            "workpaper-asset.ts",
            "--output",
            ".tmp/workpaper-proof.json",
        ],
        context=context,
        extras={"quantity": 18},
    ).get_materialize_result()
```

The TypeScript process writes a JSON proof file and, when Dagster Pipes
environment variables are present, emits a `report_asset_materialization` message
with structured metadata:

- calculated patch as JSON metadata
- WorkPaper proof as JSON metadata
- proof file path
- edited cell
- calculated total

## Asset Output

The full proof file contains:

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

Keep the full proof file in an artifact path, shared volume, object store, or
asset-adjacent storage when the pipeline needs an audit trail. Keep Dagster
metadata compact enough for the event log.

## Workflow Shape

1. Dagster materializes `bilig_workpaper_quote_asset`.
2. `PipesSubprocessClient` runs the Node WorkPaper subprocess.
3. Bilig writes inputs, recalculates formulas, exports WorkPaper JSON, restores
   it, and verifies readback.
4. The subprocess emits Dagster Pipes metadata and writes the full proof file.
5. Downstream Dagster assets consume the calculated patch or proof artifact.

Dagster owns orchestration, asset state, run history, and materialization
metadata. Bilig owns the formula workbook, recalculation, JSON serialization,
restore, and readback proof.

## When This Fits

Use it for quote approvals, pricing rules, payout checks, import validation,
data-quality calculations, and asset pipelines where spreadsheet formulas are
the most reviewable representation of business logic.

Do not use it to pretend Bilig is desktop Excel. Keep Excel, LibreOffice,
Microsoft Graph, or a domain oracle in the loop for macros, pivots, charts,
external links, and exact spreadsheet UI behavior.

## Outreach Note

If this is shared in a Dagster discussion, Slack thread, or example request,
lead with the concrete boundary it solves:

> Dagster owns the asset graph and materialization metadata. Bilig owns the
> formula workbook and returns both calculated field values and readback proof.

Do not post it as a generic spreadsheet-engine pitch. Link the runnable example
and smoke output, then ask whether a formula-backed asset using JavaScript Pipes
would be useful to Dagster users.
