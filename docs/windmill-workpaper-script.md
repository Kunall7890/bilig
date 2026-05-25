---
title: Windmill WorkPaper TypeScript script
published: true
description: Use Bilig WorkPaper inside a Windmill TypeScript script to calculate workflow fields with formula readback, JSON persistence, and restore proof.
tags: windmill, typescript, workflow-automation, workbook-api, spreadsheet-automation
canonical_url: https://proompteng.github.io/bilig/windmill-workpaper-script.html
image: /assets/github-social-preview.png
---

# Windmill WorkPaper TypeScript script

Use this when a Windmill workflow needs spreadsheet-shaped business logic, but
the formula state should be edited and verified through a TypeScript API instead
of Excel UI automation, browser grid clicks, or stale cached XLSX formula
values.

Windmill's TypeScript scripts run with a `main` entrypoint. Windmill documents
that TypeScript dependencies can be resolved directly from script imports, with
lockfiles generated from those imports at deployment time.

Official Windmill references:

- <https://www.windmill.dev/docs/advanced/dependencies_in_typescript>
- <https://www.windmill.dev/docs/getting_started/scripts_quickstart/typescript>
- <https://www.windmill.dev/docs/script_editor>
- <https://www.windmill.dev/docs/core_concepts/workflows_as_code>

## Example Artifact

The runnable source lives in:

```text
examples/windmill-workpaper-script
```

It contains:

- `src/workpaper-script.ts` for the Windmill-style `main` script
- `src/smoke.ts` for a no-Windmill local proof

Run it locally:

```sh
cd examples/windmill-workpaper-script
pnpm install --ignore-workspace --lockfile=false
pnpm run typecheck
pnpm run smoke
```

The smoke test edits `Inputs!B2`, recalculates quote formulas, serializes the
WorkPaper document, restores it, and verifies that the restored calculated
values match.

## Windmill Script

Paste the contents of `src/workpaper-script.ts` into a Windmill TypeScript
script.

The script shape is intentionally boring:

```ts
export async function main(
  quantity = 18,
  unitPrice = 125,
  discountRate = 0.1,
  taxRate = 0.08,
  unitCost = 52,
  previousQuantity = 12,
) {
  // Build a WorkPaper, edit one input cell, read dependent formulas, export
  // JSON, restore it, and return both a patch and proof.
}
```

Windmill can infer inputs from the `main` parameters and resolve
`@bilig/workpaper` from the script import. If your workspace requires pinned
versions, pin or lock the dependency using the Windmill dependency workflow your
team already uses.

## Script Output

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

## Workflow Shape

1. Trigger the Windmill flow from a schedule, webhook, approval, or manual run.
2. Run the Bilig WorkPaper TypeScript script with record values such as
   `quantity`, `unitPrice`, `discountRate`, `taxRate`, and `unitCost`.
3. Feed `patch` into the next step that writes calculated fields back to your
   system of record.
4. Keep `proof` in logs, audit storage, or a downstream approval step when the
   workflow needs readback evidence.

That keeps Windmill in charge of orchestration, retries, workers, and workflow
routing. Bilig owns the formula workbook, recalculation, JSON serialization,
restore, and readback proof.

## When This Fits

Use it for quote approvals, pricing rules, discount calculations, payout checks,
import validation, and operational fields that should be reviewable as formulas
but executed inside a Windmill workflow.

Do not use it to pretend Bilig is desktop Excel. Keep Excel, LibreOffice,
Microsoft Graph, or a domain oracle in the loop for macros, pivots, charts,
external links, and exact spreadsheet UI behavior.

## Outreach Note

If this is shared in a Windmill issue, discussion, or Hub submission, lead with
the concrete boundary it solves:

> Windmill owns the workflow. Bilig owns the formula workbook and returns both
> calculated field values and readback proof.

Do not post it as a generic spreadsheet-engine pitch. Link the runnable example
and smoke output, then ask whether a formula-backed workflow field script would
be useful as a Windmill Hub example.
