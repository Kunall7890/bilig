---
title: Prefect WorkPaper flow
published: true
description: Run Bilig WorkPaper from a Prefect flow to calculate workflow fields with formula readback, JSON persistence, and restore proof.
tags: prefect, data-orchestration, workflow-automation, workbook-api, spreadsheet-automation
canonical_url: https://proompteng.github.io/bilig/prefect-workpaper-flow.html
image: /assets/github-social-preview.png
---

# Prefect WorkPaper flow

Use this when a Prefect flow needs formula-backed workflow fields, but the
calculation should run through a Node package with a proof object instead of
Excel UI automation, browser grid clicks, or stale cached XLSX formula values.

Prefect flows and tasks are Python functions. This example keeps that boundary:
Prefect owns orchestration, retries, logs, deployments, and scheduling, while a
small Node step owns WorkPaper formula recalculation and JSON restore proof.

Official Prefect references:

- <https://docs.prefect.io/v3/concepts/flows>
- <https://docs.prefect.io/v3/how-to-guides/workflows/write-and-run>
- <https://docs.prefect.io/v3/concepts/deployments>
- <https://docs.prefect.io/llms.txt>

## Example Artifact

The runnable source lives in:

```text
examples/prefect-workpaper-flow
```

It contains:

- `flow.py` for the Prefect `@flow` and retrying `@task`
- `workpaper-quote.ts` for the TypeScript WorkPaper proof step
- `scripts/check-flow.ts` for the local wiring check

Run the TypeScript proof locally:

```sh
cd examples/prefect-workpaper-flow
pnpm install --ignore-workspace --lockfile=false
pnpm run check
pnpm run smoke
```

Run the Prefect wrapper:

```sh
python3 -m venv .venv
. .venv/bin/activate
pip install -r requirements.txt
python flow.py
```

The smoke test edits `Inputs!B2`, recalculates quote formulas, serializes the
WorkPaper document, restores it, verifies that restored calculated values match,
prints JSON, and writes `.tmp/workpaper-proof.json`.

## Prefect Flow Shape

The checked-in `flow.py` keeps the workflow boundary explicit:

```python
from prefect import flow, task


@task(retries=2, retry_delay_seconds=5)
def calculate_quote_fields(quantity: int = 18) -> dict:
    # Calls: npx tsx workpaper-quote.ts --quantity ...
    # Reads: .tmp/workpaper-proof.json
    # Raises if proof.verified is false.
    ...


@flow(name="bilig-workpaper-quote")
def bilig_workpaper_quote_flow(quantity: int = 18) -> dict:
    return calculate_quote_fields(quantity=quantity)
```

Use a deployment when the flow should run on your Prefect worker pool. Make
Node and npm dependencies part of the worker image, startup command, or pull
step so the task can run the checked-in TypeScript WorkPaper step.

## Flow Output

The proof contains:

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

The full local smoke output includes all calculated summary fields and the
proof-file path.

## Workflow Shape

1. Run `bilig-workpaper-quote` from a schedule, deployment, event, API call, or
   parent flow.
2. Pass record values such as `quantity`, `unit_price`, `discount_rate`,
   `tax_rate`, and `unit_cost`.
3. Feed `patch` into the next task or application writeback step.
4. Keep `proof` in logs, artifacts, or an audit table when the workflow needs
   readback evidence.

Prefect owns scheduling, task retries, deployments, worker selection, and run
history. Bilig owns the formula workbook, recalculation, JSON serialization,
restore, and readback proof.

## When This Fits

Use it for quote approvals, pricing rules, discount calculations, payout checks,
import validation, and data workflows where spreadsheet formulas are the most
reviewable representation of business logic.

Do not use it to pretend Bilig is desktop Excel. Keep Excel, LibreOffice,
Microsoft Graph, or a domain oracle in the loop for macros, pivots, charts,
external links, and exact spreadsheet UI behavior.

## Outreach Note

If this is shared in a Prefect issue, Slack thread, or examples PR, lead with
the concrete boundary it solves:

> Prefect owns orchestration and task history. Bilig owns the formula workbook
> and returns both calculated field values and readback proof.

Do not post it as a generic spreadsheet-engine pitch. Link the runnable example
and smoke output, then ask whether a formula-backed flow that calls a Node
WorkPaper step would be useful as a Prefect example.

