---
title: Airflow WorkPaper DAG
published: true
description: Run Bilig WorkPaper from an Apache Airflow DAG to calculate task outputs with formula readback, compact XCom proof, and JSON restore verification.
tags: airflow, data-orchestration, workflow-automation, workbook-api, spreadsheet-automation
canonical_url: https://proompteng.github.io/bilig/airflow-workpaper-dag.html
image: /assets/github-social-preview.png
---

# Airflow WorkPaper DAG

Use this when an Apache Airflow DAG needs formula-backed task outputs, but the
calculation should run through a Node package with a proof object instead of
Excel UI automation, browser grid clicks, or stale cached XLSX formula values.

Airflow's TaskFlow API lets DAG authors write tasks as decorated Python
functions and passes task outputs through XCom. This example keeps that
boundary: Airflow owns the DAG, retries, task state, XCom summary, and run
history, while a small Node step owns WorkPaper formula recalculation and JSON
restore proof.

Official Airflow references, checked against the current Airflow 3.2 docs while
keeping this example compatible with Airflow 2.10+:

- <https://airflow.apache.org/docs/apache-airflow/stable/tutorial/taskflow.html>
- <https://airflow.apache.org/docs/apache-airflow/stable/core-concepts/xcoms.html>
- <https://airflow.apache.org/docs/apache-airflow/stable/howto/operator/python.html>
- <https://airflow.apache.org/docs/apache-airflow/stable/_api/airflow/example_dags/index.html>

## Example Artifact

The runnable source lives in:

```text
examples/airflow-workpaper-dag
```

It contains:

- `dags/bilig_workpaper_quote_dag.py` for the Airflow TaskFlow DAG
- `workpaper-quote.ts` for the TypeScript WorkPaper proof step
- `scripts/check-dag.ts` for the local wiring check

Run the TypeScript proof locally:

```sh
cd examples/airflow-workpaper-dag
pnpm install --ignore-workspace --lockfile=false
pnpm run check
pnpm run smoke
```

The smoke test edits `Inputs!B2`, recalculates quote formulas, serializes the
WorkPaper document, restores it, verifies that restored calculated values match,
prints JSON, and writes `.tmp/workpaper-proof.json`.

## Airflow DAG Shape

The checked-in DAG keeps XCom compact:

```python
try:
    from airflow.sdk import dag, task
except ImportError:
    from airflow.decorators import dag, task


@dag(dag_id="bilig_workpaper_quote", schedule=None, catchup=False)
def bilig_workpaper_quote_dag() -> None:
    @task(retries=2)
    def calculate_quote_fields() -> dict:
        # Calls: npx --no-install tsx workpaper-quote.ts --quantity ...
        # Reads: .tmp/workpaper-proof.json
        # Returns a compact patch/proof summary through XCom.
        ...

    @task
    def verify_formula_proof(result: dict) -> dict:
        # Fails the run when readback or restore proof does not match.
        ...

    verify_formula_proof(calculate_quote_fields())
```

Mount or copy the example directory so Airflow can see `dags/`, then make Node
and npm dependencies part of the worker image, startup command, or deployment
artifact. The DAG can run as a scheduled workflow, manual run, or upstream task
dependency.

## DAG Output

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

The DAG returns the calculated `patch` plus a compact proof summary through
XCom. Keep the complete WorkPaper proof file in an artifact path, shared volume,
object store, or logs when the workflow needs an audit trail.

## Workflow Shape

1. Run `bilig_workpaper_quote` from a schedule, manual trigger, dataset event,
   API trigger, or parent DAG.
2. Pass record values such as `quantity`, `unit_price`, `discount_rate`,
   `tax_rate`, and `unit_cost`.
3. Feed `patch` into the next task or application writeback step.
4. Keep the proof file as the audit artifact when the workflow needs readback
   evidence.

Airflow owns scheduling, retries, dependency graph state, XCom summary, and run
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

If this is shared in an Airflow issue, Slack thread, or example discussion, lead
with the concrete boundary it solves:

> Airflow owns the DAG and task history. Bilig owns the formula workbook and
> returns both calculated field values and readback proof.

Do not post it as a generic spreadsheet-engine pitch. Link the runnable example
and smoke output, then ask whether a formula-backed DAG that calls a Node
WorkPaper step would be useful to Airflow users.
