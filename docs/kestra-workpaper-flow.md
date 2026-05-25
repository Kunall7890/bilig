---
title: Kestra WorkPaper Node flow
published: true
description: Run Bilig WorkPaper inside a Kestra Node Commands flow to calculate workflow fields with formula readback, output-file proof, and JSON restore verification.
tags: kestra, workflow-automation, node, workbook-api, spreadsheet-automation
canonical_url: https://proompteng.github.io/bilig/kestra-workpaper-flow.html
image: /assets/github-social-preview.png
---

# Kestra WorkPaper Node flow

Use this when a Kestra flow needs formula-backed workflow fields, but the
calculation should run through a Node package and leave a proof artifact instead
of relying on Excel UI automation, browser grid clicks, or stale cached XLSX
formula values.

Kestra's Node plugin runs JavaScript or TypeScript code inline or from files,
can install npm packages, and can expose output files for downstream tasks.
Kestra Blueprints are the curated reusable workflow-template path, so keep this
as a Bilig-owned proof first and only submit one focused Blueprint if the fit is
real.

Official Kestra references:

- <https://kestra.io/docs/how-to-guides/javascript>
- <https://kestra.io/plugins/plugin-script-node>
- <https://kestra.io/docs/concepts/blueprints>
- <https://kestra.io/docs/llms.txt>

## Example Artifact

The runnable source lives in:

```text
examples/kestra-workpaper-flow
```

It contains:

- `flow.yml` for the Kestra `io.kestra.plugin.scripts.node.Commands` flow
- `kestra-workpaper-flow.ts` for the TypeScript WorkPaper proof script
- `scripts/check-flow.ts` for the local flow wiring check

Run it locally:

```sh
cd examples/kestra-workpaper-flow
pnpm install --ignore-workspace --lockfile=false
pnpm run check
pnpm run smoke
```

The smoke test edits `Inputs!B2`, recalculates quote formulas, serializes the
WorkPaper document, restores it, verifies that restored calculated values match,
prints JSON, and writes `.tmp/workpaper-proof.json`.

## Kestra Flow Shape

The checked-in `flow.yml` uses a Node Commands task:

```yaml
tasks:
  - id: calculate_quote
    type: io.kestra.plugin.scripts.node.Commands
    namespaceFiles:
      enabled: true
    taskRunner:
      type: io.kestra.plugin.scripts.runner.docker.Docker
    containerImage: node:24-slim
    beforeCommands:
      - npm install @bilig/workpaper@latest
      - npm install tsx
    outputFiles:
      - workpaper-proof.json
    commands:
      - npx tsx kestra-workpaper-flow.ts --quantity "{{ inputs.quantity }}" --output workpaper-proof.json
```

Upload or sync `kestra-workpaper-flow.ts` as a namespace file, import
`flow.yml`, and run it with the default inputs or quote fields from another
task.

## Flow Output

The proof file contains:

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
output-file path.

## Workflow Shape

1. Trigger `bilig_workpaper_quote` from a schedule, webhook, API call, or parent
   flow.
2. Pass record values such as `quantity`, `unit_price`, `discount_rate`,
   `tax_rate`, and `unit_cost`.
3. Feed `patch` into the next task or application writeback step.
4. Keep `workpaper-proof.json` as the audit artifact when the workflow needs
   readback evidence.

Kestra owns scheduling, retries, task history, Docker execution, and downstream
workflow orchestration. Bilig owns the formula workbook, recalculation, JSON
serialization, restore, and readback proof.

## When This Fits

Use it for quote approvals, pricing rules, discount calculations, payout checks,
import validation, and workflow steps where spreadsheet formulas are the most
reviewable representation of business logic.

Do not use it to pretend Bilig is desktop Excel. Keep Excel, LibreOffice,
Microsoft Graph, or a domain oracle in the loop for macros, pivots, charts,
external links, and exact spreadsheet UI behavior.

## Outreach Note

If this is shared in a Kestra issue, Slack thread, or Blueprint proposal, lead
with the concrete boundary it solves:

> Kestra owns orchestration and output-file routing. Bilig owns the formula
> workbook and returns both calculated field values and readback proof.

Do not post it as a generic spreadsheet-engine pitch. Link the runnable example
and smoke output, then ask whether a formula-backed Node Commands flow would be
useful as a Kestra Blueprint.
