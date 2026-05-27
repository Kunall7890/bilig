---
title: Airbyte WorkPaper Validation
published: true
description: Validate Airbyte post-sync record and state outputs with Bilig WorkPaper formulas, JSON restore proof, and a compact patch object.
tags: airbyte, data-replication, data-quality, workbook-api, spreadsheet-automation
canonical_url: https://proompteng.github.io/bilig/airbyte-workpaper-validation.html
image: /assets/github-social-preview.png
---

# Airbyte WorkPaper Validation

Use this when an Airbyte sync has produced records and state, but the post-sync
business validation should run through reviewable formulas with readback proof
instead of Excel UI automation, browser grid clicks, or stale cached XLSX
formula values.

The Airbyte Protocol describes record and state messages as JSON envelopes. A
source `read` emits `AirbyteRecordMessage` values and `AirbyteStateMessage`
checkpoints, and state lets the next sync resume from a checkpoint instead of
starting over. Incremental syncs use a cursor to determine which records are new
or updated.

Official Airbyte references:

- <https://docs.airbyte.com/platform/understanding-airbyte/airbyte-protocol>
- <https://docs.airbyte.com/platform/using-airbyte/core-concepts/sync-modes/incremental-append-deduped>

## Example Artifact

The runnable source lives in:

```text
examples/airbyte-workpaper-validation
```

It contains:

- `fixtures/orders-airbyte-messages.jsonl` for an Airbyte-style `RECORD` and
  `STATE` stream.
- `src/airbyte-workpaper-validation.ts` for JSONL parsing, WorkPaper formula
  recalculation, JSON restore, and proof output.
- `src/smoke.ts` for the local verification path.
- `scripts/check-airbyte-recipe.ts` for recipe drift checks.

Run it locally:

```sh
cd examples/airbyte-workpaper-validation
pnpm install --ignore-workspace --lockfile=false
pnpm run check
pnpm run typecheck
pnpm run smoke
```

The smoke test reads the JSONL stream, writes the committed state cursor into
`Inputs!B2`, writes the numeric cursor proof into `Inputs!B3`, writes expected
paid amount and record count into `Inputs!B4:B5`, recalculates the workbook,
exports WorkPaper JSON, restores it, and verifies restored readback.

## Output Shape

The returned patch is deliberately compact enough for a downstream job step:

```json
{
  "stream": "orders",
  "committed_state_cursor": "2026-05-27T10:10:00Z",
  "record_count": 4,
  "gross_amount": 315,
  "paid_amount": 301.75,
  "rejected_records": 1,
  "validation_passed": true
}
```

The proof keeps the spreadsheet evidence:

```json
{
  "editedCells": ["Inputs!B2", "Inputs!B3", "Inputs!B4", "Inputs!B5"],
  "before": {
    "stateCursorMatchesRecords": false
  },
  "after": {
    "stateCursorMatchesRecords": true,
    "paidAmountMatchesExpected": true,
    "recordCountMatchesExpected": true
  },
  "afterRestore": {
    "stateCursorMatchesRecords": true
  },
  "persistedDocumentBytes": 2170,
  "verified": true
}
```

## Boundary

Airbyte owns extraction, replication, sync mode selection, destination writes,
and checkpoint semantics. Bilig owns the post-sync formula workbook, formula
recalculation, JSON persistence, restore, and readback proof.

This is not an Airbyte connector and it is not an official Airbyte integration.
Use it after records are available from Airbyte, a warehouse export, object
storage, or a job log.

Keep source-specific state semantics, destination acknowledgement, warehouse
constraints, and domain data-quality checks in the loop for production
pipelines.

## Outreach Note

If this is shared in an Airbyte issue, Slack thread, community post, or example
discussion, lead with the boundary:

> Airbyte owns sync and checkpointing. Bilig owns post-sync formula validation
> and readback proof.

Do not post it as a generic spreadsheet-engine pitch. Link the runnable example
and smoke output, then ask whether a post-sync formula validation recipe would
be useful to Airbyte users.
