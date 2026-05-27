# Airbyte WorkPaper Validation

Use this when an Airbyte sync has produced records and state, but the post-sync
business validation should run through reviewable formulas with readback proof
instead of a spreadsheet UI session or stale cached XLSX values.

The example keeps Airbyte in charge of extraction, replication, sync modes,
record messages, and checkpoint state. Bilig owns the post-sync WorkPaper model,
formula recalculation, JSON persistence, restore, and proof object.

Official Airbyte references:

- <https://docs.airbyte.com/platform/understanding-airbyte/airbyte-protocol>
- <https://docs.airbyte.com/platform/using-airbyte/core-concepts/sync-modes/incremental-append-deduped>

## Run Locally

```sh
pnpm install --ignore-workspace --lockfile=false
pnpm run check
pnpm run typecheck
pnpm run smoke
```

The smoke test reads `fixtures/orders-airbyte-messages.jsonl`, converts the
Airbyte-style `RECORD` and `STATE` messages into a WorkPaper, writes the
committed state cursor into `Inputs!B2`, recalculates summary formulas, persists
the WorkPaper JSON document, restores it, and verifies restored readback.

## Files

- `fixtures/orders-airbyte-messages.jsonl` contains the sample Airbyte record
  and state stream.
- `src/airbyte-workpaper-validation.ts` owns JSONL parsing, record validation,
  WorkPaper formula recalculation, JSON persistence, restore, and proof output.
- `src/smoke.ts` runs the fixture and asserts the expected patch.
- `scripts/check-airbyte-recipe.ts` verifies the example stays wired to the
  docs and proof contract.

## Proof Shape

The smoke output contains:

```json
{
  "patch": {
    "stream": "orders",
    "committed_state_cursor": "2026-05-27T10:10:00Z",
    "record_count": 4,
    "gross_amount": 315,
    "paid_amount": 301.75,
    "rejected_records": 1,
    "validation_passed": true
  },
  "proof": {
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
    "verified": true
  }
}
```

Airbyte owns the sync and checkpoint semantics. Use this example after records
are available from Airbyte, a warehouse export, object storage, or a job log.
Bilig is only the formula-backed validation layer.

Keep warehouse constraints, source-specific state semantics, destination write
acknowledgement, and domain data-quality checks in the loop for production
pipelines.
