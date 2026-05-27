# Airbyte WorkPaper Validation

Use this when an Airbyte sync has produced records and state, but the post-sync
business validation should run through reviewable formulas with readback proof
instead of a spreadsheet UI session or stale cached XLSX values.

The example keeps Airbyte in charge of extraction, replication, sync modes,
record messages, checkpoint state, destination writes, and job metadata. Bilig
owns the post-sync WorkPaper model, formula recalculation, JSON persistence,
restore, and proof object.

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

The smoke test reads both `fixtures/orders-airbyte-messages.jsonl` and
`fixtures/orders-airbyte-global-state-messages.jsonl`. It converts
Airbyte-style `RECORD` and `STATE` messages into a WorkPaper, proves both
stream-scoped and global state cursor readback, writes the committed state
cursor into `Inputs!B2`, recalculates summary formulas, persists the WorkPaper
JSON document, restores it, and verifies restored readback.

## Files

- `fixtures/orders-airbyte-messages.jsonl` contains the sample Airbyte record
  stream with `STREAM` state.
- `fixtures/orders-airbyte-global-state-messages.jsonl` contains the same
  record stream with `GLOBAL` state and stream-specific state inside
  `global.stream_states`.
- `src/airbyte-workpaper-validation.ts` owns JSONL parsing, record validation,
  WorkPaper formula recalculation, JSON persistence, restore, and proof output.
- `src/smoke.ts` runs the fixture and asserts the expected patch.
- `scripts/check-airbyte-recipe.ts` verifies the example stays wired to the
  docs and proof contract.

## Proof Shape

The smoke output contains separate proofs for both state shapes:

```json
{
  "streamState": {
    "patch": {
      "stream": "orders",
      "state_type": "STREAM",
      "committed_state_cursor": "2026-05-27T10:10:00Z",
      "validation_passed": true
    },
    "proof": {
      "stateCursorSource": "state.stream.stream_state.cursor",
      "afterRestore": {
        "stateCursorMatchesRecords": true
      },
      "verified": true
    }
  },
  "globalState": {
    "patch": {
      "stream": "orders",
      "state_type": "GLOBAL",
      "committed_state_cursor": "2026-05-27T10:10:00Z",
      "validation_passed": true
    },
    "proof": {
      "stateCursorSource": "state.global.stream_states[].stream_state.cursor",
      "afterRestore": {
        "stateCursorMatchesRecords": true
      },
      "verified": true
    }
  }
}
```

Airbyte owns the sync and checkpoint semantics. Use this example after records
are available from Airbyte, a warehouse export, object storage, or a job log.
Bilig is only the formula-backed validation layer.

Keep warehouse constraints, source-specific state semantics, destination write
acknowledgement, Airbyte job metadata, and domain data-quality checks in the
loop for production pipelines. Destination-level validation should read the
destination tables and job metadata directly; the JSONL fixture here is only a
portable recipe artifact.
