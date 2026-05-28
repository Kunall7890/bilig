# Meltano WorkPaper Utility

This example shows the shape of a Meltano utility command that validates
post-ELT records with `@bilig/workpaper`, writes a JSON proof file, and leaves
pipeline execution to Meltano.

Use it when Meltano owns extractor/loader orchestration, run history, plugin
installation, and job scheduling, but a downstream validation step should stay
reviewable as workbook cells and formulas.

Source path:

```text
examples/meltano-workpaper-utility
```

Official Meltano references:

- https://docs.meltano.com/concepts/plugins/#custom-utilities
- https://docs.meltano.com/reference/command-line-interface/#invoke
- https://docs.meltano.com/reference/plugin-definition-syntax/#commands
- https://docs.meltano.com/reference/plugin-definition-syntax/#executable

## Local Smoke

```sh
pnpm install --ignore-workspace --lockfile=false
pnpm run check
pnpm run typecheck
pnpm run smoke
```

The smoke test reads `fixtures/orders.jsonl`, edits the validation expectations
in a WorkPaper, recalculates formulas, writes `.tmp/workpaper-proof.json`,
restores the serialized WorkPaper document, and verifies readback:

```json
{
  "patch": {
    "command": "meltano invoke bilig-workpaper-validator:validate",
    "record_count": 4,
    "paid_amount": 301.75,
    "rejected_records": 1,
    "validation_passed": true
  },
  "proof": {
    "editedCells": ["Inputs!B2", "Inputs!B4"],
    "before": {
      "recordCountMatchesExpected": false,
      "paidAmountMatchesExpected": false
    },
    "after": {
      "recordCountMatchesExpected": true,
      "paidAmountMatchesExpected": true,
      "rejectedWithinMax": true
    },
    "afterRestore": {
      "recordCountMatchesExpected": true
    },
    "verified": true
  }
}
```

## Meltano Shape

`meltano.yml` defines a custom utility command:

```yaml
plugins:
  utilities:
    - name: bilig-workpaper-validator
      namespace: bilig_workpaper
      executable: npx
      commands:
        validate:
          args: '--yes --package @bilig/workpaper@latest --package tsx@4.21.0 tsx meltano-workpaper-validator.ts --records output/orders.jsonl --expected-record-count 4 --expected-paid-amount 301.75 --output .tmp/workpaper-proof.json'
```

In a real Meltano project, invoke it after an ELT step has produced a JSONL
export or after a destination table has been exported for validation:

```sh
meltano invoke bilig-workpaper-validator:validate
```

The checked-in `meltano-hub-utility-definition.yml` mirrors the same command in
Meltano Hub plugin-definition syntax. Keep it as review material until a Hub
submission is ready; do not treat this example as an accepted Meltano plugin.

## Boundaries

Meltano owns plugin installation, run history, schedules, environments, and the
extract/load pipeline. Bilig owns the post-ELT formula workbook, formula
recalculation, JSON serialization, restore, and readback proof.

Keep destination-table row counts, loader acknowledgements, Meltano job
metadata, and warehouse constraints in the loop for production pipelines. Use a
warehouse query or destination export as the `--records` input when validating
real data.
