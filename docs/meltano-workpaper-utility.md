---
title: Meltano WorkPaper Utility
published: true
description: Run Bilig WorkPaper as a Meltano utility command to validate post-ELT records with formula readback proof and a JSON artifact.
tags: meltano, elt, data-quality, workbook-api, spreadsheet-automation
canonical_url: https://proompteng.github.io/bilig/meltano-workpaper-utility.html
image: /assets/github-social-preview.png
---

# Meltano WorkPaper Utility

Use this when a Meltano job has produced records or a destination-table export,
but the downstream data-quality check should run through reviewable formulas
with JSON restore proof instead of a spreadsheet UI session.

Meltano supports custom utilities, `meltano invoke`, command shortcuts such as
`meltano invoke <plugin>:<command>`, and Hub plugin definitions with
`executable` and `commands` fields. Bilig fits as a post-ELT validation utility:
Meltano owns extraction, loading, environments, run history, and scheduling;
Bilig owns the formula workbook, recalculation, JSON persistence, restore, and
readback proof.

Official Meltano references:

- <https://docs.meltano.com/concepts/plugins/#custom-utilities>
- <https://docs.meltano.com/reference/command-line-interface/#invoke>
- <https://docs.meltano.com/reference/plugin-definition-syntax/#commands>
- <https://docs.meltano.com/reference/plugin-definition-syntax/#executable>

## Example Artifact

The runnable source lives in:

```text
examples/meltano-workpaper-utility
```

It contains:

- `meltano.yml` for the `bilig-workpaper-validator` custom utility command.
- `meltano-hub-utility-definition.yml` for the Hub-shaped utility metadata.
- `fixtures/orders.jsonl` for a no-key post-load record export.
- `meltano-workpaper-validator.ts` for JSONL parsing, formula recalculation,
  JSON restore, proof writing, and CLI argument handling.
- `scripts/check-meltano-recipe.ts` for recipe drift checks.

Run it locally:

```sh
cd examples/meltano-workpaper-utility
pnpm install --ignore-workspace --lockfile=false
pnpm run check
pnpm run typecheck
pnpm run smoke
```

The local smoke reads the JSONL export, edits `Inputs!B2` and `Inputs!B4`,
recalculates record-count and paid-amount checks, exports WorkPaper JSON,
restores it, verifies restored readback, and writes
`.tmp/workpaper-proof.json`.

## Meltano Shape

`meltano.yml` defines the utility as an `npx` executable with a command
shortcut:

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

In a Meltano project, the invocation is:

```sh
meltano invoke bilig-workpaper-validator:validate
```

The checked-in Hub-shaped definition mirrors the same `validate` command, but
this page is not claiming that Bilig is already listed on Meltano Hub.

## Output Shape

The smoke output contains:

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

## Boundary

Meltano owns plugin installation, environments, run history, schedules, and the
extract/load pipeline. Bilig owns the post-ELT formula workbook, formula
recalculation, JSON persistence, restore, and readback proof.

Keep destination-table row counts, loader acknowledgements, Meltano job
metadata, and warehouse constraints in the loop for production pipelines. Use a
warehouse query or destination export as the `--records` input when validating
real data.

## Outreach Note

If this is shared with Meltano users or proposed for Meltano Hub, lead with the
boundary:

> Meltano owns ELT orchestration and run history. Bilig owns post-ELT formula
> validation and readback proof.

Do not post it as a generic spreadsheet-engine pitch. Link the runnable example,
the smoke output, and the Hub-shaped utility definition, then ask whether a
post-ELT formula validation utility would be useful to Meltano users.
