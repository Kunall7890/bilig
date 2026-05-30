---
title: XLSX Cache Doctor GitHub Action
published: true
description: Detect stale cached XLSX formula values in pull requests before a backend reads the wrong number.
tags: github-actions, xlsx, formulas, ci, node
canonical_url: https://proompteng.github.io/bilig/xlsx-cache-doctor-github-action.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# XLSX Cache Doctor GitHub Action

Use this when a repository keeps `.xlsx` fixtures, pricing models, report
templates, or approval workbooks under version control and CI needs to catch
stale cached formula values before a service reads them.

The action wraps the same command you can run locally:

```sh
npm exec --package @bilig/xlsx-formula-recalc@latest -- xlsx-cache-doctor fixtures/pricing.xlsx --json
```

`xlsx-cache-doctor` imports the workbook, lists formula cells, recomputes every
formula by default, reports stale cached values, and returns suggested `--read`
targets for the follow-up recalculation proof. It also reports
`uninspectedFormulaCellCount` when a caller intentionally sets a smaller
`--inspect-limit`. It does not write a new workbook.

## Workflow

```yaml
name: workbook-cache

on:
  pull_request:
    paths:
      - "**/*.xlsx"
  workflow_dispatch:

permissions:
  contents: read

jobs:
  xlsx-cache:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v5

      - uses: proompteng/bilig@v1
        id: cache-doctor
        with:
          workbook: fixtures/pricing.xlsx
          json-output: ${{ runner.temp }}/pricing.cache-doctor.json
          fail-on-stale: "true"

      - run: |
          echo "formula cells: ${{ steps.cache-doctor.outputs.formula-count }}"
          echo "stale values: ${{ steps.cache-doctor.outputs.stale-count }}"
          echo "uninspected formulas: ${{ steps.cache-doctor.outputs.uninspected-count }}"
          echo "suggested reads: ${{ steps.cache-doctor.outputs.suggested-reads }}"

      - uses: actions/upload-artifact@v4
        if: always()
        with:
          name: xlsx-cache-doctor-report
          path: ${{ steps.cache-doctor.outputs.json }}
```

See
[`examples/xlsx-cache-doctor-ci`](https://github.com/proompteng/bilig/tree/main/examples/xlsx-cache-doctor-ci)
for a complete fixture, workflow, and committed JSON report. The example
fixture has a formula cell saved with a stale cached value so the action failure
is easy to inspect before you add it to a real repository.

The root `proompteng/bilig@v1` action is the canonical install path for
Marketplace-style discovery. The subdirectory action remains available as
`proompteng/bilig/actions/xlsx-cache-doctor@v1` for users who already copied
that path.

## Inputs

| Input             | Default | Use                                                                 |
| ----------------- | ------- | ------------------------------------------------------------------- |
| `workbook`        |         | Path to the workbook to inspect.                                    |
| `package-version` | latest  | npm version or dist-tag for `@bilig/xlsx-formula-recalc`.         |
| `inspect-limit`   | `all`   | Formula cells to recompute during inspection. Use `all` or a positive integer. |
| `json-output`     |         | Optional path for the JSON report.                                  |
| `fail-on-stale`   | `false` | Fail the job when inspected formula cells have stale cached values. |

## Outputs

| Output            | Meaning                                                           |
| ----------------- | ----------------------------------------------------------------- |
| `json`            | Path to the JSON report written by `xlsx-cache-doctor`.           |
| `formula-count`   | Total formula cells found in the workbook.                        |
| `stale-count`     | Inspected formula cells where cached and recalculated values differ. |
| `uninspected-count` | Formula cells skipped because `inspect-limit` was lower than `formula-count`. |
| `suggested-reads` | Comma-separated cells to use in the follow-up `xlsx-recalc` proof. |

## Follow-Up Proof

When the action finds stale cached values, run a narrow proof on the cells that
matter to the backend:

```sh
npm exec --package @bilig/xlsx-formula-recalc@latest -- xlsx-recalc fixtures/pricing.xlsx \
  --set Inputs!B2=48 \
  --read Summary!B7 \
  --out fixtures/pricing.recalculated.xlsx \
  --json
```

That command writes a recalculated workbook and returns the exact readback
values. Keep the action as the pull-request detector and the `xlsx-recalc`
command as the explicit proof before production adoption.
