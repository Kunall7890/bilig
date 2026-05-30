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

Install from GitHub Marketplace:
<https://github.com/marketplace/actions/xlsx-cache-doctor>

The action wraps the same command you can run locally:

```sh
npm exec --package @bilig/xlsx-formula-recalc@latest -- xlsx-cache-doctor fixtures/pricing.xlsx --json
```

`xlsx-cache-doctor` imports the workbook, lists formula cells, recomputes every
formula by default, reports stale cached values, and returns suggested `--read`
targets for the follow-up recalculation proof. It also reports
`uninspectedFormulaCellCount` when a caller intentionally sets a smaller
`--inspect-limit`. It does not write a new workbook.

## Generate The Workflow

If you want the pull-request workflow without copying YAML by hand, generate it
from npm:

```sh
mkdir -p .github/workflows
npm exec --package @bilig/xlsx-formula-recalc@latest -- \
  xlsx-cache-doctor --print-github-action "**/*.xlsx" \
  > .github/workflows/xlsx-cache-doctor.yml
```

That prints a read-only workflow using the root `proompteng/bilig@v1` action,
`actions/checkout@v5` with enough history for changed-file detection, and an
uploaded JSON artifact. It does not need secrets. By default, the generated
pull-request workflow scans changed `.xlsx` files matching the glob. Use
`--changed-files-only false` for a scheduled or manual-dispatch full scan,
`--inspect-limit 50` for a deliberately sampled first pass, `--json-output` if
your artifact path is different, and `--fail-on-stale true` when you are ready
for the check to block pull requests. The generated workflow starts in
report-only mode.

## Workflow

```yaml
name: workbook-cache

on:
  pull_request:
    paths:
      - '**/*.xlsx'
  workflow_dispatch:

permissions:
  contents: read

jobs:
  xlsx-cache:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v5
        with:
          fetch-depth: 0

      - uses: proompteng/bilig@v1
        id: cache-doctor
        with:
          workbooks: '**/*.xlsx'
          changed-files-only: 'true'
          json-output: ${{ runner.temp }}/pricing.cache-doctor.json
          fail-on-stale: 'true'

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

For the live GitHub reviewer path, use the
[XLSX Cache Doctor demo PR](https://github.com/proompteng/xlsx-cache-doctor-demo/pull/1).
The green check runs `proompteng/bilig@v1`, inspects one changed workbook,
asserts 60 formula cells and 1 stale cached formula value, and uploads the JSON
report artifact.

The job summary is meant to be the reviewer artifact. It shows workbook count,
formula count, stale count, the first stale cells with cached and recalculated
values, and a copy-paste `xlsx-recalc --read ... --json` command for the
follow-up check. The action also writes GitHub warning annotations for the first
stale cells so the failure is visible before someone opens the JSON artifact.

The root `proompteng/bilig@v1` action is the canonical install path and is
published as
[XLSX Cache Doctor on GitHub Marketplace](https://github.com/marketplace/actions/xlsx-cache-doctor).
The subdirectory action remains available as
`proompteng/bilig/actions/xlsx-cache-doctor@v1` for users who already copied
that path.

For stricter supply-chain policy, pin the action to a full commit SHA in your
workflow after you test the example:

```yaml
- uses: proompteng/bilig@<full-commit-sha>
```

Keep `permissions: contents: read` unless your workflow adds its own write
steps. The cache doctor itself does not need secrets, pull-request comments, or
a write token.

## Inputs

| Input                | Default | Use                                                                              |
| -------------------- | ------- | -------------------------------------------------------------------------------- |
| `workbook`           |         | Path to one workbook to inspect. Kept for existing copied workflows.             |
| `workbooks`          |         | Glob, comma list, or newline list of XLSX workbooks to inspect.                  |
| `changed-files-only` | `false` | Only inspect matched XLSX files changed in the pull request or current git diff. |
| `package-version`    | latest  | npm version or dist-tag for `@bilig/xlsx-formula-recalc`.                        |
| `inspect-limit`      | `all`   | Formula cells to recompute during inspection. Use `all` or a positive integer.   |
| `json-output`        |         | Optional path for the JSON report.                                               |
| `fail-on-stale`      | `false` | Fail the job when inspected formula cells have stale cached values.              |

## Outputs

| Output              | Meaning                                                                                         |
| ------------------- | ----------------------------------------------------------------------------------------------- |
| `json`              | Path to the JSON report written by `xlsx-cache-doctor`.                                         |
| `workbook-count`    | Number of matched XLSX workbooks inspected.                                                     |
| `formula-count`     | Total formula cells found across inspected workbooks.                                           |
| `stale-count`       | Inspected formula cells where cached and recalculated values differ.                            |
| `uninspected-count` | Formula cells skipped because `inspect-limit` was lower than `formula-count`.                   |
| `suggested-reads`   | First 25 workbook-qualified cells for the follow-up proof; the JSON report keeps the full list. |

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
