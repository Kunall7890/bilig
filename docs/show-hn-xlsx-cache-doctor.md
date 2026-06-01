---
title: 'Show HN: XLSX Cache Doctor catches stale Excel formula values in CI'
description: 'A narrow Bilig launch note for finding stale cached formula values in XLSX files before Node services, queues, agents, or pull requests trust them.'
---

# Show HN: XLSX Cache Doctor catches stale Excel formula values in CI

When a Node script edits an input cell in an `.xlsx` file, common file libraries
can save the workbook while leaving old formula results in place. The next API,
queue worker, test, or agent can read the cached number and think it is fresh.

XLSX Cache Doctor is the smallest Bilig path for that problem. It recalculates a
workbook in Node, compares the cached formula value with the fresh value, and
prints the exact cells that changed.

```sh
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- bilig-evaluate --door xlsx-cache --json
```

Expected shape:

```json
{
  "schemaVersion": "bilig-evaluator.v1",
  "door": "xlsx-cache",
  "verified": true,
  "evidence": {
    "target": "Summary!B2",
    "before": 60000,
    "after": 72000,
    "staleCachedFormulaCount": 1,
    "suggestedReads": ["Summary!B2"]
  }
}
```

For a real workbook:

```sh
npm exec --package @bilig/xlsx-formula-recalc@latest -- xlsx-cache-doctor pricing.xlsx --json
```

For pull requests, start report-only:

```yaml
- uses: actions/setup-node@v6
  with:
    node-version: '22'
    package-manager-cache: false
- uses: proompteng/bilig@v1
  with:
    workbooks: '**/*.xlsx'
    changed-files-only: 'true'
    package-version: '0.131.3'
    fail-on-stale: 'false'
```

There is a demo pull request here:
<https://github.com/proompteng/xlsx-cache-doctor-demo/pull/1>.

## What it proves

- A workbook formula cell had a cached value before recalculation.
- Bilig recalculated the formula in Node.
- The report names the workbook, sheet, cell, cached value, fresh value, and
  suggested cells to read next.
- The GitHub Action can run in report-only mode before a project decides to fail
  pull requests.

## What it does not prove

- It is not a complete Excel clone.
- It does not run macros, pivots, charts, add-ins, or external data refreshes.
- It is not a hosted spreadsheet service.
- It should not be the only gate for irreversible money movement or regulated
  workflows until your formula set has fixtures and review.

## Where WorkPaper fits

Use the cache doctor when a saved `.xlsx` file is still the contract.

Use WorkPaper or `@bilig/headless` when your service or agent should own the
workbook model directly:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door workpaper-service --json
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

Those paths write inputs, recalculate formulas, read outputs, persist JSON, and
verify readback without opening Excel, LibreOffice, Google Sheets, or a browser
grid.

## Feedback I want

I am looking for hard workbook cases, not applause:

- SheetJS, ExcelJS, `xlsx-populate`, or template workflows that write inputs but
  read stale formula values later.
- CI repositories with committed `.xlsx` fixtures where stale cached results
  hide real changes.
- Reduced workbooks where the detector skips formulas it should inspect.
- Formulas that differ from Excel in a way that should become a public fixture.

Open a fixture issue:
<https://github.com/proompteng/bilig/issues/new?template=workbook_fixture.yml>

Project:
<https://github.com/proompteng/bilig>

## Share copy

Title:

```text
Show HN: XLSX Cache Doctor catches stale Excel formula values in CI
```

Short body:

```text
I built a narrow check for a boring spreadsheet bug: Node writes an XLSX input,
the formula cache stays stale, and the next service reads the old value.

One command runs locally and returns the formula cell, cached value, fresh value,
and suggested cells to read next:

npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- bilig-evaluate --door xlsx-cache --json

It does not claim full Excel parity. It is a report-first detector for stale
formula caches, with a GitHub Action that can run in report-only mode.

I am looking for reduced workbooks where this misses a stale value or disagrees
with Excel.
```
