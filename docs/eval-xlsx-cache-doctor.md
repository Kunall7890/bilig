---
title: Evaluate stale XLSX formula caches
published: true
description: Copy-paste evaluator for checking stale cached XLSX formula values in Node without Excel, LibreOffice, browser automation, or repo cloning.
tags: node, xlsx, formulas, spreadsheet, ci, evaluator
canonical_url: https://proompteng.github.io/bilig/eval-xlsx-cache-doctor.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Evaluate stale XLSX formula caches

Use this when the thing in your hands is an `.xlsx` file and you need to know
whether stored formula results are stale before a backend, CI job, or agent
trusts them.

This evaluator is read-only by default. It imports the workbook in Node,
recalculates formula cells in memory, compares those values with the cached
values stored in the file, and returns JSON that an agent or CI job can inspect.
It does not open Excel, LibreOffice, or a browser UI.

## One command

```sh
npm exec --package @bilig/xlsx-formula-recalc@latest -- xlsx-cache-doctor --demo --json
```

## Expected proof

The current demo prints this shape:

```json
{
  "schemaVersion": "xlsx-cache-doctor.v1",
  "mode": "demo",
  "input": "generated demo workbook",
  "formulaCellCount": 1,
  "inspectedFormulaCellCount": 1,
  "uninspectedFormulaCellCount": 0,
  "inspectionLimit": "all",
  "staleCachedFormulaCount": 1,
  "cacheStatusSummary": {
    "inspected": 1,
    "stale": 1,
    "fresh": 0,
    "missingCache": 0,
    "unsupportedRecalculation": 0
  },
  "suggestedReads": ["Summary!B2"],
  "formulas": [
    {
      "target": "Summary!B2",
      "formula": "=Inputs!B2*Inputs!B3",
      "cachedValue": 60000,
      "literalRecalculatedValue": 72000,
      "cacheStatus": "stale",
      "staleCachedValue": true
    }
  ],
  "commandSucceeded": true,
  "inspectionCompleted": true,
  "recalculationCompleted": true,
  "excelParity": "not_proven"
}
```

For the demo, the important checks are `commandSucceeded: true`,
`inspectionCompleted: true`, `recalculationCompleted: true`,
`uninspectedFormulaCellCount: 0`, and `staleCachedFormulaCount: 1`.
For a real workbook, treat `suggestedReads` as the next cells to prove with
`xlsx-recalc` after the detector tells you which cached results are stale.

`staleCachedValue: null` is intentionally split into explicit status buckets.
`cacheStatus: "missing-cache"` means the workbook did not store a cached value
for that formula. `cacheStatus: "unsupported-recalculation"` means Bilig did not
produce a comparable literal value for that cell. Neither should be mixed with
confirmed stale caches.

`excelParity: "not_proven"` is intentional. This check proves fresh Bilig
formula readback, not desktop Excel parity. Keep an Excel, LibreOffice, Graph,
or golden-file oracle for customer-critical workbook flows.

## Try your workbook

```sh
npm exec --package @bilig/xlsx-formula-recalc@latest -- xlsx-cache-doctor pricing.xlsx --json
```

If the workbook is large and you pass `--inspect-limit`, do not treat the report
as complete coverage unless `uninspectedFormulaCellCount` is `0`.

## Use the same report in Node

```ts
import { readFile } from "node:fs/promises";
import { inspectXlsxCache } from "@bilig/xlsx-formula-recalc";

const report = inspectXlsxCache(await readFile("pricing.xlsx"), {
  fileName: "pricing.xlsx",
});

if (report.staleCachedFormulaCount > 0) {
  throw new Error(
    report.formulas
      .filter((formula) => formula.cacheStatus === "stale")
      .map((formula) => formula.target)
      .join(", "),
  );
}
```

The API returns the same `schemaVersion`, `cacheStatusSummary`,
`staleCachedFormulaCount`, `suggestedReads`, and per-formula `cacheStatus` fields
as the CLI JSON.

## Turn the report into a fresh XLSX proof

After the detector returns `suggestedReads`, prove the actual cells your service
depends on:

```sh
npm exec --package @bilig/xlsx-formula-recalc@latest -- xlsx-recalc pricing.xlsx \
  --set Inputs!B2=48 \
  --set Inputs!B3=1500 \
  --read Summary!B2 \
  --out pricing.recalculated.xlsx \
  --json
```

Use sheet-qualified A1 references. Keep the adapter strict: known input cells,
known output cells, and tests around the exported workbook.

## Put it in CI

Use the repository action when stale cached formula values should be visible on
pull requests:

```yaml
- uses: proompteng/bilig@v1
  with:
    workbooks: "**/*.xlsx"
    changed-files-only: "true"
    fail-on-stale: "false"
```

The action writes JSON and Markdown reports, adds a job summary, and exposes
`formula-count`, `stale-count`, `uninspected-count`, `suggested-reads`, `json`,
and `markdown` outputs. Set `fail-on-stale: "true"` only after the team agrees
that stale workbook caches should block merges.

## What this proves

- the package can import the XLSX workbook in Node
- formula cells can be enumerated without a spreadsheet UI
- formula values can be recalculated in memory
- stale cached values are reported with exact cell addresses
- skipped formulas are visible through `uninspectedFormulaCellCount`
- the next `xlsx-recalc --read` targets are suggested explicitly

## What this does not prove

This is not a claim of complete Excel compatibility. It does not prove macros,
pivots, charts, unsupported formulas, locale-specific Excel behavior,
external-link freshness, or exact desktop Excel UI behavior.

## After the proof

- Star Bilig if this found a stale cached formula problem:
  <https://github.com/proompteng/bilig/stargazers>
- Watch releases if you want formula and XLSX compatibility updates:
  <https://github.com/proompteng/bilig/subscription>
- Report the exact adoption blocker:
  <https://github.com/proompteng/bilig/discussions/new?category=general>

## Related

- [Evaluate XLSX formula recalculation](eval-xlsx-recalc.md)
- [XLSX Cache Doctor GitHub Action](xlsx-cache-doctor-github-action.md)
- [Stale XLSX formula cache in Node.js](stale-xlsx-formula-cache-node.md)
- [Agent XLSX recalculation without LibreOffice](agent-xlsx-formula-recalculation-without-libreoffice.md)
- [Where Bilig is not Excel-compatible yet](where-bilig-is-not-excel-compatible-yet.md)
