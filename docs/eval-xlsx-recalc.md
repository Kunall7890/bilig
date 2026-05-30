---
title: Evaluate XLSX formula recalculation
published: true
description: Copy-paste evaluator for teams with an XLSX file that needs fresh formula values in Node without Excel, LibreOffice, or browser automation.
tags: node, xlsx, formulas, spreadsheet, evaluator
canonical_url: https://proompteng.github.io/bilig/eval-xlsx-recalc.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Evaluate XLSX formula recalculation

Use this when the thing in your hands is an `.xlsx` file. The narrow question is
whether Node can edit known input cells, recalculate formulas, write a new XLSX,
and return a proof object without opening Excel, LibreOffice, or a browser UI.

## One command

```sh
npm exec --package @bilig/xlsx-formula-recalc@latest -- xlsx-recalc --demo --json
```

## Expected proof

The current demo prints this shape:

```json
{
  "mode": "demo",
  "input": "generated demo workbook",
  "output": "bilig-formula-recalc-demo.xlsx",
  "edits": 2,
  "externalWorkbooks": 0,
  "reads": {
    "Summary!B2": {
      "tag": 1,
      "value": 72000
    }
  },
  "warnings": [],
  "commandSucceeded": true,
  "recalculationCompleted": true,
  "excelParity": "not_proven",
  "expectedReadback": {
    "Summary!B2": 72000
  },
  "expectedValueMatched": true
}
```

The exact output file name can change if you pass your own `--out` path. The
important checks are `commandSucceeded: true`, `recalculationCompleted: true`,
an empty or understood `warnings` array, and the recalculated cell value under
`reads`. `expectedValueMatched: true` is only a demo-fixture check. It is not an
Excel parity claim; real workbooks still report `excelParity: "not_proven"`
unless you compare against your own Excel, LibreOffice, or Graph oracle.

The JSON is clean for CI and agents. Star, release-watch, and
adoption-blocker links stay in prose after the proof so the machine output does
not look like a growth prompt.

## Inspect your workbook first

If you already have the workbook but do not know the right output cells yet,
start with inspection:

```sh
npm exec --package @bilig/xlsx-formula-recalc@latest -- xlsx-cache-doctor pricing.xlsx --json
```

That command does not write `pricing.recalculated.xlsx`. It imports the
workbook, lists formula cells, recomputes every formula by default, reports
stale cached values, and suggests `--read` targets for the real proof command.
If you intentionally pass `--inspect-limit`, require
`uninspectedFormulaCellCount: 0` before treating the report as complete
coverage.

Expected shape:

```json
{
  "formulaCellCount": 12,
  "inspectedFormulaCellCount": 12,
  "uninspectedFormulaCellCount": 0,
  "inspectionLimit": "all",
  "staleCachedFormulaCount": 3,
  "suggestedReads": ["Summary!B7"],
  "formulas": [
    {
      "target": "Summary!B7",
      "formula": "=Inputs!B2*Inputs!B3",
      "cachedValue": 60000,
      "literalRecalculatedValue": 72000,
      "staleCachedValue": true
    }
  ],
  "commandSucceeded": true,
  "inspectionCompleted": true,
  "recalculationCompleted": true,
  "excelParity": "not_proven"
}
```

## Try your workbook

```sh
npm exec --package @bilig/xlsx-formula-recalc@latest -- xlsx-recalc pricing.xlsx \
  --set Inputs!B2=48 \
  --set Inputs!B3=1500 \
  --read Summary!B7 \
  --out pricing.recalculated.xlsx \
  --json
```

Use sheet-qualified A1 references. Keep your adapter strict: known input cells,
known output cells, and tests around the exported workbook.

## Put it in CI

Use the repository action when stale cached formula values should block a pull
request:

```yaml
- uses: proompteng/bilig@v1
  with:
    workbook: fixtures/pricing.xlsx
    fail-on-stale: "true"
```

The action writes a JSON report, adds a job summary, and exposes
`formula-count`, `stale-count`, `uninspected-count`, and `suggested-reads`
outputs. See
[XLSX Cache Doctor GitHub Action](xlsx-cache-doctor-github-action.md).

## What this proves

- the package can import an XLSX workbook in Node
- known input cells can be edited from a command
- dependent formulas can be recalculated and read back
- the edited workbook can be written back to XLSX bytes
- warnings are visible instead of hidden behind a "success" message

## What this does not prove

This is not a claim of complete Excel parity. It does not prove macros, pivots,
charts, unsupported formulas, locale-specific Excel behavior, external-link
freshness, or exact desktop Excel UI behavior. Keep a golden workbook fixture
and an Excel or LibreOffice oracle test for customer-critical file flows.

## After the proof

- Star Bilig if this solved the XLSX recalculation problem:
  <https://github.com/proompteng/bilig/stargazers>
- Watch releases if you want compatibility and formula updates:
  <https://github.com/proompteng/bilig/subscription>
- Report the exact adoption blocker:
  <https://github.com/proompteng/bilig/discussions/new?category=general>

## Related

- [XLSX formula recalculation in Node.js](xlsx-formula-recalculation-node.md)
- [Curlable XLSX recalculation proof](xlsx-recalculation-proof.md)
- [External workbook recalculation proof](external-workbook-recalc-proof.md)
- [Agent XLSX recalculation without LibreOffice](agent-xlsx-formula-recalculation-without-libreoffice.md)
- [Where Bilig is not Excel-compatible yet](where-bilig-is-not-excel-compatible-yet.md)
