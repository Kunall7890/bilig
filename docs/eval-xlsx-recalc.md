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
  "verified": true
}
```

The exact output file name can change if you pass your own `--out` path. The
important checks are `verified: true`, an empty or understood `warnings` array,
and the recalculated cell value under `reads`.

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
- [Agent XLSX recalculation without LibreOffice](agent-xlsx-formula-recalculation-without-libreoffice.md)
- [Where Bilig is not Excel-compatible yet](where-bilig-is-not-excel-compatible-yet.md)
