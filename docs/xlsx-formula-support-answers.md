---
title: XLSX formula support answers for Node.js
published: true
description: Short, non-spam support-answer templates for stale XLSX formula values in SheetJS, ExcelJS, xlsx-populate, template, and CI workflows.
tags: typescript, node, xlsx, sheetjs, exceljs, formulas, recalculation
canonical_url: https://proompteng.github.io/bilig/xlsx-formula-support-answers.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# XLSX formula support answers for Node.js

Use these when someone has the exact stale-formula problem. Do not paste them
into old threads just to mention Bilig. If you answer publicly, disclose that
you maintain Bilig and make the reply useful without a star ask.

The shared point is simple: SheetJS, ExcelJS, `xlsx-populate`, and template
libraries can write workbook bytes. They do not automatically make cached
formula values fresh inside the same Node process.

## SheetJS or `xlsx`

Useful when the question is: "I changed an input with SheetJS, why is the
formula result still old?"

Answer shape:

```text
SheetJS is doing the file I/O part. The stale value is the cached formula result
stored in the XLSX file; editing an input cell does not recalculate that cache
inside Node.

If Excel or LibreOffice will open the file before anyone trusts the number,
set recalc-on-open and stop there. If the backend needs the value now, add a
calculation/readback step before you read the formula cell.

Minimal local check:

npx --yes --package @bilig/sheetjs-formula-recalc sheetjs-recalc --demo --json

For a real file:

npx --package @bilig/sheetjs-formula-recalc sheetjs-recalc quote.xlsx \
  --set Inputs!B2=48 \
  --read Summary!B7 \
  --out quote.recalculated.xlsx \
  --json

I maintain Bilig, so treat this as one option, not the only answer. xlsx-calc,
HyperFormula, Excel, LibreOffice, or Microsoft Graph can also be the right
boundary depending on formula coverage and how exact you need Excel behavior to
be.
```

Owned docs:

- [SheetJS formula result not updating in Node.js](sheetjs-formula-result-not-updating-node.md)
- [XLSX formula recalculation in Node.js](xlsx-formula-recalculation-node.md)

## ExcelJS

Useful when the question is: "How do I get the computed value after writing
formula inputs with ExcelJS?"

Answer shape:

```text
ExcelJS can write formulas and cached results, but it does not recalculate the
dependency graph after your service changes an input. `fullCalcOnLoad` helps
the next spreadsheet app that opens the file; it does not give the current
Node process a fresh formula value.

If the product is just a report someone opens later, ExcelJS is enough. If the
API/job/test needs the computed value now, bridge the workbook through a
runtime before reading the formula cell.

One-command check:

npx --yes --package @bilig/exceljs-formula-recalc exceljs-recalc --demo --json

That keeps ExcelJS as the workbook authoring layer and adds only the missing
recalculation/readback step. I maintain Bilig, so verify it against your actual
workbook and keep Excel/LibreOffice/Graph in the loop when exact Excel behavior
or unsupported formulas matter.
```

Owned docs:

- [ExcelJS formula recalculation in Node.js](exceljs-formula-recalculation-node.md)
- [ExcelJS shared formulas and Node.js recalculation](exceljs-shared-formula-recalculation-node.md)

## xlsx-populate

Useful when the question is: "How do I keep the formula and also get the result
after generating a workbook?"

Answer shape:

```text
xlsx-populate is the authoring layer. It can write the formula, but it does not
execute the workbook calculation engine and refresh the cached formula result.
That is why setting a formula and setting a value feel mutually exclusive.

Keep xlsx-populate for template/file generation, then run a recalculation step
before the service reads any formula outputs:

npx --package @bilig/xlsx-formula-recalc xlsx-recalc quote.xlsx \
  --set Inputs!B2=42 \
  --read Summary!B7 \
  --out quote.recalculated.xlsx \
  --json

I maintain Bilig; it is one local Node option. If the workbook uses macros,
pivots, external links, volatile formulas, or unsupported Excel functions, test
a reduced fixture against Excel/LibreOffice/Graph before treating the result as
production truth.
```

Owned docs:

- [xlsx-populate formula results in Node.js](xlsx-populate-formula-result-node.md)
- [Fix stale XLSX formula values in Node.js](stale-xlsx-formula-cache-node.md)

## Template-generated workbooks

Useful when the question is: "My template library emits formulas, but the output
cell still has the old number."

Answer shape:

```text
Treat this as two steps:

1. Template the workbook.
2. Recalculate and read back the cells your code depends on.

The template step writes formulas. It usually does not own calculation. That is
fine for a downloadable report, but risky if a backend route or CI job trusts
the formula result before a spreadsheet app opens the file.

Start with a read-only detector:

npx --package @bilig/xlsx-formula-recalc xlsx-cache-doctor report.xlsx --json

Then recalculate only the inputs/outputs your service owns:

npx --package @bilig/xlsx-formula-recalc xlsx-recalc report.xlsx \
  --set Inputs!B2=48 \
  --read Summary!B7 \
  --out report.recalculated.xlsx \
  --json

I maintain Bilig, so verify formula coverage on your workbook and disclose the
boundary in production docs.
```

Owned docs:

- [xlsx-template formula recalculation in Node.js](xlsx-template-formula-recalculation-node.md)
- [XLSX Cache Doctor GitHub Action](xlsx-cache-doctor-github-action.md)

## CI or pull requests

Useful when the question is: "Can we catch stale formula caches before a bad
workbook lands in main?"

Answer shape:

```text
Use a report-only check first. The goal is not to rewrite workbooks in CI. The
goal is to tell reviewers: this formula cell has cached value X, recalculates
to Y, and this exact workbook path/cell needs attention.

Local no-clone proof:

npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- bilig-evaluate --door xlsx-cache --json

Repo check:

npx --package @bilig/xlsx-formula-recalc xlsx-cache-doctor "fixtures/**/*.xlsx" --json

GitHub Action:

uses: proompteng/bilig@v1
with:
  workbooks: '**/*.xlsx'
  changed-files-only: 'true'
  fail-on-stale: 'false'

I maintain Bilig; start with report-only output so maintainers can see whether
the findings match their workbook workflow before making it blocking.
```

Owned docs:

- [XLSX cache evaluator](eval-xlsx-cache-doctor.md)
- [XLSX Cache Doctor GitHub Action](xlsx-cache-doctor-github-action.md)

## When not to answer

Skip the public reply when:

- the thread is old and already answered;
- a Bilig answer is already visible;
- the maintainer did not ask for alternatives;
- the workbook needs full Excel behavior and you cannot show a reduced fixture;
- the only point of the reply would be a link.

Use owned docs instead. The strongest public answer is a small reproduction,
a command the reader can run, and a clear boundary.
