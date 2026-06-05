---
title: Bilig formula bug clinic
published: true
description: Send a reduced public XLSX, ExcelJS shared-formula, stale cached formula, or WorkPaper blocker so it can become a Bilig fixture, test, docs page, or example.
tags: typescript, node, exceljs, xlsx, formulas, open source
canonical_url: https://proompteng.github.io/bilig/formula-bug-clinic.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Bilig formula bug clinic

If a workbook formula bug is blocking your Node service, send the smallest
public case that proves it. The goal is not to collect private spreadsheets.
The goal is to turn real failures into public fixtures that future evaluators
can run.

Start with the cache check when the symptom is "Node changed cells, but the
formula output stayed old":

```sh
npm exec --package @bilig/xlsx-formula-recalc@latest -- \
  xlsx-cache-doctor ./reduced.xlsx --json
```

That command runs locally, does not upload the workbook, and reports exact
stale-cache evidence: `target`, `formula`, `cachedValue`,
`literalRecalculatedValue`, `cacheStatus`, and `suggestedReads`. If it finds a
stale cell, paste the relevant JSON object into the fixture form or discussion
instead of describing the bug in prose.

Good cases:

- an ExcelJS, SheetJS, or `xlsx-populate` pipeline writes inputs but keeps a
  stale cached formula value;
- an ExcelJS workflow writes inputs but formula readback is stale;
- an XLSX uses shared formulas and the imported formula text is wrong;
- a workbook works in Excel but fails in a local Node formula runtime;
- a WorkPaper JSON restore changes a calculated value;
- an agent or MCP tool writes a cell but cannot prove the recalculated output;
- a service route needs one missing formula family, import detail, or example.

Open the fixture form when the reduced public fixture is ready:
<https://github.com/proompteng/bilig/issues/new?template=workbook_fixture.yml>.

Discuss the shape first if you are still reducing the case:
<https://github.com/proompteng/bilig/discussions/414>.

## Generate a local report

Use the narrowest command that matches the blocker:

| Blocker | First local command | What to paste |
| --- | --- | --- |
| Stale cached XLSX formula value after Node edits | `npm exec --package @bilig/xlsx-formula-recalc@latest -- xlsx-cache-doctor ./reduced.xlsx --json` | The stale formula object with cell address, cached value, recalculated value, and `suggestedReads`. |
| Pull requests can commit stale workbook fixtures | `npm exec --package @bilig/xlsx-formula-recalc@latest -- xlsx-cache-doctor --print-github-action "**/*.xlsx"` | The generated report-only workflow, or the GitHub Action report artifact. |
| WorkPaper import, formula, or persistence mismatch | `npm exec --package @bilig/headless@0.163.0 -- bilig-formula-clinic ./reduced.xlsx --cells "Summary!B7,Inputs!B2"` | The Markdown clinic report with requested cells, formula samples, warnings, and actual readback. |

If the workbook is already reduced and the stale-cache check is not enough, run
the clinic reporter locally and paste the Markdown output into the fixture form.
It reads the file on your machine and does not upload workbook contents.

```sh
npm exec --package @bilig/headless@0.163.0 -- bilig-formula-clinic ./reduced.xlsx \
  --cells "Summary!B7,Inputs!B2"
```

That is the lowest-friction path for package users. It imports the workbook,
samples formulas, reads the requested cells through WorkPaper, and prints a
Markdown report.

If you want to pin or edit the reporter script directly:

```sh
mkdir bilig-formula-clinic
cd bilig-formula-clinic
npm init -y
npm pkg set type=module
npm install @bilig/headless
npm install --save-dev tsx typescript @types/node
curl -fsSLo formula-clinic-report.ts \
  https://proompteng.github.io/bilig/formula-clinic-report.ts
npx tsx formula-clinic-report.ts ./reduced.xlsx \
  --cells "Summary!B7,Inputs!B2"
```

Use `--cells` for the output cells that prove the bug. The report includes
import warnings, formula samples, requested readback, and a fixture checklist.

## What to send

Send one reduced public fixture, not the whole production workbook.

Include:

- package version or commit tested;
- sheet names and exact cells or ranges;
- formulas involved;
- input values before and after the edit;
- expected output from Excel, LibreOffice, Graph, an existing service, or a
  manual check;
- actual Bilig output, import error, stale cached value, or missing API;
- for stale-cache cases: the exact `target`, `formula`, `cachedValue`,
  `literalRecalculatedValue`, and `cacheStatus` from `xlsx-cache-doctor`;
- the shortest command or script that maintainers can run.

Do not attach confidential workbooks, customer data, financial models, or files
that cannot be redistributed in a public test corpus. Replace names and numbers
with neutral values while keeping the same formula shape.

Good discussion summary:

```text
Reduced public workbook attached or linked. xlsx-cache-doctor reports stale
Sheet1!B61: formula =A61*10, cached 999, recalculated 600. The service reads
Sheet1!B61 after changing Sheet1!A61. Expected output comes from Excel save.
```

Bad discussion summary:

```text
My spreadsheet is wrong. Can Bilig support it?
```

## Why this helps

Stars usually follow evidence, not claims. A reduced workbook fixture is better
than a marketing post because it gives maintainers something concrete to merge:

- a regression test;
- an XLSX import/export corpus case;
- a formula compatibility note;
- a WorkPaper JSON persistence fixture;
- a service-route example;
- an MCP or agent-tool transcript.

When a case lands, the issue can point to the commit, release, and docs page
that fixed it. That is the evidence a skeptical backend developer can inspect
before adopting the package.

## Fast local check

For stale cached XLSX values, first verify whether the backend is reading an old
stored value instead of a fresh calculation:

```sh
npm exec --package @bilig/xlsx-formula-recalc@latest -- \
  xlsx-cache-doctor ./reduced.xlsx --json
```

If you need to see the exact output shape first, use the committed transcript:
[XLSX Cache Doctor proof transcript](xlsx-cache-doctor-proof-transcript.md).
If the check belongs in pull requests, start from the report-only
[XLSX Cache Doctor GitHub Action](xlsx-cache-doctor-github-action.md).

For a pure WorkPaper case, reduce it to a script:

```sh
mkdir bilig-fixture-check
cd bilig-fixture-check
npm init -y
npm pkg set type=module
npm install @bilig/headless
npm install --save-dev tsx typescript @types/node
```

If the script is short enough to paste into an issue, it is probably a good
fixture.

## Useful references

- [Submit a workbook fixture](submit-workbook-fixture.md)
- [ExcelJS shared formulas and Node.js recalculation](exceljs-shared-formula-recalculation-node.md)
- [Fix stale XLSX formula values in Node.js](stale-xlsx-formula-cache-node.md)
- [XLSX formula recalculation in Node.js](xlsx-formula-recalculation-node.md)
- [Where Bilig is not Excel-compatible yet](where-bilig-is-not-excel-compatible-yet.md)

If this helped you reduce a workbook bug but a gap still blocks adoption,
open one concrete blocker or fixture note:
<https://github.com/proompteng/bilig/discussions/new?category=general>.
