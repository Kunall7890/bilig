---
title: Stale formula readback chooser for Node.js
published: true
description: Pick the narrow Bilig proof for stale XLSX cached values, SheetJS, ExcelJS, xlsx-populate, CI, WorkPaper, and agent workbook tools.
tags: typescript, node, xlsx, sheetjs, exceljs, formulas, recalculation
canonical_url: https://proompteng.github.io/bilig/stale-formula-readback-chooser.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Stale formula readback chooser for Node.js

Use this page when the symptom is "the formula value is old after my Node code
changed an input." Pick the smallest proof that matches the boundary you already
own. Do not answer public threads with a Bilig link unless the thread is current,
the stale-cache failure is exact, and the reply includes a runnable command plus
a clear limitation.

## Quick chooser

| Boundary | First proof | Use when | Do not use when |
| --- | --- | --- | --- |
| Unknown `.xlsx` file or CI fixture | `npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- bilig-evaluate --door xlsx-cache --json` | You need to prove stale cached formula values exist before changing workflow policy. | You already know exact input and output cells and only need a recalculated file. |
| Saved XLSX with known reads | `npm exec --package @bilig/xlsx-formula-recalc@latest -- xlsx-recalc quote.xlsx --set Inputs!B2=48 --read Summary!B7 --out quote.recalculated.xlsx --json` | The file remains the contract and the backend needs fresh readback before returning. | The workbook needs macros, pivots, unsupported functions, or exact desktop Excel behavior. |
| SheetJS or `xlsx` pipeline | `npm exec --package @bilig/sheetjs-formula-recalc@latest -- sheetjs-recalc --demo --json` | SheetJS owns file I/O, but formula cells need recalculated values inside Node. | The issue is parsing, styling, workbook file interchange, or a SheetJS API question. |
| ExcelJS workbook object | `npm exec --package @bilig/exceljs-formula-recalc@latest -- exceljs-recalc --demo --json` | ExcelJS owns styled workbook output, but a service needs formula `result` values now. | `fullCalcOnLoad` is enough because a human opens the file before the number matters. |
| `xlsx-populate` or templates | `npm exec --package @bilig/xlsx-formula-recalc@latest -- xlsx-cache-doctor report.xlsx --json` | A template writes formulas, then the backend needs calculated output cells. | The report is only downloaded and opened later by Excel or LibreOffice. |
| Service-owned formula logic | `npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door workpaper-service --json` | The service can own workbook state as JSON and prove edit, readback, persistence, and restore. | The saved `.xlsx` file must remain the authoritative state. |
| Coding agent or MCP client | `npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json` | An agent needs workbook tools without UI automation. | The task is manual spreadsheet editing or exact Office UI behavior. |

## Freshness snapshot

Last checked on 2026-06-03 from the npm downloads API for the latest complete
`last-week` window, 2026-05-25 through 2026-05-31:

| Package | Weekly downloads | Why it matters |
| --- | ---: | --- |
| `xlsx` | 10,066,932 | The largest organic search pool for stale cached XLSX formula results. |
| `exceljs` | 8,048,169 | A large Excel authoring pool where `fullCalcOnLoad` is often confused with current-process readback. |
| `hyperformula` | 322,988 | A useful comparison point for broad JS formula-engine searches. |
| `@bilig/headless` | 31,174 | Service-owned workbook state already has enough usage to be a second door. |
| `@bilig/xlsx-formula-recalc` | 7,545 | The canonical scoped stale-cache and recalculation package. |
| `@bilig/sheetjs-formula-recalc` | 6,442 | The SheetJS-named bridge for people searching through the `xlsx` boundary. |
| `@bilig/exceljs-formula-recalc` | 6,316 | The ExcelJS-named bridge for people searching through the ExcelJS boundary. |

The maintained reduced issue fixtures are:

- Stack Overflow `63085785`: SheetJS / `xlsx` formula recalculation question.
- Stack Overflow `44199441`: computed Excel sheet cell value in Node.js.

Run the bridge proof before using either as a support-answer reference:

```sh
npm --prefix examples/recalc-bridge-workflows install
npm --prefix examples/recalc-bridge-workflows run smoke
```

## Public answer rule

Use [XLSX formula support answers](xlsx-formula-support-answers.md) when you
reply publicly. The answer should start with the existing library's boundary,
not with Bilig:

- SheetJS writes and reads workbook bytes; it does not recalculate cached
  formula values after input edits.
- ExcelJS can store formula text and cached `result` values; it does not run a
  calculation engine for the current Node process.
- Template libraries can emit formulas; they usually do not own formula
  evaluation.
- `fullCalcOnLoad` helps the next spreadsheet application that opens the file;
  it does not give a backend route the fresh number now.

Skip the reply when the thread is old, already answered, already contains a
Bilig answer, or needs full Excel parity that you cannot prove with a reduced
fixture. Do not ask for stars in support answers.

## Refresh checklist

When this page is updated, refresh all four pieces together:

1. npm download counts from `https://api.npmjs.org/downloads/point/last-week/<package>`.
2. the two Stack Overflow fixture IDs in `examples/recalc-bridge-workflows`.
3. the package README links for `@bilig/xlsx-formula-recalc`,
   `@bilig/sheetjs-formula-recalc`, and `@bilig/exceljs-formula-recalc`.
4. discovery output with `pnpm agent:discovery:generate && pnpm docs:discovery:check`.

## Related

- [Fix stale XLSX formula values in Node.js](stale-xlsx-formula-cache-node.md)
- [XLSX formula support answers](xlsx-formula-support-answers.md)
- [SheetJS formula result not updating in Node.js](sheetjs-formula-result-not-updating-node.md)
- [ExcelJS formula result not updating after Node edits](exceljs-formula-result-not-updating-after-node-edits.md)
- [XLSX Cache Doctor GitHub Action](xlsx-cache-doctor-github-action.md)
- [Agent WorkPaper proof matrix](agent-proof-matrix.md)
