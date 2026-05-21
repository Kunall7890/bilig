# @bilig/excel-fixtures

Checked-in Excel compatibility fixtures and parity inputs for bilig.

## Install

```bash
npm install @bilig/excel-fixtures
```

## Package entrypoints

- ESM: `./dist/index.js`
- Types: `./dist/index.d.ts`

## macOS Desktop Excel Oracle

This package owns the reusable local Excel oracle surface for compatibility tests. It can build and run an `osascript` harness that opens a workbook in Microsoft Excel for Mac, writes formulas, forces `calculate full rebuild`, reads typed cell values, and closes without saving.

Default package tests do not launch Excel. To run the live oracle test on a Mac with Excel installed:

```bash
BILIG_EXCEL_ORACLE_RUN=1 pnpm --filter @bilig/excel-fixtures test
```

This package is part of the [bilig](https://github.com/proompteng/bilig) monorepo.
