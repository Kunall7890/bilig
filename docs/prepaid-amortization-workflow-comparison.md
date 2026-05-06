# Prepaid amortization workflow comparison

This is the realistic bar for making Bilig better than a single referenced prepaid-amortization spreadsheet. The goal is not to reproduce one workbook by hardcoding its title, vendor list, formulas, colors, or cell addresses. The goal is to make the normal Bilig workflow better for the general class of prepaid schedules.

## Goal

Bilig should handle prepaid-amortization workbooks through reusable product behavior:

- import ordinary `.xlsx` files through the workbook import path
- preserve workbook structure and visible metadata where the runtime model supports it
- preserve formulas, literal values, dates as spreadsheet serials, number formats, widths, row heights, merges, and mappable cell styling
- make prepaid schedule setup easier through a reusable template flow, not a file-specific clone
- expose formula and error state clearly enough that users can audit monthly amortization, remaining balances, and edge cases

## Current reference workbook

The referenced prepaid schedule is a normal workbook-style schedule, not a special protocol:

- one visible sheet named `Prepaid Schedule`
- used range `A1:U100`
- title and header rows followed by prepaid item rows
- formulas for month counts, monthly amortization, month-by-month expense, total amortization, and remaining balance
- visible column sizing and header fill styling

Those facts are useful as a comparison fixture, but they must not be used as dispatch conditions in product code.

## Bilig behavior after the import-fidelity slice

The normal XLSX importer preserves:

- sheet names and sheet order
- formulas from workbook cells
- literal string, number, boolean, and date-serial values
- cell number-format codes
- column widths from `!cols`
- row heights from `!rows`
- merge ranges from `!merges`
- mappable cell styles exposed by SheetJS, currently solid RGB fills, font family, font size, bold, italic, underline, font color, horizontal and vertical alignment, wrapping, indent, and common border styles
- prepaid template generation from parameters, including schedule year, custom items, and row count
- prepaid template status formulas that surface blank rows, missing dates, missing amount, reversed dates, over-amortized balances, complete rows, not-started rows, and in-progress rows

This is generic import behavior. It is not scoped to prepaid files and does not inspect workbook names, sheet names, vendor names, or the reference workbook's formula layout.

## Verified fixtures

The browser workflow now covers two generated prepaid-amortization workbook shapes through the normal import UI:

- `reference-style-prepaid`: a multi-sheet dashboard/tracking/amortization/categories workbook with cross-sheet formulas, widths, row heights, and merged title rows
- `single-sheet-daily-prepaid`: a one-sheet daily-prorated amortization workbook with month formulas, totals, remaining-balance formulas, widths, row heights, and a merged title row

The external referenced workbook can be verified without committing or product-hardcoding the local file by setting:

```bash
BILIG_REFERENCE_PREPAID_XLSX=/Users/gregkonush/Downloads/Prepaid\ Expense\ Template.xlsx pnpm test:browser -- e2e/tests/web-shell-import.pw.ts --workers=1 --grep "external referenced prepaid workbook"
```

## Bilig advantages to prove

For prepaid amortization specifically, Bilig should beat a static spreadsheet when it can demonstrate:

- faster setup for a new schedule with a reusable template flow
- visible formula preservation after import instead of opaque pasted values
- better error detection for invalid dates, zero-month schedules, missing costs, and bad remaining-balance formulas
- clearer row-level status formulas for missing dates, missing amount, reversed date ranges, and over-amortized balances
- safer editing through normal workbook undo and structured operations
- reusable formatting and sizing controls that survive import/export and snapshot restore
- browser tests proving the flow works in the actual UI, not just in package-level unit tests

## Known gaps

These are not complete yet:

- imported comments are still ignored
- imported defined names are still ignored
- theme and indexed colors are not fully resolved when SheetJS does not expose RGB colors
- live Google Sheets collaboration and Sheets-specific formula behavior are not parity targets yet
- the toolbar currently applies the default prepaid template preset; it does not yet expose template parameters in the UI
- the prepaid template generator covers custom years, items, and row counts, but not multiple accounting conventions yet
- generated browser fixtures cover multiple prepaid-amortization imports, but a committed reference workbook fixture is intentionally not added; the local reference workbook is covered by the optional `BILIG_REFERENCE_PREPAID_XLSX` verifier above

## Completion bar

The prepaid-amortization goal is only complete when all of these are true:

- two or more prepaid-amortization XLSX fixtures import through the normal UI path: covered by `pnpm test:browser -- e2e/tests/web-shell-import.pw.ts --workers=1`
- the local referenced workbook imports through the normal UI path: covered by `BILIG_REFERENCE_PREPAID_XLSX=/Users/gregkonush/Downloads/Prepaid\ Expense\ Template.xlsx pnpm test:browser -- e2e/tests/web-shell-import.pw.ts --workers=1 --grep "external referenced prepaid workbook"`
- formulas, values, dimensions, merges, and mappable formatting survive import: covered by `pnpm exec vitest run packages/excel-import/src/__tests__/excel-import.test.ts` plus the browser import tests for formulas, values, and widths
- the reusable template path is parameterized and covered by tests: covered by `pnpm exec vitest run apps/web/src/__tests__/workbook-prepaid-template.test.ts`
- formula and error visibility is covered by focused tests and a real browser flow: covered by `pnpm test:browser -- e2e/tests/web-shell-toolbar.pw.ts --workers=1 --grep "prepaid amortization template|prepaid template row validation"`
- this comparison document is updated with exact passing commands and any remaining unsupported gaps
