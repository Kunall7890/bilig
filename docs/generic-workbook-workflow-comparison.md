# Generic workbook workflow comparison

The goal is domain-neutral workbook workflow support. Bilig must import and operate on ordinary `.xlsx` workbooks without branching on a workbook name, sheet name, vendor label, business process, or one reference file.

## Goal

Bilig should improve the normal workbook flow:

- import arbitrary `.xlsx` files through the same UI path
- preserve sheet order, formulas, literal values, date serials, number formats, column widths, row heights, merged ranges, and mappable cell styles
- keep formulas visible in the formula bar after import
- keep formatting, border, merge, hide, unhide, freeze, autofit, undo, and redo operations available after import
- verify behavior with multiple generated workbook shapes plus an optional local external workbook smoke test

## Verified Generic Fixtures

The browser import path uses generated fixtures that are intentionally not tied to one business domain:

- `multi-sheet-operations`: dashboard, ledger, rollforward, and lookup sheets with cross-sheet formulas, dimensions, and merged title rows
- `single-sheet-planning`: one-sheet monthly planning schedule with date serials, month formulas, totals, dimensions, and a merged title row
- optional `BILIG_REFERENCE_WORKBOOK_XLSX`: any local workbook path supplied by the operator; expected sheets and sample cells are derived from the workbook contents instead of hardcoded names

The import unit tests cover the same generic behavior below the UI layer and assert that no file-specific dispatch is required.

## Current Generic Advantages

- Imported formulas remain editable and inspectable as formulas.
- Imported values keep primitive workbook types instead of being flattened into display strings.
- Column widths, row heights, and merged ranges are represented in workbook metadata.
- Mappable fills, font settings, alignment, and borders are converted into Bilig style records.
- Unsupported workbook features produce explicit warnings instead of pretending to import perfectly.

## Known Gaps

- Comments and defined names are still ignored.
- Indexed and theme colors are only imported when the library exposes enough direct color data.
- The generated fixtures do not prove every real-world workbook shape.
- This does not claim Google Sheets collaboration parity.
- There is no domain-specific built-in template in this goal.

## Completion Bar

This goal is complete only when all of these are true:

- generic import unit coverage passes for multiple workbook shapes
- browser import coverage passes through the visible UI for multiple generated workbook shapes
- the optional external verifier can validate any supplied local `.xlsx` without encoding that workbook's domain
- the toolbar does not expose a domain-specific hardcoded template
- `pnpm run ci` is green on the final committed tree
