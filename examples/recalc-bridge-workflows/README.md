# Recalc bridge workflows

This example is for developers already using SheetJS/`xlsx`, `xlsx-populate`,
or ExcelJS who hit stale formula values after editing workbook inputs in Node.

It proves the boundary in one command:

1. create a formula-backed `.xlsx` pricing workbook;
2. edit the input cells through SheetJS/`xlsx`;
3. edit the same input cells through `xlsx-populate`;
4. edit the same input cells through ExcelJS;
5. show the stale cached formula value each library still sees;
6. run Bilig recalculation and verify the fresh formula value.

## Run

```sh
npm install
npm run smoke
```

Expected output includes:

```json
{
  "verified": true
}
```

## Why this exists

SheetJS, `xlsx-populate`, and ExcelJS are useful file/workbook libraries. They
are not in-process Excel calculation engines. If your service changes
`Inputs!B2` and `Inputs!B3`, a dependent formula such as `Summary!B2` can still
show the old cached value until another calculation step runs.

Use `xlsx-formula-recalc` when you have XLSX bytes from SheetJS or
`xlsx-populate`. Use `exceljs-formula-recalc` when you need the recalculated
values patched back onto an ExcelJS workbook object.
