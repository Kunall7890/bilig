# @bilig/excel-import

CSV/XLSX-to-`WorkbookSnapshot` import helpers and supported-subset XLSX export helpers for bilig.

## Install

```sh
pnpm add @bilig/excel-import @bilig/headless
```

## XLSX To WorkPaper

```ts
import { readFileSync } from "node:fs";
import { WorkPaper } from "@bilig/headless";
import { importXlsx } from "@bilig/excel-import";

const imported = importXlsx(
  new Uint8Array(readFileSync("model.xlsx")),
  "model.xlsx",
);

const workbook = WorkPaper.buildFromSnapshot(imported.snapshot, {
  evaluationTimeoutMs: 30_000,
  useColumnIndex: true,
});
```

Use `WorkPaper.buildFromSnapshot()` for imported XLSX files. It preserves the
workbook metadata that Excel formulas need, including defined names, table
metadata, and structured-reference translations. `WorkPaper.buildFromSheets()`
is intentionally metadata-free.
