---
title: Google Sheets QUERY and SORTN formulas in Node.js
published: true
description: Run Google Sheets-style QUERY and SORTN formulas over local WorkPaper ranges in Node.js, with verified readback and a clear boundary around provider-backed imports.
tags: google sheets, query, sortn, node, formulas, workpaper
canonical_url: https://proompteng.github.io/bilig/google-sheets-query-sortn-node-workpaper.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Google Sheets QUERY and SORTN formulas in Node.js

Use `@bilig/workpaper` when a Node service or coding agent needs Google
Sheets-style formula behavior over workbook state it already owns.

This is not a Google Sheets connector. It does not fetch a live spreadsheet,
read Drive permissions, or run a remote Visualization API query. It evaluates
local WorkPaper ranges, recalculates formulas, reads the result back, and can
persist the WorkPaper document as JSON.

## What works locally

Bilig supports the useful service-side subset:

- `QUERY(range, "select ... where ... order by ... limit ... offset ...", headers)`
- `QUERY` `group by` with `sum(column)` and `count(column)`
- `QUERY` `label` for selected columns and supported aggregate output headers
- `SORTN(range, n, tie_mode, sort_column_or_range, ascending, ...)`
- `COUNTUNIQUEIFS(...)`
- `ARRAYFORMULA(...)` spill evaluation

That covers the common backend job: take a small model, group or filter it,
sort the rows, read the calculated output, and save the state.

Unsupported `QUERY` clauses fail closed. Do not expect `pivot`, `having`,
`format`, `options`, arbitrary SQL, or live Google data fetching.

Provider-backed imports such as `IMPORTDATA`, `IMPORTRANGE`, `IMPORTHTML`,
`IMPORTXML`, `IMPORTFEED`, and `GOOGLEFINANCE` are a separate boundary. Without
a host adapter, they return a blocked result instead of pretending to have
network or account access.

## Run the proof

From an empty directory:

```sh
mkdir bilig-query-sortn
cd bilig-query-sortn
npm init -y
npm pkg set type=module
npm install @bilig/workpaper
npm install -D tsx typescript @types/node
cat > query-sortn.ts <<'EOF'
import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from "@bilig/workpaper";

type CellValue = {
  value?: unknown;
};

function readNumber(cell: unknown, label: string): number {
  if (typeof cell === "object" && cell !== null && typeof (cell as CellValue).value === "number") {
    return (cell as CellValue).value;
  }

  throw new Error(`expected ${label} to be numeric, got ${JSON.stringify(cell)}`);
}

const workbook = WorkPaper.buildFromSheets({
  Deals: [
    ["Region", "Segment", "Revenue"],
    ["West", "SMB", 60000],
    ["East", "Enterprise", 45000],
    ["West", "Enterprise", 140000],
    ["East", "SMB", 30000],
    ["West", "SMB", 36000],
  ],
  Summary: [
    ["Metric", "Value"],
    [
      "Top region revenue",
      '=INDEX(QUERY(Deals!A1:C6,"select A,sum(C) where C >= 30000 group by A order by sum(C) desc label A \'Region\', sum(C) \'Revenue\'",1),2,2)',
    ],
    ["Top deal", "=INDEX(SORTN(Deals!A2:C6,1,0,3,FALSE),1,3)"],
  ],
});

const summary = workbook.getSheetId("Summary");
const deals = workbook.getSheetId("Deals");
if (summary === undefined || deals === undefined) {
  throw new Error("missing sheet");
}

const before = readNumber(workbook.getCellValue({ sheet: summary, row: 1, col: 1 }), "before top region");
workbook.setCellContents({ sheet: deals, row: 3, col: 2 }, 190000);

const after = readNumber(workbook.getCellValue({ sheet: summary, row: 1, col: 1 }), "after top region");
const topDeal = readNumber(workbook.getCellValue({ sheet: summary, row: 2, col: 1 }), "top deal");
const saved = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }));
const restored = createWorkPaperFromDocument(parseWorkPaperDocument(saved));
const restoredSummary = restored.getSheetId("Summary");
if (restoredSummary === undefined) {
  throw new Error("missing restored Summary sheet");
}

const afterRestore = readNumber(
  restored.getCellValue({ sheet: restoredSummary, row: 1, col: 1 }),
  "restored top region",
);

console.log(
  JSON.stringify(
    {
      before,
      after,
      afterRestore,
      topDeal,
      persistedDocumentBytes: saved.length,
      verified: before === 236000 && after === 286000 && afterRestore === 286000 && topDeal === 190000,
    },
    null,
    2,
  ),
);

workbook.dispose();
restored.dispose();
EOF
npx tsx query-sortn.ts
```

Expected output:

```json
{
  "before": 236000,
  "after": 286000,
  "afterRestore": 286000,
  "topDeal": 190000,
  "persistedDocumentBytes": 1273,
  "verified": true
}
```

The exact byte count can change with runtime metadata. The important proof is
that `QUERY` recalculated after the edit, `SORTN` found the new top deal, and
the restored WorkPaper still read the same calculated value.

## When to use this

Use this path for:

- pricing or quote models that group rows by region, tier, or customer segment
- import validation that filters suspicious rows before accepting a file
- agent tools that need spreadsheet formulas without driving a browser grid
- backend tests where Google credentials would make the run flaky

Use Google Sheets or its API when the sheet is a shared hosted document,
permissions matter, or users expect to edit the source data in Google
Workspace.

## Related paths

- [Google Sheets API boundary](google-sheets-api-alternative-node-workpaper.md)
- [Formula language notes](formula-language.md)
- [Agent adoption kit](agent-adoption-kit.md)
- [Evaluate Bilig as an Agent MCP workbook tool](eval-agent-mcp.md)
- [Where Bilig is not Excel-compatible yet](where-bilig-is-not-excel-compatible-yet.md)

If this matches the workflow, run the proof before adopting it. If the proof
passes, star or watch the repository so other backend and agent engineers can
find the package:
<https://github.com/proompteng/bilig/stargazers>.

If a missing clause blocks adoption, open one concrete formula fixture:
<https://github.com/proompteng/bilig/discussions/new?category=general>.

## Sources

- Google Sheets QUERY function:
  <https://support.google.com/docs/answer/3093343>
- Google Sheets SORTN function:
  <https://support.google.com/docs/answer/7354624>
