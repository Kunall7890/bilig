---
title: Quote approval WorkPaper API in Node
published: true
description: Run a quote approval API smoke from an empty Node.js directory with @bilig/headless, formula recalculation, JSON persistence, and restored readback.
tags: typescript, node, spreadsheet, formulas, api
canonical_url: https://proompteng.github.io/bilig/quote-approval-workpaper-api.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Quote approval WorkPaper API in Node

Use this page when you want a production-shaped `@bilig/headless` proof instead
of a toy arithmetic workbook.

The smoke runs a quote approval workflow:

1. Build a two-sheet WorkPaper with `Inputs` and `Summary`.
2. Write quote input cells: units, list price, discount, unit cost, and minimum
   margin.
3. Recalculate formulas for net revenue, gross margin, and approval decision.
4. Serialize the WorkPaper document as JSON.
5. Restore that JSON and verify the restored workbook still matches the
   recalculated result.

No browser grid, spreadsheet account, OAuth setup, or repo clone is required.

## Run It From An Empty Directory

```sh
mkdir bilig-quote-approval
cd bilig-quote-approval
npm init -y
npm pkg set type=module
npm install @bilig/headless
npm install -D tsx typescript @types/node
curl -fsSLo quote-approval-api.ts \
  https://raw.githubusercontent.com/proompteng/bilig/main/examples/serverless-workpaper-api/quote-approval-api.ts
npx tsx quote-approval-api.ts
```

Expected shape:

```json
{
  "route": "Quote approval WorkPaper API",
  "inputCells": {
    "units": "Inputs!B2",
    "listPrice": "Inputs!B3",
    "discount": "Inputs!B4",
    "unitCost": "Inputs!B5",
    "minimumMargin": "Inputs!B6"
  },
  "before": {
    "netRevenue": 43200,
    "grossMargin": 0.2963,
    "decision": "review"
  },
  "edit": {
    "input": {
      "units": 40,
      "listPrice": 1200,
      "discount": 0.05,
      "unitCost": 760,
      "minimumMargin": 0.3
    },
    "after": {
      "netRevenue": 45600,
      "grossMargin": 0.3333,
      "decision": "approved"
    },
    "checks": {
      "decisionChanged": true,
      "formulasPersisted": true,
      "inputPersisted": true,
      "restoredMatchesAfter": true
    }
  },
  "verified": true,
  "star": "https://github.com/proompteng/bilig/stargazers",
  "watchReleases": "https://github.com/proompteng/bilig/subscription",
  "adoptionBlocker": "https://github.com/proompteng/bilig/discussions/new?category=general",
  "nextStep": "If this proof matches your workflow, open a concrete blocker or adoption note: https://github.com/proompteng/bilig/discussions/new?category=general"
}
```

The exact serialized byte count can move between releases. The important parts
are:

- `decisionChanged: true`
- `formulasPersisted: true`
- `inputPersisted: true`
- `restoredMatchesAfter: true`
- `verified: true`

## What This Proves

This is the service boundary that matters for backend adoption:

- input JSON maps to known workbook cells
- formulas recalculate after the write
- the returned values come from formula readback, not a screenshot
- the persisted JSON still contains formulas
- a restored WorkPaper returns the same decision

That is the shape behind pricing rules, discount approval, payout checks,
budget guardrails, import validation, and agent tools that need exact readback.

## What This Does Not Prove

It does not prove full Excel compatibility. It does not prove formatting,
charts, collaboration, or broad XLSX file fidelity. For those boundaries, read
the [compatibility limits](where-bilig-is-not-excel-compatible-yet.md) and the
[production adoption checklist](production-adoption-checklist-headless-workpaper.md).

It also does not prove every formula family you need is implemented. If this API
shape is right but a formula, persistence shape, or framework boundary blocks a
trial, open a concrete adoption note in the
[workflow feedback discussion](https://github.com/proompteng/bilig/discussions/157).

## Use It In A Service

The full example is
[`examples/serverless-workpaper-api`](https://github.com/proompteng/bilig/tree/main/examples/serverless-workpaper-api).
It includes:

- a web-standard `Request` / `Response` route handler
- a quote approval route
- a Vercel Function smoke
- a Next.js App Router smoke
- framework adapters
- persistence adapter examples

Run the wider proof from a repo checkout:

```sh
pnpm --dir examples/serverless-workpaper-api install --ignore-workspace
pnpm --dir examples/serverless-workpaper-api run test
pnpm --dir examples/serverless-workpaper-api run framework-adapters
pnpm --dir examples/serverless-workpaper-api run persistence-adapters
```

## Next Pages

- [Try `@bilig/headless` in Node](try-bilig-headless-in-node.md)
- [Serverless WorkPaper API route](serverless-workpaper-api-route.md)
- [Node service WorkPaper recipe](node-service-workpaper-recipe.md)
- [Five Node.js workbook automation examples](workbook-automation-examples-node.md)
- [What the WorkPaper benchmark proves](what-workpaper-benchmark-proves.md)
- [Where bilig is not Excel-compatible yet](where-bilig-is-not-excel-compatible-yet.md)

If it almost matches but a gap blocks adoption, use the adoption blocker form:
<https://github.com/proompteng/bilig/discussions/new?category=general>.
