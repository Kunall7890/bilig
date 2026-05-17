---
title: Why use Bilig?
published: true
description: A plain explanation of when Bilig is worth using, when it is not, and how to prove the fit with a runnable WorkPaper check.
tags: typescript, node, spreadsheet, formulas, workbook
canonical_url: https://proompteng.github.io/bilig/why-use-bilig.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Why use Bilig?

Use Bilig when a spreadsheet is already the clearest way to review the logic,
but production cannot depend on a person opening Excel or Google Sheets.

That usually means a backend service, queue worker, serverless route, CLI, test,
or agent tool needs to do four things:

1. write inputs into known cells;
2. recalculate dependent formulas;
3. read the exact output cells back;
4. save the edited workbook state and restore it later.

That is the narrow reason `@bilig/headless` exists. It gives TypeScript code a
`WorkPaper` object for workbook-shaped business logic.

## The problem it replaces

Teams often start with a spreadsheet because formulas are easier to audit than a
pile of duplicated application code. Then the spreadsheet escapes into a service
boundary:

- a pricing model needs to run from an API route;
- an import checker needs formula-backed validation;
- a payout or commission model needs tests and reproducible readback;
- an agent needs to edit assumptions and prove what changed;
- an XLSX template has formulas, but Node file libraries only expose stale
  cached values.

The usual failure mode is not that spreadsheets are bad. The failure mode is
that the spreadsheet becomes invisible to production code. Someone rewrites the
math in JavaScript, drives a UI with screenshots, or trusts cached XLSX formula
values that were calculated before the current input edit.

Bilig is for the cases where you want to keep the workbook shape and make it a
real runtime artifact.

## What you get

`@bilig/headless` gives you:

- programmatic sheets, cells, formulas, and structural edits;
- formula readback after writes;
- WorkPaper JSON export and restore;
- a local MCP server for agent tools;
- XLSX import/export through the `@bilig/headless/xlsx` subpath;
- runnable examples that prove write -> recalc -> read -> persist -> restore.

The important word is prove. A useful integration does not stop after writing a
formula. It reads the calculated cell, serializes the WorkPaper, restores it,
and checks that the same value comes back.

## When not to use it

Do not use Bilig when:

- the product is a visual spreadsheet grid for humans;
- the only job is reading or writing `.xlsx` files;
- you need full Excel macro, chart, pivot, or desktop automation behavior;
- a mature formula engine with the broadest function surface matters more than
  WorkPaper persistence;
- a hosted collaborative spreadsheet is the actual product requirement.

For those cases, start with a browser grid, SheetJS, ExcelJS, HyperFormula,
Microsoft Graph Excel, or Google Sheets API. Bilig is not trying to replace all
of them.

## Fast fit check

Run the npm proof in a blank directory:

```sh
mkdir bilig-headless-eval
cd bilig-headless-eval
npm init -y
npm pkg set type=module
npm install @bilig/headless
npm install -D tsx typescript @types/node
curl -fsSLo quickstart.ts https://proompteng.github.io/bilig/npm-eval.ts
npx tsx quickstart.ts
```

If it prints `verified: true`, the package can create a workbook, edit one
input, recalculate, persist JSON, restore it, and read the same calculated value
back.

Then run the route-shaped example:

```sh
git clone --depth 1 https://github.com/proompteng/bilig.git
cd bilig/examples/serverless-workpaper-api
npm install
npm run smoke
```

That example is closer to a real service: request-shaped inputs go into cells,
the quote decision recalculates, and the restored WorkPaper has to match the
post-edit value.

## One-sentence version

Bilig lets Node services and agents run spreadsheet-shaped business logic as a
testable WorkPaper: edit cells, recalculate formulas, read values, and persist
the workbook state without opening a spreadsheet app.

If that is the job you have, star or bookmark the repo so you can find it again:
<https://github.com/proompteng/bilig/stargazers>.

If the shape is close but blocked, open a concrete adoption blocker:
<https://github.com/proompteng/bilig/discussions/new?category=general>.
