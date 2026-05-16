---
title: Show HN: Bilig runs small formula workbooks in Node
published: true
description: A maintainer note for Bilig with the npm check, benchmark numbers, limits, and the feedback that would make the project more useful.
tags: show-hn, typescript, node, spreadsheet, agents
canonical_url: https://proompteng.github.io/bilig/show-hn-formula-workbooks-node-services.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Show HN: Bilig runs small formula workbooks in Node

Bilig came out of a very plain problem: people like reviewing pricing rules,
payout checks, and import validators in spreadsheet form, but the code path
still has to run on a server.

I did not want another service that drives Excel through clicks and then trusts
a screenshot. I wanted a small object model that can load a workbook, change a
few inputs, recalculate formulas, read the result, and save the same workbook
state as JSON.

## Try the npm package

This starts from an empty directory and uses the published npm package. The
current checked package version is `@bilig/headless@0.16.24`.

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

Expected shape:

```json
{
  "before": 24000,
  "after": 38400,
  "afterRestore": 38400,
  "sheets": ["Inputs", "Summary"],
  "verified": true
}
```

The important line is `"verified": true`. The script changed an input cell,
read the recalculated formula value, saved WorkPaper JSON, restored it, and got
the same calculated output again.

## The part I care about

Evaluating `=A1+B1` is table stakes. The useful part is the whole loop:

- put typed inputs into stable cells
- recalculate dependent formulas after edits
- read computed values back from the same workbook state
- save formulas and values as JSON
- restore the workbook in CI and prove the answer did not change

Bilig exposes a `WorkPaper` object because the workbook state has to be part
of the contract, not an image someone eyeballed after the fact.

## Current numbers

The checked benchmark artifact currently records `76/100` mean-latency wins
against HyperFormula-style comparable workloads, and `75/100` workloads winning
both mean and p95.

The caveat is visible on purpose:
`lookup-approximate-sorted-large` is slower at p95 by `2.626x`.
Browser grid rendering is not part of this benchmark.

Read the benchmark note:
[what the WorkPaper benchmark proves](what-workpaper-benchmark-proves.md).

## Stuff it does not do

Bilig is not a finished Excel clone. It does not claim full Excel formula
parity, chart fidelity, macro execution, collaborative editing, or
faster p95 on every workload.

Use HyperFormula first when you primarily need a mature broad formula engine.
Use SheetJS or ExcelJS first when the main job is file reading, writing, or
styling. Use Google Sheets API first when a shared hosted spreadsheet and human
collaboration are the product requirement.

Use `@bilig/headless` when your Node code owns the workbook state and needs
formula readback, persistence, and restore checks.

## What would help

The most useful feedback is concrete:

- the workflow you tried
- the formula or workbook shape that blocked you
- whether the npm check worked on your machine
- the smallest example that would make you consider it for a real service

Open feedback here:
<https://github.com/proompteng/bilig/discussions/new?category=general>.

If this matches a service workflow you want to revisit later, star or bookmark
the repository:
<https://github.com/proompteng/bilig/stargazers>.

## Shareable post

Suggested HN title:

```text
Show HN: Bilig runs small formula workbooks in Node
```

Suggested short body:

```text
I maintain Bilig.

The use case is boring but real: sometimes a pricing rule, payout check, or
import validator is easiest to review as cells and formulas, but the production
path still has to run in Node.

I did not want a worker clicking around Excel or Google Sheets and then trusting
a screenshot. Bilig gives you a WorkPaper object instead: write inputs,
recalculate, read values back, save JSON, restore it, and check the same answer
again.

The npm check starts from an empty directory and proves that loop with the
published package.

It is not an Excel clone. It will not run macros or preserve every weird XLSX
artifact. The current benchmark artifact says 76/100 mean wins against
HyperFormula-style comparable workloads, with the p95 misses called out instead
of hidden.

I am looking for blunt rejection reasons from people who have shipped
spreadsheet-backed services: missing formulas, XLSX cases, API shape, or the
benchmark that would make you trust or reject it quickly.
```
