---
title: Show HN: Bilig runs small formula workbooks in Node
published: true
description: A plain maintainer note for Bilig with the npm check, benchmark artifact, limits, and concrete feedback ask.
tags: show-hn, typescript, node, spreadsheet, agents
canonical_url: https://proompteng.github.io/bilig/show-hn-formula-workbooks-node-services.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Show HN: Bilig runs small formula workbooks in Node

I built Bilig because I kept running into the same awkward gap: the business
logic was easiest to discuss as cells and formulas, but the actual service
could not depend on a person opening a spreadsheet app.

The use cases are not glamorous: quote checks, payout rules, import validators,
small revenue models. Somebody wants to review the formulas. The backend still
needs to write inputs, recalculate, read the answer, and keep a durable record
of exactly what ran.

That is the narrow thing Bilig is trying to do.

## Try the npm package

This starts from an empty directory and uses the published package. The current
checked version is `@bilig/headless@0.16.25`.

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

The interesting line is `"verified": true`. The script changes an input cell,
reads a recalculated value, serializes the WorkPaper JSON, restores it, and
gets the same calculated value again.

## What this is

Bilig is a small workbook-state API for Node. The point is the full loop, not
just `=A1+B1`:

- write typed inputs into known cells
- recalculate dependent formulas
- read calculated values back from the same state
- save formulas and values as JSON
- restore the workbook later and check the answer again

The API is built around a `WorkPaper` object because I want the workbook state
to be the artifact under test. A screenshot is not enough for this kind of
workflow.

## Current numbers

The checked benchmark artifact currently records `78/100` mean-latency wins on
HyperFormula-style comparable workloads, with `74/100` workloads winning on
both mean and p95.

The caveat matters: `single-formula-edit-recalc` is slower at p95 by `2.608x`.
Browser grid rendering is not part of this benchmark.

Read the benchmark note:
[what the WorkPaper benchmark proves](what-workpaper-benchmark-proves.md).

## What this is not

Bilig is not a finished Excel clone. It does not claim full Excel formula
parity, chart fidelity, macro execution, collaborative editing, or
faster p95 on every workload.

If you mainly need a mature broad formula engine, start with HyperFormula. If
the problem is XLSX reading, writing, or styling, start with SheetJS or ExcelJS.
If the product is a shared hosted spreadsheet, use Google Sheets.

Use `@bilig/headless` when your Node code owns the workbook state and needs
formula readback, persistence, and restore checks.

## What would help

The feedback I care about is the kind that would make you reject it quickly:

- a formula family that is missing
- a workbook shape that breaks the model
- a Node runtime or deployment target where the package is annoying
- an API shape that makes this hard to wire into a real service
- a benchmark you would want before trusting it

Open feedback here:
<https://github.com/proompteng/bilig/discussions/new?category=general>.

If this is a problem you might come back to, star or bookmark the repository:
<https://github.com/proompteng/bilig/stargazers>.

## Shareable post

Suggested HN title:

```text
Show HN: Bilig runs small formula workbooks in Node
```

Suggested short body:

```text
I maintain Bilig. It is a small Node library for running formula-backed
workbook state without opening Excel or Google Sheets.

The use case is fairly boring: quote checks, payout rules, import validators,
small revenue models. People want to review those rules as cells and formulas,
but the service still needs a real API path: write inputs, recalculate, read
the result, save the state, and restore it later.

Bilig exposes that as a WorkPaper object. The quick npm check starts from an
empty directory and proves the loop with the published package.

It is not an Excel clone. It does not run macros, preserve every XLSX artifact,
or claim full Excel compatibility. The current benchmark artifact says 78/100
mean wins on HyperFormula-style comparable workloads, and the p95 misses are
called out on the page.

I am looking for rejection reasons from people who have shipped spreadsheet-ish
backend workflows: missing formulas, XLSX cases, API shape, runtime constraints,
or the benchmark that would make you trust or reject it faster.
```
