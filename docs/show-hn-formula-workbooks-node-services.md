---
title: Show HN: formula workbooks for Node services
published: true
description: A plain maintainer note for Bilig with the npm check, benchmark numbers, limits, and the feedback that would make the project more useful.
tags: show-hn, typescript, node, spreadsheet, agents
canonical_url: https://proompteng.github.io/bilig/show-hn-formula-workbooks-node-services.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Show HN: formula workbooks for Node services

Bilig is a TypeScript WorkPaper runtime for a familiar awkward case: the
calculation is easiest to review as cells and formulas, but the service needs to
run it from Node instead of from Excel, Google Sheets, or browser automation.

The fit is narrow on purpose: pricing rules, quote approval, payout checks,
budget guardrails, import validation, and tool calls that need to change workbook
inputs and read the calculated result back.

## Try the npm package

This starts from an empty directory and uses the published npm package. The
current checked package version is `@bilig/headless@0.16.22`.

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

The important line is `"verified": true`: the script edited an input cell, read
the recalculated formula value, serialized the workbook as WorkPaper JSON, and
restored it with the same calculated output.

## Why not just a formula parser

The useful boundary is not just evaluating `=A1+B1`. A service or agent usually
needs the whole loop:

- map typed inputs to stable workbook cells
- recalculate dependent formulas after edits
- read computed values back from the workbook runtime
- persist formulas and values as JSON
- restore the workbook and prove the same output in CI

Bilig exposes a `WorkPaper` object because the workbook state matters as much as
the scalar formula result.

## Evidence

The checked benchmark artifact currently records `76/100` mean-latency wins
against HyperFormula-style comparable workloads, and `74/100` workloads winning
both mean and p95.

The caveat is intentionally visible: `named-expression-change` is slower at
p95 by `3.497x`. Browser grid rendering is outside this benchmark.

Read the benchmark note:
[what the WorkPaper benchmark proves](what-workpaper-benchmark-proves.md).

## What it is not

Bilig is not a finished Excel clone. It does not claim full Excel formula
parity, chart fidelity, macro execution, collaborative spreadsheet editing, or
faster p95 on every workload.

Use HyperFormula first when you primarily need a mature broad formula engine.
Use SheetJS or ExcelJS first when the main job is file reading, writing, or
styling. Use Google Sheets API first when a shared hosted spreadsheet and human
collaboration are the product requirement.

Use `@bilig/headless` when a Node service or tool owns the workbook state and
needs formula readback, persistence, and restore checks.

## If you are evaluating it

The most useful feedback is concrete:

- the workflow you tried
- the formula or workbook shape that blocked you
- whether the npm check worked on your machine
- the smallest example that would make you try it in a real service

Open feedback here:
<https://github.com/proompteng/bilig/discussions/new?category=general>.

If this matches a service or tool workflow you want to revisit, star or bookmark
the repository:
<https://github.com/proompteng/bilig/stargazers>.

## Shareable post

Suggested HN title:

```text
Show HN: Formula workbooks for Node services and agent tools
```

Suggested short body:

```text
I built Bilig because I kept hitting the same awkward shape: the business rule
was clearest as a small workbook, but the service needed to run it in Node and
test the result in CI.

The npm check starts from an empty project, edits an input cell, reads the
recalculated formula value, serializes WorkPaper JSON, restores it, and checks
the same output again.

It is not an Excel clone. The current benchmark artifact says 76/100 mean wins
against HyperFormula-style comparable workloads and `74/100` workloads winning both mean and p95. One
visible p95 holdout is named-expression-change.

I am looking for concrete misses: formula coverage, XLSX import/export,
persistence shape, MCP/tool use, or a benchmark that would make you trust or
reject it faster.
```
