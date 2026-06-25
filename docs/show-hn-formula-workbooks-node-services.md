---
title: 'Bilig maintainer note: formula WorkPapers for Node services and agents'
published: true
description: A plain maintainer note for Bilig with the WorkPaper npm check, benchmark artifact, limits, and open questions.
tags: typescript, node, spreadsheet, agents, workpaper
canonical_url: https://proompteng.github.io/bilig/show-hn-formula-workbooks-node-services.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Bilig maintainer note: formula WorkPapers for Node services and agents

I built Bilig for the rules that live in cells long after the product around
them moved to code: quotes, payout checks, approvals, import validation, budget
guards.

Those workflows should not need Excel screen driving. A service or coding agent
should be able to change an input cell, recalculate the workbook, read the
answer, and keep the workbook state under test.

Bilig is the WorkPaper runtime for that loop. `@bilig/workpaper` is the
canonical package for service code, CLI evaluators, MCP tools, and no-key
readback checks.

That is the whole pitch: cells stay reviewable, Node gets an API, and agents get
readback instead of screenshots.

## Fastest check

This uses the latest published package and starts from an empty directory:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

The output should look like this:

```json
{
  "schemaVersion": "bilig-evaluator.v1",
  "door": "agent-mcp",
  "verified": true
}
```

The line that matters is `"verified": true`. The evaluator starts a WorkPaper
tool server, discovers tools, edits an input cell, recalculates a dependent
formula, exports WorkPaper JSON, restores it, and checks that readback still
matches.

For direct library use:

```sh
npm install @bilig/workpaper
```

The 90-second quickstart is here:
[try Bilig WorkPaper in Node](try-bilig-headless-in-node.md).

## What Bilig is

- write typed inputs into known cells
- recalculate dependent formulas
- read calculated values back from the same state
- save formulas and values as JSON
- restore the workbook later and check the answer again

The API is built around a `WorkPaper` object because the workbook state is the
artifact under test. Saved-file tools are useful when an existing workbook file
is the contract, but the product is the WorkPaper runtime.

## Current numbers

The checked benchmark artifact currently says Bilig wins `100/100` comparable
workloads on mean latency against the HyperFormula-style baseline. It wins
`100/100` on both mean and p95.

The worst p95 row is not hidden: `sheet-rename-dependencies` is the current worst p95 row at `0.792x`.
Browser grid rendering is not part of this benchmark.

Read the benchmark note:
[what the WorkPaper benchmark proves](what-workpaper-benchmark-proves.md).

## What this is not

Bilig is not Excel in Node. It does not run macros, preserve every workbook
artifact, cover every Excel formula, do collaborative editing, or prove future
p95 cases without adding them to the checked suite.

If you mainly need a mature broad formula engine, start with HyperFormula. If
the problem is XLSX reading, writing, or styling, start with SheetJS or ExcelJS.
If the product is a shared hosted spreadsheet, use Google Sheets.

Use `@bilig/workpaper` when your Node code can own workbook state and needs
formula readback, persistence, and restore checks.

## What would help

I am looking for rejection reasons:

- a formula family that blocks a real workbook
- a workbook shape that breaks the model
- a runtime or deployment target where the package is painful
- an API shape that makes this awkward in a real service
- a benchmark you would need before trusting it

Open feedback here:
<https://github.com/proompteng/bilig/discussions/new?category=general>.

## Review checklist

Before adopting it, verify the narrow path you need:

- `bilig-evaluate --door workpaper-service --json` passes for service-owned
  WorkPaper state.
- `bilig-evaluate --door agent-mcp --json` passes for tool discovery, write,
  readback, export, and restore.
- The benchmark page matches the checked artifact instead of an isolated number.
- The compatibility page rules out any Excel-only feature your workflow needs.
- Any real blocked workbook is reduced into a fixture before it becomes an
  accuracy claim.
