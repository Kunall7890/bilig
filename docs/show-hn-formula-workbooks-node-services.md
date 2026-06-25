---
title: 'Show HN: Bilig - formula WorkPapers for Node services and agents'
published: true
description: A plain maintainer note for Bilig with the WorkPaper npm check, benchmark artifact, limits, and open questions.
tags: show-hn, typescript, node, spreadsheet, agents
canonical_url: https://proompteng.github.io/bilig/show-hn-formula-workbooks-node-services.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Show HN: Bilig - formula WorkPapers for Node services and agents

I built Bilig for the rules that live in cells long after the product around
them moved to code: quotes, payout checks, approvals, import validation, budget
guards.

Those workflows should not need Excel screen driving. A service or coding agent
should be able to change an input cell, recalculate the workbook, read the
answer, and keep the workbook state under test.

Bilig is the WorkPaper runtime for that loop. `@bilig/headless` owns the
formula-backed workbook state. `@bilig/workpaper` packages the CLI, MCP server,
and no-key evaluators around the same model.

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
npm install @bilig/headless
```

The 90-second quickstart is here:
[try Bilig headless in Node](try-bilig-headless-in-node.md).

## What Bilig is

- write typed inputs into known cells
- recalculate dependent formulas
- read calculated values back from the same state
- save formulas and values as JSON
- restore the workbook later and check the answer again

The API is built around a `WorkPaper` object because the workbook state is the
artifact under test. The XLSX cache tools are useful when a file boundary is the
problem, but the product is the WorkPaper runtime.

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

Use `@bilig/headless` when your Node code can own workbook state and needs
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

## Shareable post

Suggested HN title:

```text
Show HN: Bilig - formula WorkPapers for Node services and agents
```

Suggested first comment, only if the story is live and the maintainer can stay
in the thread:

```text
I built Bilig because a lot of pricing, approval, payout, and import rules are
still workbook-shaped, but backend services and coding agents need a state API
rather than screenshots.

The quickest check is:

`npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json`

It writes an input cell, recalculates, reads a dependent formula, saves and
restores WorkPaper JSON, and returns `verified: true`.

There is also a public no-key Hugging Face Space for the MCP readback path:
https://huggingface.co/spaces/gregkonush/bilig-workpaper-mcp-readback

Useful feedback: API friction, missing formula semantics, MCP/readback shape, or
real workbook cases that should become fixtures. This is not a full Excel
clone; XLSX cache doctor is one doorway into the WorkPaper runtime.
```
