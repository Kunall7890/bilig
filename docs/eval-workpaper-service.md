---
title: Evaluate WorkPaper in a Node service
published: true
description: Copy-paste evaluator for backend services that need formula workbook state, input writes, readback, JSON persistence, and restore proof.
tags: node, typescript, workpaper, formulas, evaluator
canonical_url: https://proompteng.github.io/bilig/eval-workpaper-service.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Evaluate WorkPaper in a Node service

Use this when the calculation model belongs in code, not in a user-edited Excel
file. The evaluator starts from an empty directory, creates a small WorkPaper
service, writes one input, reads a dependent formula, serializes the WorkPaper
document, restores it, and verifies the same result.

## One command

```sh
npm create @bilig/workpaper@latest pricing-workpaper && cd pricing-workpaper && npm install && npm run smoke
```

## Expected proof

The starter smoke prints this shape:

```json
{
  "before": 24000,
  "after": 38400,
  "afterRestore": 38400,
  "sheets": ["Inputs", "Summary"],
  "bytes": 999,
  "verified": true,
  "star": "https://github.com/proompteng/bilig/stargazers",
  "watchReleases": "https://github.com/proompteng/bilig/subscription",
  "adoptionBlocker": "https://github.com/proompteng/bilig/discussions/new?category=general",
  "nextStep": "If this proof matches your workflow, open a concrete blocker or adoption note: https://github.com/proompteng/bilig/discussions/new?category=general"
}
```

The byte count can change by package version. The invariant is that `after`
comes from the dependent formula cell and `afterRestore` matches `after`.

## What this proves

- a service can own workbook-shaped business logic as WorkPaper JSON
- input cells can be changed through an API instead of a UI
- dependent formulas recalculate before the service responds
- exported WorkPaper state can be restored and re-read
- the proof object is small enough for tests, logs, or agent handoff

## What this does not prove

This does not prove full XLSX fidelity, desktop Excel behavior, database
durability, or a visual spreadsheet editor. Use this path when the service owns
the formulas and JSON state. Use the XLSX evaluator when a real `.xlsx` file is
the source of truth.

## After the proof

- Star Bilig if this is the service shape you needed:
  <https://github.com/proompteng/bilig/stargazers>
- Watch releases for API and formula runtime updates:
  <https://github.com/proompteng/bilig/subscription>
- Report the exact adoption blocker:
  <https://github.com/proompteng/bilig/discussions/new?category=general>

## Related

- [Try Bilig WorkPaper in Node](try-bilig-headless-in-node.md)
- [Create a Bilig WorkPaper starter](create-bilig-workpaper.md)
- [WorkPaper service recipe](node-service-workpaper-recipe.md)
- [Quote approval WorkPaper API](quote-approval-workpaper-api.md)
