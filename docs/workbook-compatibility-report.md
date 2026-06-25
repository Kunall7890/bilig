---
title: Workbook Compatibility Report
published: true
description: Local XLSX workbook risk report for Node services and coding agents that need to know which workbook features require investigation before trusting Bilig.
tags: xlsx, formulas, compatibility, agents, node
canonical_url: https://proompteng.github.io/bilig/workbook-compatibility-report.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Workbook Compatibility Report

Use this before wiring a real `.xlsx` workbook into a Node service or coding
agent. The report answers one question:

> If I point an agent or Node service at this workbook, what known risks should
> I investigate before I trust the outputs?

It is an inspector, not a grader. It does not say that a workbook is Excel
compatible, and it does not print a compatibility percentage.

## One command

```sh
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- \
  workbook-compatibility-report workbook.xlsx --json
```

For a no-project proof:

```sh
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- \
  workbook-compatibility-report --demo --json
```

The same proof is available through the evaluator door:

```sh
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- \
  bilig-evaluate --door workbook-compatibility --json
```

For an agent that will keep working through MCP after the risk report, use the
XLSX preflight transcript:

```sh
pnpm --dir examples/headless-workpaper run agent:mcp-xlsx-risk-preflight
```

That path starts `bilig-workpaper-mcp --from-xlsx`, calls
`analyze_workbook_risk`, edits one input through
`set_cell_contents_and_readback`, verifies dependent formula readback, and
exports the WorkPaper document.

## Fields to trust

The v1 report stays small on purpose:

- `verified: true`: the report completed locally.
- `workbook.formulaCellCount`: how much formula surface was inspected.
- `findings.unsupportedFunctions`: formulas that need review before readback is
  trusted.
- `findings.externalLinks`: linked workbook references and unresolved count.
- `findings.macroModules`: preserved VBA payloads that Bilig does not execute.
- `findings.volatileFunctions`: functions like `NOW` and `OFFSET`.
- `findings.pivotTables`: pivots and unsupported pivot surfaces.
- `findings.staleCachedFormulas`: cached formula values that changed when
  recalculated.
- `cacheInspection.uninspectedFormulaCellCount`: formulas left unchecked when a
  caller deliberately sets `--inspect-limit`.
- `risk.level` and `risk.reasons`: low, medium, or high with concrete reasons.
- `excelParity: "not_proven"`: the report does not certify desktop Excel
  behavior.

## Demo output

The checked-in proof artifact is
[`workbook-compatibility-report.json`](workbook-compatibility-report.json).
The important shape is:

```json
{
  "schemaVersion": "bilig-workbook-compatibility-report.v1",
  "verified": true,
  "workbook": {
    "sheetCount": 2,
    "formulaCellCount": 3
  },
  "findings": {
    "unsupportedFunctions": [{ "name": "CUBEVALUE", "count": 1 }],
    "externalLinks": { "count": 0, "unresolvedCount": 0, "refreshedCount": 0 },
    "macroModules": { "count": 0, "byteLength": 0 },
    "volatileFunctions": [{ "name": "NOW", "count": 1 }],
    "pivotTables": { "count": 0, "unsupportedCount": 0, "cacheOnlyCount": 0 },
    "staleCachedFormulas": { "count": 2 },
    "missingCachedFormulaValues": { "count": 1 }
  },
  "risk": {
    "level": "high",
    "reasons": ["unsupported functions: CUBEVALUE (1)"]
  },
  "excelParity": "not_proven"
}
```

The demo workbook is intentionally imperfect. A perfect workbook proves little;
this one proves the report can call out a known unsupported function, a volatile
function, and cache states without pretending to know full Excel parity.

## Human output

Without `--json`, the CLI prints a compact readout:

```text
Workbook analyzed. Risk level: HIGH
Findings:
- Unsupported functions: CUBEVALUE (1)
- External links: 0
- Macro modules: 0
- Pivot tables: 0
- Volatile functions: NOW (1)
- Formula cells: 3
- Stale cached formulas: 2
- Missing cached formula values: 1
This report identifies workbook features that may require investigation before using Bilig in a service or agent workflow. It is not an Excel compatibility certification.
```

## What this proves

- the workbook can be imported locally without a spreadsheet UI
- formula cells can be counted and inspected
- unsupported function, external-link, macro, volatile, pivot, and cache signals
  are machine-readable
- limited inspection is visible and raises risk when formula cells remain
  unchecked
- the report completed without uploading the workbook

## What this does not prove

This is not an Excel compatibility certification. It does not execute VBA,
refresh pivots, refresh external data sources, certify chart behavior, or prove
desktop Excel UI behavior. Do not add `compatibilityScore`,
`excelCompatibilityPercent`, or similar score fields around this report.

## Related

- [Agent XLSX risk preflight](agent-xlsx-risk-preflight.md)
- [Workbook Compatibility Report transcript](workbook-compatibility-report-transcript.md)
- [Where Bilig is not Excel-compatible yet](where-bilig-is-not-excel-compatible-yet.md)
