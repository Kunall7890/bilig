---
title: Agent XLSX Risk Preflight
published: true
description: Run a local MCP XLSX risk diagnostic before a coding agent edits formulas, then require formula readback and WorkPaper export proof.
tags: agents, mcp, xlsx, formulas, workpaper
canonical_url: https://proompteng.github.io/bilig/agent-xlsx-risk-preflight.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Agent XLSX Risk Preflight

Use this when an agent already has a real `.xlsx` file and is about to automate
spreadsheet edits. The first useful question is not whether the write call
succeeded. It is whether the imported workbook has risk indicators that should
change the plan before any edit is trusted.

This path starts the Bilig WorkPaper MCP server from the XLSX, calls the
read-only `analyze_workbook_risk` tool, then proves a formula edit through
`set_cell_contents_and_readback` and `export_workpaper_document`.

## Run It

From a cloned checkout:

```sh
pnpm --dir examples/headless-workpaper install --ignore-workspace
pnpm --dir examples/headless-workpaper run agent:mcp-xlsx-risk-preflight
```

The example builds a small `pricing-risk-preflight.xlsx`, starts the published
MCP binary, imports the XLSX into a persisted WorkPaper JSON file, analyzes
workbook risk, edits `Inputs!B3`, reads back `Summary!B3`, and exports the
WorkPaper document.

The underlying server command is:

```sh
npm exec --package @bilig/workpaper@latest -- \
  bilig-workpaper-mcp \
  --from-xlsx pricing-risk-preflight.xlsx \
  --workpaper pricing-risk-preflight.workpaper.json \
  --writable
```

## Required Tool Order

An agent should use this order for a private workbook:

1. `tools/list`
2. `tools/call` `analyze_workbook_risk` with `inspectLimit: "all"`
3. Decide whether the workbook can continue through WorkPaper or needs Excel,
   LibreOffice, Microsoft Graph, or a human/oracle path.
4. `tools/call` `set_cell_contents_and_readback` for one small input edit.
5. `tools/call` `export_workpaper_document` for handoff or persistence proof.

The risk step is fixed to the XLSX file passed at server startup. It is
read-only, local, and does not upload the workbook.

## Expected Proof

A passing run prints a compact JSON object:

```json
{
  "schemaVersion": "bilig-agent-xlsx-risk-preflight.v1",
  "transport": "stdio",
  "risk": {
    "schemaVersion": "bilig-workbook-compatibility-report.v1",
    "verified": true,
    "fileName": "pricing-risk-preflight.xlsx",
    "formulaCellCount": 3,
    "excelParity": "not_proven"
  },
  "readback": {
    "editedCell": "Inputs!B3",
    "beforeExpectedArr": 60000,
    "afterExpectedArr": 96000,
    "restoredExpectedArr": 96000,
    "persisted": true,
    "restoredReadbackMatchesAfter": true
  },
  "verified": true
}
```

Those fields matter more than the exact serialized byte count. The result is
usable only when the risk diagnostic is `verified: true`, `excelParity` remains
`"not_proven"`, the dependent formula readback changed from `60000` to `96000`,
and restored state still reads `96000`.

## What It Proves

- the agent used a real XLSX file, not a hand-waved workbook description
- the local MCP tool surface exposed `analyze_workbook_risk`
- workbook risk indicators were inspected before edits
- a dependent formula was recalculated after `Inputs!B3` changed
- the edited WorkPaper was exported or persisted for another process to check

## What It Does Not Prove

This is not an Excel compatibility certification. It does not execute VBA,
refresh pivots, refresh external data, prove chart layout, or certify desktop
Excel UI behavior. If the risk report flags unsupported functions, macros,
pivots, external links, or workbook features that Bilig does not cover, keep
Excel, LibreOffice, Microsoft Graph, or a spreadsheet-specific oracle in the
loop.

## Related

- [MCP WorkPaper tool server](mcp-workpaper-tool-server.md)
- [Workbook Compatibility Report](workbook-compatibility-report.md)
- [Agent WorkPaper proof matrix](agent-proof-matrix.md)
- [Stale formula readback chooser](stale-formula-readback-chooser.md)
- [Where Bilig is not Excel-compatible yet](where-bilig-is-not-excel-compatible-yet.md)
