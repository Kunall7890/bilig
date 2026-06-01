---
title: Open Multi-Agent WorkPaper MCP example
published: true
description: Connect Open Multi-Agent to Bilig WorkPaper MCP tools so an agent can edit workbook inputs, verify formula readback, and persist JSON without spreadsheet UI automation.
tags: open-multi-agent, mcp, spreadsheet-agent, workbook-api, typescript
canonical_url: https://proompteng.github.io/bilig/open-multi-agent-workpaper-mcp.html
image: /assets/github-social-preview.png
---

# Open Multi-Agent WorkPaper MCP example

Use this when an Open Multi-Agent workflow needs spreadsheet formulas but should
not drive Excel, Google Sheets, or a browser UI. Bilig keeps the workbook as a
file-backed WorkPaper and exposes only explicit MCP tools for reads, writes,
formula validation, display-value readback, and JSON export.

## Open Multi-Agent example

The Open Multi-Agent integration example was merged here:

- <https://github.com/open-multi-agent/open-multi-agent/pull/247>

It uses Open Multi-Agent's `connectMCPTools()` helper to launch
`bilig-workpaper-mcp` over stdio, registers the returned tools with an
`Agent`, and asks the agent to:

1. list the workbook sheets;
2. read a calculated summary cell;
3. set one input cell;
4. read the calculated summary cell again;
5. report whether the WorkPaper recalculated and persisted the edit.

The upstream example pins the Bilig package version in the command instead of
exposing the npm package spec to the model. In a new project, pin deliberately
after checking the current Bilig evaluator.

## No-key Bilig check first

Before wiring an Open Multi-Agent model provider, run Bilig's package-owned MCP
evaluator. It does not need a Gemini, OpenAI, or other model key:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

A good run returns `verified: true`, `editedCell: "Inputs!B3"`,
`dependentCell: "Summary!B3"`, formula readback `60000 -> 96000`, JSON export,
and restart readback.

## Local command shape

The MCP server command used by the example is:

```sh
npm exec --yes --package @bilig/workpaper@latest -- \
  bilig-workpaper-mcp \
  --workpaper ./pricing.workpaper.json \
  --init-demo-workpaper \
  --writable
```

For production, replace `@latest` with the current version you verified and keep
the WorkPaper file under the application's normal data directory.

## Agent contract

Give the agent a narrow tool contract:

```text
Use Bilig WorkPaper MCP tools to inspect and edit formula workbooks. Always
verify a write by reading the recalculated output cell afterward. Keep the
final answer short and include the before and after values.
```

Do not let the model claim success from a write call alone. A valid result
needs both mutation and readback evidence.

## Proof shape

Ask for a proof object like this:

```json
{
  "editedCell": "Inputs!B3",
  "before": {
    "Summary!B3": "60000"
  },
  "after": {
    "Summary!B3": "96000"
  },
  "persistedDocumentBytes": 1000,
  "verified": true,
  "limitations": [
    "Pinned @bilig/workpaper version should be refreshed deliberately.",
    "The demo WorkPaper is local to the process unless you choose a stable file path."
  ]
}
```

`verified` should only be true after the dependent formula output was read back
from the WorkPaper after the input edit.

## Related Bilig docs

- [Agent framework workbook tools](agent-framework-workbook-tools.md)
- [MCP WorkPaper tool server](mcp-workpaper-tool-server.md)
- [Headless WorkPaper agent handbook](headless-workpaper-agent-handbook.md)
- [Agent WorkPaper tool-calling recipe](agent-workpaper-tool-calling-recipe.md)
- [Why agents need workbook APIs](why-agents-need-workbook-apis.md)
