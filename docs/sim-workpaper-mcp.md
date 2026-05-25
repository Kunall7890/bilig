---
title: Sim WorkPaper MCP setup
published: true
description: Connect Bilig WorkPaper to Sim MCP tools so workflows and agents can edit workbook inputs, verify formula readback, and pass proof into later blocks.
tags: sim, mcp, spreadsheet-agent, workbook-api, workflow-automation
canonical_url: https://proompteng.github.io/bilig/sim-workpaper-mcp.html
image: /assets/github-social-preview.png
---

# Sim WorkPaper MCP setup

Use this when a Sim workflow needs spreadsheet-shaped business logic, but the
formula state should live behind explicit WorkPaper tools instead of Excel UI
automation, browser grid clicks, or stale XLSX cached values.

Sim's MCP tool setup adds external MCP servers from **Settings -> MCP Tools**.
Sim documents Streamable HTTP server URLs, connection testing, Agent-block tool
use, and a standalone MCP Tool block for deterministic calls.

Official Sim references:

- <https://docs.sim.ai/mcp>
- <https://docs.sim.ai/mcp/deploy-workflows>

## Fastest smoke test: hosted Streamable HTTP

Use this when you only need to prove that Sim can discover and call the Bilig
WorkPaper tools.

In Sim:

1. Open **Settings -> MCP Tools**.
2. Click **Add**.
3. Set **Server Name** to `bilig-workpaper`.
4. Set **Server URL** to `https://bilig.proompteng.ai/mcp`.
5. Leave headers empty.
6. Keep transport as Streamable HTTP.
7. Click **Test Connection** and confirm seven tools are discovered.
8. Save the server.

The hosted endpoint is stateless and request-local. It proves tool discovery and
formula readback, but it does not persist a private project file.

## Agent block proof

Use this when the workflow should let the model choose the WorkPaper tool calls.

1. Open an Agent block.
2. Add tools from the `bilig-workpaper` MCP server.
3. Select all seven tools.
4. Use a prompt that requires readback and persistence proof:

```text
Use the Bilig WorkPaper MCP tools. List the tools, read the sample sheets, set
Inputs!B3 to 0.4, read Summary!B3, export the WorkPaper document, and return
editedCell, before, after, afterRestore, persistedDocumentBytes, verified, and
limitations. Do not claim success from a write call alone.
```

The useful Bilig tools are:

- `list_sheets`
- `read_range`
- `read_cell`
- `set_cell_contents`
- `get_cell_display_value`
- `export_workpaper_document`
- `validate_formula`

Expected proof fields include:

```json
{
  "editedCell": "Inputs!B3",
  "dependentCell": "Summary!B3",
  "before": 60000,
  "after": 96000,
  "verified": true
}
```

`verified` should only be true after the dependent formula output is read back.

## Standalone MCP Tool block

Use Sim's standalone MCP Tool block when the workflow step should be
deterministic instead of model-selected.

One practical shape:

1. `read_cell` or `read_range` to capture the current input and dependent output.
2. `set_cell_contents` with `Inputs!B3 = 0.4`.
3. `get_cell_display_value` for `Summary!B3`.
4. `export_workpaper_document` so downstream blocks can store or inspect the
   WorkPaper proof object.

That keeps the calculation repeatable: Sim owns the workflow routing, and Bilig
owns the formula workbook state and readback contract.

## Private workbook state

The hosted endpoint is only a smoke test. For a private or writable project
WorkPaper, expose your own Bilig WorkPaper MCP endpoint on a domain that your
Sim workspace can reach, then add that URL in **Settings -> MCP Tools**.

For self-hosted Sim deployments with domain allowlisting, include the private
Bilig MCP host in `ALLOWED_MCP_DOMAINS`.

Before putting a private endpoint behind Sim, prove the file-backed local MCP
contract from a terminal:

```sh
npx -y --package @bilig/workpaper@latest bilig-mcp-challenge --json
```

Expected local proof:

```json
{
  "transport": "stdio-json-rpc",
  "editedCell": "Inputs!B3",
  "dependentCell": "Summary!B3",
  "before": 60000,
  "after": 96000,
  "afterRestart": 96000,
  "verified": true
}
```

## Boundaries

- Sim connects to MCP server URLs over Streamable HTTP. Do not paste a local
  stdio command into Sim's Server URL field.
- Hosted Streamable HTTP is stateless. Use a private reachable Bilig MCP
  endpoint when a workflow needs durable workbook state.
- Agent blocks let the model choose tools. Use the standalone MCP Tool block for
  structured, repeatable workflow steps.
- Keep Excel or another workbook oracle in the loop for macros, pivots, charts,
  external links, and layout fidelity.

## Related Bilig docs

- [Agent MCP workbook evaluator](eval-agent-mcp.md)
- [MCP WorkPaper tool server](mcp-workpaper-tool-server.md)
- [MCP client setup](mcp-client-setup.md)
- [Agent framework workbook tools](agent-framework-workbook-tools.md)
- [Headless WorkPaper agent handbook](headless-workpaper-agent-handbook.md)
