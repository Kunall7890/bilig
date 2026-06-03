---
title: Agent WorkPaper proof matrix
published: true
description: Pick the smallest Bilig WorkPaper proof for coding agents, MCP clients, AI SDK tools, OpenAI tool calls, LangGraph, Semantic Kernel, Mastra, and XLSX recalculation.
tags: ai-agents, mcp, workpaper, spreadsheet-agent, proof
canonical_url: https://proompteng.github.io/bilig/agent-proof-matrix.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Agent WorkPaper Proof Matrix

Use this page before an agent drives Excel, LibreOffice, Google Sheets, or a
browser grid. Pick the smallest proof that writes an input, recalculates a
dependent formula, reads the value back, and preserves enough state for another
process to check the result.

If you only run one command, run the agent MCP evaluator:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

Expected invariants:

```json
{
  "schemaVersion": "bilig-evaluator.v1",
  "door": "agent-mcp",
  "verified": true,
  "evidence": {
    "editedCell": "Inputs!B3",
    "dependentCell": "Summary!B3",
    "before": 60000,
    "after": 96000,
    "afterRestore": 96000,
    "afterRestart": 96000,
    "persistedDocumentBytes": 1162,
    "checks": {
      "listedFileBackedTools": true,
      "listedResourcesAndPrompts": true,
      "dependentCellChanged": true,
      "persistedToDisk": true,
      "restartReadbackMatchesAfter": true
    }
  }
}
```

## Proof Matrix

| Proof | Command or asset | Expected JSON field | What it proves | What it does not prove |
| --- | --- | --- | --- | --- |
| WorkPaper service | `npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door workpaper-service --json` | `door: "workpaper-service"`, `verified: true` | Node can edit a WorkPaper input, recalculate a formula, export JSON, restore it, and verify readback. | MCP discovery, private workbook compatibility, macros, pivots, charts, or Excel UI behavior. |
| Agent MCP evaluator | `npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json` | `door: "agent-mcp"`, `listedResourcesAndPrompts`, `restartReadbackMatchesAfter` | A coding agent or MCP client can discover workbook tools, write a cell, read a formula value, persist state, and restart from disk. | Hosted auth, arbitrary client UX, or full workbook compatibility. |
| Provider-backed formula boundary | `npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario provider-backed --json` | `scenario: "provider-backed"`, `provider-backed-adapter-missing`, `adapterBackedDiagnosticsCleared` | Provider formulas such as `IMPORTRANGE` fail closed until the host supplies an adapter, then verify readback. | Live Google Sheets authorization or remote provider availability. |
| XLSX cache detector | `npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- bilig-evaluate --door xlsx-cache --json` | `door: "xlsx-cache"`, `verified: true` | A saved `.xlsx` can be inspected for stale cached formula values without mutating the file. | In-process business state or JSON WorkPaper persistence. |
| XLSX recalculation | `npm exec --package @bilig/xlsx-formula-recalc@latest -- xlsx-recalc --demo --json` | `recalculationCompleted: true` | An XLSX file boundary can be edited, recalculated, exported, and reimported for readback. | A full Excel clone, macro execution, charts, pivots, or desktop layout fidelity. |
| ExcelJS recalculation | `npx --package @bilig/exceljs-formula-recalc exceljs-recalc --demo --json` | `commandSucceeded: true`, `recalculationCompleted: true`, `expectedValueMatched: true` | An existing ExcelJS workbook can get fresh formula readback after Node edits. | ExcelJS styling/export behavior, desktop Excel parity, or every Excel formula. |
| MCP Inspector | `npx -y @modelcontextprotocol/inspector@latest --cli npm exec --yes --package @bilig/workpaper@latest -- bilig-workpaper-mcp --method tools/list` | tool names such as `read_workpaper_summary`, `set_workpaper_input_cell` | A neutral MCP client can inspect the packaged stdio server before a user adds it to an agent host. | Private workbook persistence unless the file-backed config is used. |
| File-backed MCP server | `npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable` | `set_cell_contents_and_readback`, `export_workpaper_document` | A local agent can use a persistent WorkPaper JSON file and general workbook tools. | Hosted multi-user storage or secret management. |
| Vercel AI SDK `generateText()` | `pnpm --dir examples/headless-workpaper run agent:ai-sdk-generate-text` | `apiShape: "AI SDK generateText -> tool -> execute"` | AI SDK tools can return before/after/restore WorkPaper proof from a `generateText()` loop. | Provider model quality or production prompt behavior. |
| Vercel AI SDK `streamText()` | `pnpm --dir examples/headless-workpaper run agent:ai-sdk-stream-text` | `apiShape: "AI SDK streamText -> tool -> execute"`, `streamChunkTypes` | Streaming tool calls can carry the same WorkPaper proof while the model streams final text. | Browser UI streaming, telemetry retention, or non-deterministic provider output. |
| OpenAI Responses function call | `pnpm --dir examples/headless-workpaper run agent:openai-responses` | `function_call_output`, `verified: true` | OpenAI tool calling can wrap WorkPaper readback as a structured function result. | Hosted remote MCP app review or ChatGPT UI behavior. |
| OpenAI Agents SDK hosted MCP | `pnpm --dir examples/headless-workpaper run agent:openai-agents-sdk-hosted-mcp` | `MCPServerStreamableHttp`, `set_cell_contents_and_readback` | An OpenAI Agents SDK agent can call the hosted Streamable HTTP MCP endpoint. | Private writable workbook state. Use local stdio for that. |
| LangGraph ToolNode | `pnpm --dir examples/headless-workpaper run agent:langgraph-toolnode` | `SpreadsheetAgentProof`, `restoredMatchesAfter` | Graph state can carry the proof object instead of only a scalar formula value. | Whether every graph architecture should store workbook state. |
| Semantic Kernel MCP plugin | `python examples/semantic-kernel-workpaper-mcp/semantic_kernel_workpaper_mcp.py --smoke` | `verified: true`, `pluginBoundary` | A .NET-oriented agent path can keep WorkPaper behind a plugin or MCP boundary. | .NET package publication or C# parity for every TypeScript example. |
| Mastra tool | `pnpm --dir examples/mastra-workpaper-tool run smoke` | `Mastra createTool -> execute -> WorkPaper readback` | The repo-local Mastra example uses a real `createTool()` wrapper around formula readback. | A new Mastra catalog submission, issue, or PR. Do not duplicate that outreach. |

## Selection Rules

Use `agent-mcp` first when the caller is an agent, MCP client, tool host, or
integration reviewer. It proves discovery, write/readback, resources, prompts,
and restart state in one command.

Use the Vercel AI SDK, OpenAI, LangGraph, Semantic Kernel, or Mastra examples
only after the generic evaluator passes. Those examples prove host fit, not a
stronger workbook runtime.

Use the XLSX and ExcelJS paths when a saved file or ExcelJS object is already
the contract. Do not force a WorkPaper model when the job is mostly workbook
formatting, image embedding, or file metadata.

## Limits

Bilig is not a desktop Excel replacement. Keep Excel, LibreOffice, Microsoft
Graph, or a spreadsheet-specific oracle in the loop for macros, pivots, charts,
external links, unsupported formulas, locale-specific Excel behavior, or exact
manual UI workflows.

## Related

- [Evaluate Bilig as an agent MCP workbook tool](eval-agent-mcp.md)
- [MCP spreadsheet formula server for coding agents](mcp-spreadsheet-formula-server-for-coding-agents.md)
- [Vercel AI SDK spreadsheet tool: generateText and streamText with formula readback](vercel-ai-sdk-spreadsheet-tool-formula-readback.md)
- [ExcelJS formula result not updating after Node edits](exceljs-formula-result-not-updating-after-node-edits.md)
- [Workbook tools for agent frameworks](agent-framework-workbook-tools.md)
- [Compatibility limits](where-bilig-is-not-excel-compatible-yet.md)
