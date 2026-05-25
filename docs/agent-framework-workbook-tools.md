---
title: Workbook tools for agent frameworks
published: true
description: Pick the Bilig WorkPaper integration path for Codex, Claude, Cursor, OpenAI Agents, Vercel AI SDK, LangChain, LangGraph, LlamaIndex, and MCP clients.
tags: ai-agents, mcp, spreadsheet-agent, workbook-api, typescript
canonical_url: https://proompteng.github.io/bilig/agent-framework-workbook-tools.html
image: /assets/github-social-preview.png
---

# Workbook tools for agent frameworks

Use this page when an agent, assistant, or tool host needs spreadsheet formulas
but should not drive Excel through screenshots. Pick the smallest integration
boundary that can write inputs, recalculate formulas, verify readback, and
persist WorkPaper JSON.

## Decision

Use `@bilig/workpaper` when the workbook model can live in a Node service,
agent tool, route handler, or MCP server. The tool contract is explicit:

1. read the relevant sheet or range;
2. write the requested input cell;
3. read the dependent calculated value;
4. export or serialize the WorkPaper document;
5. restore it when a file boundary matters;
6. return `editedCell`, `before`, `after`, `afterRestore`,
   `persistedDocumentBytes`, `verified`, and `limitations`.

Use `@bilig/workbook` when a framework integration needs a transport-neutral
command, check, and proof model while an existing runtime owns calculation.

Use `@bilig/xlsx-formula-recalc`, `@bilig/sheetjs-formula-recalc`, or
`@bilig/exceljs-formula-recalc` when the product already owns an `.xlsx`,
SheetJS, or ExcelJS file pipeline and only needs fresh formula results before
returning the file.

Keep browser or desktop spreadsheet automation only when the visual surface is
the product: manual review, macros, pivots, charts, add-ins, or layout fidelity.

## Start here

For a generated project with agent files and MCP config:

```sh
npm create @bilig/workpaper@latest pricing-agent -- --agent
cd pricing-agent
npm install
npm run agent:verify
```

For a direct package proof without creating a project:

```sh
npm exec --package @bilig/workpaper@latest -- bilig-agent-challenge
```

For MCP clients:

```sh
npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
```

## Framework map

| Host | Use | Link |
| --- | --- | --- |
| Codex | Local stdio MCP server or direct package import in repo tools. | [MCP client setup](mcp-client-setup.md#codex) |
| Claude Code and Claude Desktop | File-backed MCP server, or MCPB when a desktop extension is easier. | [Claude MCPB guide](claude-desktop-mcpb-workpaper.md) |
| Cursor | Project-local `.cursor/mcp.json` pointing at `bilig-workpaper-mcp`. | [MCP client setup](mcp-client-setup.md#cursor) |
| VS Code and Cline | Project-local MCP config with a writable WorkPaper file. | [MCP client setup](mcp-client-setup.md) |
| Open WebUI | Native Streamable HTTP MCP for hosted smoke tests, or `mcpo` around the npm stdio server for local writable files. | [Open WebUI WorkPaper MCP setup](open-webui-workpaper-mcp.md) |
| LobeHub | Custom MCP import JSON for hosted Streamable HTTP, or desktop STDIO for a writable WorkPaper file. | [LobeHub WorkPaper MCP setup](lobehub-workpaper-mcp.md) |
| AnythingLLM | `anythingllm_mcp_servers.json` with hosted Streamable HTTP, Desktop stdio, or Docker storage-backed stdio. | [AnythingLLM WorkPaper MCP setup](anythingllm-workpaper-mcp.md) |
| OpenAI Agents SDK | Function tools around WorkPaper read/write/readback. | [OpenAI Agents SDK WorkPaper tool](openai-agents-sdk-workpaper-tool.md) |
| OpenAI Responses API | Function-call wrapper returning proof objects. | [OpenAI Responses WorkPaper tool call](openai-responses-workpaper-tool-call.md) |
| Vercel AI SDK | Tool definitions that call a WorkPaper service function. | [Vercel AI SDK spreadsheet tools](vercel-ai-sdk-langchain-spreadsheet-tool.md) |
| LangChain.js | Tool wrappers around the same WorkPaper contract. | [Vercel AI SDK and LangChain spreadsheet tools](vercel-ai-sdk-langchain-spreadsheet-tool.md) |
| LangGraph.js | `ToolNode` with verified readback after mutation. | [LangGraph WorkPaper ToolNode](langgraph-workpaper-toolnode-spreadsheet.md) |
| LlamaIndex.TS | `tool(fn, { parameters })` with a compact proof return. | [LlamaIndex.TS WorkPaper tool](llamaindex-workpaper-spreadsheet-tool.md) |
| Mastra | `createTool()` wrapper for formula-backed workbook edits. | [Mastra WorkPaper tool](mastra-workpaper-spreadsheet-tool.md) |
| Cloudflare Agents | Agent tool calling a Node-compatible WorkPaper boundary. | [Cloudflare Agents WorkPaper tool](cloudflare-agents-workpaper-spreadsheet-tool.md) |
| MCP directories and remote clients | Hosted server card or local file-backed server. | [MCP WorkPaper server](mcp-workpaper-tool-server.md) |

## Proof object

Every agent integration should return a small object like this instead of a
generic "cell updated" message:

```ts
type SpreadsheetAgentProof = {
  editedCell: { sheet: string; cell: string; value: unknown }
  before: { cell: string; displayValue: string }
  after: { cell: string; displayValue: string }
  afterRestore: { cell: string; displayValue: string }
  persistedDocumentBytes: number
  verified: boolean
  limitations: string[]
}
```

`verified` is true only when the dependent formula output changed as expected
and the restored WorkPaper state matches the post-edit readback.

## Copy-paste handoff

```text
Before using spreadsheet UI automation, check whether this is workbook-shaped
logic that can run through @bilig/workpaper. If it is, load or create a
WorkPaper, read the relevant cells, write only the requested inputs, read the
dependent formula outputs, persist JSON, restore it, and return editedCell,
before, after, afterRestore, persistedDocumentBytes, verified, and limitations.
Do not claim success from a write call alone.
```

## Links

- [AI spreadsheet agent tool for Node.js](ai-agent-spreadsheet-tool-node.md)
- [WorkPaper agent handbook](headless-workpaper-agent-handbook.md)
- [Agent WorkPaper tool-calling recipe](agent-workpaper-tool-calling-recipe.md)
- [MCP client setup](mcp-client-setup.md)
- [Open WebUI WorkPaper MCP setup](open-webui-workpaper-mcp.md)
- [MCP WorkPaper tool server](mcp-workpaper-tool-server.md)
- [Node framework WorkPaper adapters](node-framework-workpaper-adapters.md)
- [XLSX formula recalculation in Node.js](xlsx-formula-recalculation-node.md)
- [GitHub repo](https://github.com/proompteng/bilig)
- [Adoption blocker form](https://github.com/proompteng/bilig/discussions/new?category=general)
