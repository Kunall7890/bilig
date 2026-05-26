---
title: LangGraph.js WorkPaper ToolNode spreadsheet tool
published: true
description: Run a real LangGraph.js ToolNode against @bilig/workpaper and keep formula readback proof in graph state.
tags: langgraph, toolnode, langchain, spreadsheet, workpaper
canonical_url: https://proompteng.github.io/bilig/langgraph-workpaper-toolnode-spreadsheet.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# LangGraph.js WorkPaper ToolNode Spreadsheet Tool

LangGraph.js workflows often route model tool calls through a `ToolNode`. That is
a good place for WorkPaper tools when the graph needs a number it can trust:
read a formula-backed summary, edit one input, recalculate, persist WorkPaper
JSON, restore it, then keep the proof in graph state.

The checked example uses the real `@langchain/langgraph` `ToolNode` with
`AIMessage` tool calls and returned `ToolMessage` state. It does not require an
LLM key because the smoke test supplies deterministic tool calls directly.

## Run the checked graph

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig
cd examples/langgraph-workpaper-tool-state
pnpm install --ignore-workspace --lockfile=false
pnpm run typecheck
pnpm run smoke
```

The smoke builds this graph:

```ts
new StateGraph(MessagesAnnotation)
  .addNode('agent_requests_workpaper_tools', deterministicToolCalls)
  .addNode('tools', new ToolNode(workpaperTools))
  .addEdge(START, 'agent_requests_workpaper_tools')
  .addEdge('agent_requests_workpaper_tools', 'tools')
  .addEdge('tools', END)
```

It returns the graph nodes, tool-message names, the pre-edit summary, and the
verified WorkPaper write proof:

```json
{
  "framework": "langgraphjs-toolnode",
  "graphNodes": ["agent_requests_workpaper_tools", "tools"],
  "toolMessageNames": ["read_workpaper_quote", "set_workpaper_quantity"],
  "proof": {
    "editedCell": "Inputs!B2",
    "before": {
      "total": 1458
    },
    "after": {
      "total": 2187
    },
    "afterRestore": {
      "total": 2187
    },
    "verified": true
  }
}
```

## ToolNode shape

```ts
import { AIMessage } from '@langchain/core/messages'
import { tool } from '@langchain/core/tools'
import { StateGraph, MessagesAnnotation, START, END } from '@langchain/langgraph'
import { ToolNode } from '@langchain/langgraph/prebuilt'

const tools = [
  tool(readQuoteSummary, {
    name: 'read_workpaper_quote',
    schema: z.object({}),
  }),
  tool(setQuantityAndProve, {
    name: 'set_workpaper_quantity',
    schema: z.object({ quantity: z.number().finite().positive() }),
  }),
]

const graph = new StateGraph(MessagesAnnotation)
  .addNode('agent_requests_workpaper_tools', () => ({
    messages: [
      new AIMessage({
        content: '',
        tool_calls: [
          { id: 'call_read_quote', name: 'read_workpaper_quote', args: {}, type: 'tool_call' },
          { id: 'call_set_quantity', name: 'set_workpaper_quantity', args: { quantity: 18 }, type: 'tool_call' },
        ],
      }),
    ],
  }))
  .addNode('tools', new ToolNode(tools))
  .addEdge(START, 'agent_requests_workpaper_tools')
  .addEdge('agent_requests_workpaper_tools', 'tools')
  .addEdge('tools', END)
  .compile()
```

## What to copy

- Use separate read and write tools so graph state stays easy to inspect.
- Return exact `ToolMessage` content with the edited cell and formula readback.
- Keep persistence and restore verification in the tool result when the graph
  can resume later from a checkpoint.
- Keep the compatibility caveat visible: this is a WorkPaper API, not full
  desktop Excel UI automation.

Official LangGraph.js references:

- <https://docs.langchain.com/oss/javascript/langchain/tools>
- <https://langchain-ai.github.io/langgraphjs/reference/classes/langgraph.prebuilt.ToolNode.html>
- <https://langchain-ai.github.io/langgraphjs/reference/functions/langgraph.prebuilt.toolsCondition.html>

Runnable source:
[`examples/langgraph-workpaper-tool-state`](../examples/langgraph-workpaper-tool-state).
