---
title: Microsoft Agent Framework WorkPaper MCP tools
published: true
description: Use Microsoft Agent Framework MCP tools with Bilig WorkPaper so Python or .NET agents can edit workbook inputs, recalculate formulas, and verify readback without spreadsheet UI automation.
tags: microsoft-agent-framework, agents, mcp, spreadsheet automation, workpaper
canonical_url: https://proompteng.github.io/bilig/microsoft-agent-framework-workpaper-mcp.html
image: /assets/github-social-preview.png
---

# Microsoft Agent Framework WorkPaper MCP Tools

Microsoft Agent Framework can connect agents to local and remote MCP servers.
That makes it a direct fit for Bilig WorkPaper when an agent needs workbook
formula logic but should not drive Excel, LibreOffice, Google Sheets, or a
browser grid.

Bilig owns the WorkPaper state, formula recalculation, JSON persistence, and
read-after-write proof. Agent Framework owns the Python or .NET agent host that
loads MCP tools.

Official references:

- <https://learn.microsoft.com/en-us/agent-framework/>
- <https://learn.microsoft.com/en-us/agent-framework/agents/tools/local-mcp-tools>
- <https://github.com/microsoft/agent-framework>
- <https://github.com/microsoft/autogen>

## Run The No-Key Smoke

Use the package-owned evaluator first:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

Expected proof shape:

```json
{
  "door": "agent-mcp",
  "verified": true,
  "evidence": {
    "editedCell": "Inputs!B3",
    "dependentCell": "Summary!B3",
    "before": 60000,
    "after": 96000,
    "afterRestore": 96000,
    "afterRestart": 96000
  }
}
```

After editing this recipe, run the static guard:

```sh
python examples/microsoft-agent-framework-workpaper-mcp/scripts/check-microsoft-agent-framework-recipe.py
```

The guard is intentionally static. It protects the public recipe text without
requiring a particular preview or stable Agent Framework package API to be
installed in CI.

## Local Python MCP Shape

Agent Framework's Python docs show `MCPStdioTool` for local stdio MCP servers.
Use it to launch Bilig's file-backed WorkPaper MCP server as an argument array:

```python
from agent_framework import Agent, MCPStdioTool

bilig_tools = MCPStdioTool(
    name="Bilig WorkPaper",
    command="npm",
    args=[
        "exec",
        "--yes",
        "--package",
        "@bilig/workpaper@latest",
        "--",
        "bilig-workpaper-mcp",
        "--workpaper",
        "./pricing.workpaper.json",
        "--init-demo-workpaper",
        "--writable",
    ],
)

async with Agent(
    client=client,
    name="PricingAgent",
    instructions="Edit workbook inputs only through Bilig tools and verify formula readback.",
    tools=bilig_tools,
) as agent:
    result = await agent.run(
        "Set Inputs!B3 to 0.4 and report Summary!B3 after recalculation.",
    )
```

The useful proof is the tool transcript, not the final model sentence. The
agent must call `set_cell_contents_and_readback` or call `set_cell_contents`,
then `read_cell` or `read_range`, then `export_workpaper_document`.

## Hosted Streamable HTTP Shape

Agent Framework's Python docs also show `MCPStreamableHTTPTool` for HTTP MCP
servers. Use Bilig's hosted no-key endpoint for connector smoke tests:

```python
from agent_framework import MCPStreamableHTTPTool

bilig_tools = MCPStreamableHTTPTool(
    name="Bilig hosted WorkPaper MCP",
    url="https://bilig.proompteng.ai/mcp",
)
```

The hosted endpoint is request-local. It is useful for MCP discovery and demo
readback proof, but it does not persist private workbook files. Use the stdio
command when an agent must write `./pricing.workpaper.json`.

## .NET Boundary

For .NET hosts, follow the Agent Framework MCP guidance with the official MCP
C# SDK, convert discovered MCP tools into `AITool` values, and keep the same
Bilig proof contract:

1. list tools from the Bilig MCP server;
2. edit only the requested input cell;
3. read the dependent formula cell after recalculation;
4. export the WorkPaper document;
5. reopen or restore when a file boundary matters.

Do not accept an agent answer that only says the cell was updated.

## Proof Loop

A valid Microsoft Agent Framework integration should prove this loop:

- discover `set_cell_contents_and_readback`;
- read `Inputs!B3` and `Summary!B3`;
- write `Inputs!B3` to `0.4`;
- read `Summary!B3` changing from `60000` to `96000`;
- export the WorkPaper JSON document;
- restart or restore and confirm `Summary!B3` is still `96000`.

Return a compact object with `editedCell`, `before`, `after`, `afterRestore`,
`persistedDocumentBytes`, `verified`, and `limitations`.

## Boundary

This recipe is owned by Bilig. No upstream Microsoft Agent Framework PR or issue was opened for this page.

This proves the MCP boundary and the WorkPaper readback contract. It does not
claim full desktop Excel compatibility, Office macro execution, pivot refresh,
external link refresh, private workbook storage on the hosted endpoint, or that
every Agent Framework preview package version has the exact same import names.
