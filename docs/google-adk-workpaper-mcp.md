---
title: Google ADK WorkPaper MCP tools
published: true
description: Use Google ADK McpToolset to launch Bilig's WorkPaper MCP server, edit workbook inputs, and verify formula readback without spreadsheet UI automation.
tags: google-adk, ai-agents, mcp, workbook-api, spreadsheet-automation
canonical_url: https://proompteng.github.io/bilig/google-adk-workpaper-mcp.html
image: /assets/github-social-preview.png
---

# Google ADK WorkPaper MCP tools

Use this when a Google Agent Development Kit app needs spreadsheet formulas
through MCP. ADK should own agent orchestration; Bilig should own the workbook
truth: read a range, write one input, read the dependent formula output, persist
WorkPaper JSON, and return the proof.

Official Google ADK references:

- <https://adk.dev/tools-custom/mcp-tools/>
- <https://adk.dev/integrations/>
- <https://github.com/google/adk-python>

## Run The Smoke Test

The runnable source lives in:

```text
examples/google-adk-workpaper-mcp
```

Run it from the repo root:

```sh
uv run --python 3.12 --with google-adk --with mcp \
  python examples/google-adk-workpaper-mcp/google_adk_workpaper_mcp.py \
  --output .tmp/google-adk-workpaper-proof.json
```

Expected top-level result:

```json
{
  "framework": "google-adk",
  "toolset": "McpToolset",
  "packageSpec": "@bilig/workpaper@latest",
  "verified": true,
  "beforeExpectedArr": 60000,
  "afterExpectedArr": 96000,
  "restoredExpectedArr": 96000
}
```

The command starts Bilig's file-backed MCP server through `npm exec`, then ADK's
`McpToolset` discovers and calls:

1. `read_range`
2. `set_cell_contents_and_readback`
3. `export_workpaper_document`

The write changes `Inputs!B3` to `0.4` and reads `Summary!A1:B4`. The dependent
`Summary!B3` value changes from `60000` to `96000`, and the restored WorkPaper
readback stays `96000`.

## Minimal Agent Shape

ADK's Python MCP integration uses `McpToolset` with `StdioConnectionParams`.
Keep `tool_filter` narrow when the workbook task only needs read/write/export
tools.

```python
from google.adk.agents import LlmAgent
from google.adk.tools.mcp_tool import McpToolset
from google.adk.tools.mcp_tool.mcp_session_manager import StdioConnectionParams
from mcp import StdioServerParameters

root_agent = LlmAgent(
    name="workpaper_pricing_agent",
    model="gemini-2.5-flash",
    instruction=(
        "Use Bilig WorkPaper MCP tools for workbook edits. "
        "Do not claim success until readback and restored state match."
    ),
    tools=[
        McpToolset(
            connection_params=StdioConnectionParams(
                server_params=StdioServerParameters(
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
                ),
                timeout=30,
            ),
            tool_filter=[
                "read_range",
                "set_cell_contents_and_readback",
                "export_workpaper_document",
            ],
        )
    ],
)
```

That is the agent-facing shape. The checked-in smoke test uses the same
`McpToolset` and direct tool calls so the proof runs without a Gemini key.

## First Proof Command

If you only need to decide whether Bilig fits before wiring ADK, run the no-key
agent MCP evaluator first:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

Trust the workflow only when it returns `verified: true` with:

- `Inputs!B3` edited;
- `Summary!B3` changing from `60000` to `96000`;
- `restoredReadbackMatchesAfter: true`;
- persisted WorkPaper document bytes.

## Boundary

This is for formula-backed service or agent workflows where WorkPaper JSON can
represent the workbook state. It is a fit for quote approval, payout checks,
pricing rules, import validation, and forecast gates.

It is not a claim that Bilig replaces desktop Excel for macros, add-ins, pivot
tables, or visual workbook review. For raw `.xlsx` files, start with:

```sh
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- \
  bilig-evaluate --door xlsx-cache --json
```

No upstream Google ADK PR or issue was opened for this guide; it is an owned
Bilig integration surface backed by public ADK MCP docs and a local no-key proof.
