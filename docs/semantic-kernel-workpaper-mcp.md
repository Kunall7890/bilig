---
title: Semantic Kernel WorkPaper MCP plugin
published: true
description: Import Bilig WorkPaper MCP tools into Microsoft Semantic Kernel so Python agents can edit workbook inputs, recalculate formulas, and verify readback without spreadsheet UI automation.
tags: semantic-kernel, microsoft-agent-framework, mcp, spreadsheet, workpaper
canonical_url: https://proompteng.github.io/bilig/semantic-kernel-workpaper-mcp.html
image: /assets/github-social-preview.png
---

# Semantic Kernel WorkPaper MCP Plugin

Microsoft Semantic Kernel can import tools from MCP servers. That makes it a
good host for Bilig when a Python agent needs formula-backed spreadsheet logic
but should not drive Excel, LibreOffice, Google Sheets, or a browser grid.

Semantic Kernel owns the MCP plugin host. Bilig owns the WorkPaper file,
formula recalculation, JSON persistence, and read-after-write proof.

Official Semantic Kernel references:

- <https://learn.microsoft.com/en-us/semantic-kernel/concepts/plugins/adding-mcp-plugins>
- <https://devblogs.microsoft.com/agent-framework/semantic-kernel-adds-model-context-protocol-mcp-support-for-python/>

## Run The No-Key Smoke

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig
uv run --python 3.12 --with 'semantic-kernel[mcp]' \
  python examples/semantic-kernel-workpaper-mcp/semantic_kernel_workpaper_mcp.py \
  --workpaper .tmp/pricing.workpaper.json \
  --output .tmp/semantic-kernel-workpaper-proof.json
```

The command starts the published `@bilig/workpaper@latest` MCP server. It does
not require an OpenAI, Microsoft, Excel, LibreOffice, or Google Sheets key.

Expected top-level output:

```json
{
  "framework": "semantic-kernel-mcp",
  "pluginName": "BiligWorkPaper",
  "packageSpec": "@bilig/workpaper@latest",
  "dependentCell": "Summary!B3",
  "beforeExpectedArr": 60000,
  "afterExpectedArr": 96000,
  "afterRestartExpectedArr": 96000,
  "verified": true
}
```

The full proof includes the imported tool names, edited cell, before/after
input content, dependent formula readback, persisted WorkPaper path, and a
restart readback from the same JSON file.

Use `--local-source` only when you are changing the local TypeScript MCP server
and need to verify unreleased behavior from the repository checkout.

## Plugin Shape

```python
from semantic_kernel import Kernel
from semantic_kernel.connectors.mcp import MCPStdioPlugin

async with MCPStdioPlugin(
    name="BiligWorkPaper",
    description="Bilig WorkPaper spreadsheet formula tools",
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
    load_prompts=False,
    request_timeout=30,
) as plugin:
    kernel = Kernel()
    kernel.add_plugin(plugin)

    proof = await plugin.call_tool(
        "set_cell_contents",
        sheetName="Inputs",
        address="B3",
        value="=0.4",
    )
```

Use formula strings such as `=0.4` when the target MCP host requires each
parameter schema to have a single primitive `type`. Bilig still accepts JSON
number, boolean, and null arguments from MCP clients that support union-typed
tool parameters. After writing, read the dependent formula cell such as
`Summary!B3` and reopen the same WorkPaper JSON file before reporting success.

## Why This Fits Semantic Kernel

Semantic Kernel agents already treat tools as plugin functions. A WorkPaper MCP
server is the right boundary when formula work must be deterministic:

- exact sheet and cell addresses;
- one cell edit per tool call;
- recalculation before returning;
- persisted WorkPaper JSON;
- dependent formula readback before trusting the result;
- restart readback from the persisted WorkPaper JSON;
- no spreadsheet screenshots or cached XLSX formula values.

Use this for quote approval, payout rules, budget gates, import validation,
forecast checks, and any agent workflow where spreadsheet formulas are the
reviewable business logic.

## Boundary

This proves Semantic Kernel can import and call Bilig's MCP tools. It does not
claim full desktop Excel compatibility, Office macro execution, external link
refresh, or mutation of arbitrary private spreadsheets without your own
WorkPaper JSON file.

Runnable source:
[`examples/semantic-kernel-workpaper-mcp`](https://github.com/proompteng/bilig/tree/main/examples/semantic-kernel-workpaper-mcp).
