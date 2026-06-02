---
title: Pydantic AI WorkPaper MCP tools
published: true
description: Use Pydantic AI MCPToolset to launch Bilig's WorkPaper MCP server, edit workbook inputs, and validate typed formula readback.
tags: pydantic-ai, ai-agents, mcp, workbook-api, spreadsheet-automation
canonical_url: https://proompteng.github.io/bilig/pydantic-ai-workpaper-mcp.html
image: /assets/github-social-preview.png
---

# Pydantic AI WorkPaper MCP tools

Use this when a Pydantic AI app needs spreadsheet formulas through MCP and the
result should be validated as typed Python data before an agent trusts it.

Official Pydantic AI references:

- <https://ai.pydantic.dev/mcp/>
- <https://ai.pydantic.dev/tools/>
- <https://github.com/pydantic/pydantic-ai>

## Run The Smoke Test

The runnable source lives in:

```text
examples/pydantic-ai-workpaper-mcp
```

Run it from the repo root:

```sh
uv run --python 3.12 --with pydantic-ai --with mcp --with fastmcp \
  python examples/pydantic-ai-workpaper-mcp/pydantic_ai_workpaper_mcp.py \
  --output .tmp/pydantic-ai-workpaper-proof.json
```

Expected top-level result:

```json
{
  "framework": "pydantic-ai",
  "toolset": "MCPToolset",
  "packageSpec": "@bilig/workpaper@latest",
  "verified": true,
  "beforeExpectedArr": 60000,
  "afterExpectedArr": 96000,
  "restoredExpectedArr": 96000
}
```

The command starts Bilig's file-backed MCP server through `npm exec`, then
Pydantic AI's `MCPToolset` calls:

1. `list_tools`
2. `read_range`
3. `set_cell_contents_and_readback`
4. `export_workpaper_document`

The write changes `Inputs!B3` to `0.4` and reads `Summary!A1:B4`. The dependent
`Summary!B3` value changes from `60000` to `96000`, and restored WorkPaper
readback stays `96000`.

## Minimal Code Shape

The current Pydantic AI path is `MCPToolset` with a `fastmcp` stdio transport,
not the deprecated `MCPServerStdio` helper.

```python
from fastmcp.client.transports import StdioTransport
from pydantic_ai.mcp import MCPToolset

toolset = MCPToolset(
    StdioTransport(
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
    tool_error_behavior="error",
    read_timeout=30,
)

write = await toolset.direct_call_tool(
    "set_cell_contents_and_readback",
    {
        "sheetName": "Inputs",
        "address": "B3",
        "value": 0.4,
        "readbackRange": "Summary!A1:B4",
    },
)
```

The checked-in example validates the proof with a Pydantic model:

```python
class WorkPaperProof(BaseModel):
    framework: str = "pydantic-ai"
    toolset: str = "MCPToolset"
    before_expected_arr: float = Field(alias="beforeExpectedArr")
    after_expected_arr: float = Field(alias="afterExpectedArr")
    restored_expected_arr: float = Field(alias="restoredExpectedArr")
    verified: bool
```

This keeps the contract explicit: the agent can explain the proof, but the
truth comes from readback checks and typed validation.

## Boundary

This is for formula-backed service or agent workflows where WorkPaper JSON can
represent workbook state. It is a fit for quote approval, payout checks, pricing
rules, import validation, and forecast gates.

It is not a claim that Bilig replaces desktop Excel for macros, add-ins, pivot
tables, or visual workbook review. For raw `.xlsx` files, start with:

```sh
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- \
  bilig-evaluate --door xlsx-cache --json
```
