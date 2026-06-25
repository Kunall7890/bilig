---
title: Agno WorkPaper MCP tools
published: true
description: Use Agno MCPTools to launch Bilig's WorkPaper MCP server, edit workbook inputs, and verify formula readback without spreadsheet UI automation.
tags: agno, agents, mcp, workbook-api, spreadsheet-automation
canonical_url: https://proompteng.github.io/bilig/agno-workpaper-mcp.html
image: /assets/github-social-preview.png
---

# Agno WorkPaper MCP tools

Use this when an Agno agent needs spreadsheet formulas through MCP. The workbook
should be a tool boundary: read a range, write one input, read the dependent
formula output, persist WorkPaper JSON, and return the proof.

Official Agno references:

- <https://docs.agno.com/tools/mcp>
- <https://docs.agno.com/agents/introduction>
- <https://github.com/agno-agi/agno>

## Run The Smoke Test

The runnable source lives in:

```text
examples/agno-workpaper-mcp
```

Run it from the repo root:

```sh
uv run --python 3.12 --with agno --with mcp --with openai \
  python examples/agno-workpaper-mcp/agno_workpaper_mcp.py \
  --output .tmp/agno-workpaper-proof.json
```

Expected top-level result:

```json
{
  "framework": "agno",
  "toolkit": "MCPTools",
  "packageSpec": "@bilig/workpaper@latest",
  "verified": true,
  "beforeExpectedArr": 60000,
  "afterExpectedArr": 96000,
  "restoredExpectedArr": 96000
}
```

The command starts:

```sh
npm exec --yes --package @bilig/workpaper@latest -- \
  bilig-workpaper-mcp \
  --workpaper /tmp/pricing.workpaper.json \
  --init-demo-workpaper \
  --writable
```

Then Agno imports the MCP tools through `MCPTools` and calls:

1. `list_sheets`
2. `read_range`
3. `set_cell_contents_and_readback`
4. `export_workpaper_document`

The write changes `Inputs!B3` to `0.4` and reads `Summary!A1:B4`. The dependent
`Summary!B3` value changes from `60000` to `96000`, and the restored WorkPaper
readback stays `96000`.

## Minimal Code Shape

```python
from agno.tools.mcp import MCPTools
from mcp import StdioServerParameters

server_params = StdioServerParameters(
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

async with MCPTools(server_params=server_params, timeout_seconds=30) as tools:
    await tools.initialize()
    functions = tools.get_functions()
    result = await functions["set_cell_contents_and_readback"].entrypoint(
        sheetName="Inputs",
        address="B3",
        value=0.4,
        readbackRange="Summary!A1:B4",
    )
```

The checked-in example parses the JSON result, verifies
`restoredReadbackMatchesAfter`, confirms the persisted WorkPaper file, and fails
closed if the calculated value does not match.

## Optional Agent Summary

The proof does not need an LLM key. If you want an Agno `Agent` to summarize the
proof, set `OPENAI_API_KEY` and add `--agent`:

```sh
uv run --python 3.12 --with agno --with mcp --with openai \
  python examples/agno-workpaper-mcp/agno_workpaper_mcp.py --agent
```

Keep the workbook validation outside the model response. The agent may explain
the proof, but `verified: true` should come from the MCP readback checks.

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
