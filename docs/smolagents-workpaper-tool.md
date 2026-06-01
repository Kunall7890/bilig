---
title: smolagents WorkPaper Tool
published: true
description: Give Hugging Face smolagents a narrow Bilig WorkPaper tool that edits inputs, recalculates formulas, and returns verified readback without spreadsheet UI automation.
tags: smolagents, hugging-face, ai-agents, workbook-api, spreadsheet-automation
canonical_url: https://proompteng.github.io/bilig/smolagents-workpaper-tool.html
image: /assets/github-social-preview.png
---

# smolagents WorkPaper Tool

Use this when a Hugging Face `smolagents` agent needs spreadsheet-style
business logic, but the spreadsheet work should be a small verified tool call
instead of browser automation.

`smolagents` tools are Python classes with a name, description, input schema,
output type, and a `forward()` method. This example keeps the tool surface
small: one tool runs Bilig's local agent-MCP evaluator, and one tool calls the
public Hugging Face Space fixture. Both return structured readback data an agent
can inspect before it trusts a formula-backed decision.

Official smolagents references:

- <https://huggingface.co/docs/smolagents/main/en/tutorials/tools>
- <https://huggingface.co/docs/smolagents/main/en/guided_tour>
- <https://github.com/huggingface/smolagents>

## Example Artifact

The runnable source lives in:

```text
examples/smolagents-workpaper-tool
```

It contains:

- `smolagents_workpaper_tool.py` for the two `Tool` subclasses
- `scripts/check-smolagents-recipe.py` for a static recipe guard
- `README.md` with the no-key smoke command

Run the proof locally:

```sh
cd examples/smolagents-workpaper-tool
uv run --python 3.12 --with smolagents \
  python smolagents_workpaper_tool.py --output .tmp/smolagents-workpaper-proof.json
```

Expected top-level result:

```json
{
  "framework": "smolagents",
  "toolName": "verify_workpaper_formula_readback",
  "door": "agent-mcp",
  "packageSpec": "@bilig/workpaper@latest",
  "verified": true
}
```

Call the live Hugging Face Space fixture:

```sh
cd examples/smolagents-workpaper-tool
uv run --python 3.12 --with smolagents \
  python smolagents_workpaper_tool.py --mode space --win-rate 0.4 \
  --output .tmp/smolagents-workpaper-space.json
```

Expected top-level result:

```json
{
  "framework": "smolagents",
  "toolName": "read_workpaper_space_formula",
  "space": "gregkonush/bilig-workpaper-mcp-readback",
  "verified": true
}
```

## Tool Shape

The checked-in classes are normal smolagents tools:

```python
from smolagents import Tool


class BiligWorkPaperFormulaProofTool(Tool):
    name = "verify_workpaper_formula_readback"
    output_type = "object"

    def forward(self, package_spec: str):
        ...


class BiligWorkPaperSpaceReadbackTool(Tool):
    name = "read_workpaper_space_formula"
    output_type = "object"

    def forward(self, win_rate: float = 0.4):
        ...
```

The local tool runs:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

That command edits `Inputs!B3`, recalculates `Summary!B3`, serializes the
WorkPaper document, restores it, restarts the file-backed MCP server, and
returns `verified: true` with explicit checks. The Space tool calls:

```text
https://gregkonush-bilig-workpaper-mcp-readback.hf.space/gradio_api/call/v2/prove_workpaper_readback
```

For `win_rate: 0.4`, the Space reads `Summary!B3` back as `96000`.

## Why This Fits smolagents

`smolagents` is code-first. The agent can call a compact workbook proof tool,
inspect the returned object, and decide what to do next from Python. Bilig keeps
spreadsheet state behind a deterministic API:

- exact sheet and cell addresses;
- read-after-write formula proof;
- WorkPaper JSON persistence;
- restore proof before trusting the result;
- no screenshot or desktop spreadsheet session.

Use this for formula-backed quote approval, payout rules, budget gates, import
validation, and forecast checks where a spreadsheet is the reviewable business
logic but the runtime belongs in code.

## Agent Wiring

After the smoke proof passes, wire the tool into a `CodeAgent`:

```python
from smolagents import CodeAgent, InferenceClientModel
from smolagents_workpaper_tool import (
    BiligWorkPaperFormulaProofTool,
    BiligWorkPaperSpaceReadbackTool,
)

model = InferenceClientModel()
agent = CodeAgent(
    tools=[BiligWorkPaperFormulaProofTool(), BiligWorkPaperSpaceReadbackTool()],
    model=model,
)
```

Ask the agent to call `verify_workpaper_formula_readback` for local npm/MCP
verification or `read_workpaper_space_formula` for the hosted fixture before
relying on a formula-backed decision. The tools return the edited cell,
before/after values, restore result, serialized document size, limitations, and
checks.

## Boundary

This proof uses `@bilig/workpaper@latest` from npm and verifies the WorkPaper
write/recalc/read/persist loop. It does not claim full desktop Excel
compatibility, Office macro support, or mutation of a private workbook file.

For private persisted workbook state, use Bilig's file-backed MCP server:

```sh
npm exec --package @bilig/workpaper@latest -- \
  bilig-workpaper-mcp \
  --workpaper ./pricing.workpaper.json \
  --init-demo-workpaper \
  --writable
```

For `.xlsx` files, use the XLSX-specific evaluator:

```sh
npm exec --package @bilig/xlsx-formula-recalc@latest -- xlsx-recalc --demo --json
```
