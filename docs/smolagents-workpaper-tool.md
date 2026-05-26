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
deliberately narrow: one tool runs Bilig's WorkPaper proof command and returns
structured evidence that the formula readback changed and survived JSON
restore.

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

- `smolagents_workpaper_tool.py` for the `Tool` subclass
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
  "packageSpec": "@bilig/workpaper@latest",
  "verified": true
}
```

## Tool Shape

The checked-in tool is a normal smolagents tool:

```python
from smolagents import Tool


class BiligWorkPaperFormulaProofTool(Tool):
    name = "verify_workpaper_formula_readback"
    output_type = "object"

    def forward(self, package_spec: str):
        ...
```

The tool runs:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-challenge --json
```

That command edits `Inputs!B2`, recalculates `Summary!B2`, serializes the
WorkPaper document, restores it, and returns `verified: true` with explicit
checks.

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
from smolagents_workpaper_tool import BiligWorkPaperFormulaProofTool

model = InferenceClientModel()
agent = CodeAgent(
    tools=[BiligWorkPaperFormulaProofTool()],
    model=model,
)
```

Ask the agent to call `verify_workpaper_formula_readback` before relying on a
formula-backed decision. The tool returns the edited cell, dependent cell,
before/after values, restore result, serialized document size, limitations, and
links for starring, watching releases, or filing an adoption blocker.

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
