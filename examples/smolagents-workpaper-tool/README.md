# smolagents WorkPaper tool

Use this example when a Hugging Face `smolagents` agent needs spreadsheet-style
business logic without opening Excel, LibreOffice, Google Sheets, or a browser
spreadsheet grid.

The recipe exposes two narrow smolagents `Tool` classes:

- run Bilig's local agent-MCP evaluator from npm;
- call the public Hugging Face Space readback fixture;
- edit a workbook input cell;
- recalculate a dependent formula;
- serialize and restore WorkPaper JSON;
- return a compact `verified: true` object that an agent can inspect.

No model key is needed for the smoke test. The script exercises the tool class
directly so the workbook boundary is proven before wiring it into a `CodeAgent`.

## Run the local evaluator

```sh
uv run --python 3.12 --with smolagents \
  python smolagents_workpaper_tool.py --output .tmp/smolagents-workpaper-proof.json
```

Expected top-level output:

```json
{
  "framework": "smolagents",
  "toolName": "verify_workpaper_formula_readback",
  "door": "agent-mcp",
  "packageSpec": "@bilig/workpaper@latest",
  "verified": true
}
```

The full output includes the edited cell, dependent formula cell, before/after
values, restart readback, serialized document size, discovered MCP tools, and
checks.

## Call the live Space

Use this path when a smolagents agent needs a no-key hosted readback fixture:

```sh
uv run --python 3.12 --with smolagents \
  python smolagents_workpaper_tool.py \
  --mode space \
  --win-rate 0.4 \
  --output .tmp/smolagents-workpaper-space.json
```

Expected top-level output:

```json
{
  "framework": "smolagents",
  "toolName": "read_workpaper_space_formula",
  "space": "gregkonush/bilig-workpaper-mcp-readback",
  "verified": true
}
```

For the default win rate, the Space edits `Inputs!B3` and reads `Summary!B3`
back as `96000`.

## Why This Fits smolagents

`smolagents` agents write and execute Python code. That works well for Bilig
because workbook operations should be a small tool surface, not dozens of
spreadsheet UI actions. The agent gets one callable proof function and receives
structured evidence it can branch on:

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

Ask the agent to call `verify_workpaper_formula_readback` before trusting a
formula-backed quote, payout, budget, or import-validation result.

## Boundary

The local smoke test uses `@bilig/workpaper@latest` from npm and verifies the
WorkPaper write/recalc/read/persist loop. The Space smoke test calls the public
hosted fixture. Neither path claims full desktop Excel compatibility or mutates
a user spreadsheet file.

For private workbook state, use Bilig's file-backed MCP server or build a
service route around `@bilig/workpaper` so the agent tool writes your own
WorkPaper JSON document.

Official smolagents references:

- <https://huggingface.co/docs/smolagents/main/en/tutorials/tools>
- <https://huggingface.co/docs/smolagents/main/en/guided_tour>
