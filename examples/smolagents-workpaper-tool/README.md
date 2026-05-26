# smolagents WorkPaper tool

Use this example when a Hugging Face `smolagents` agent needs spreadsheet-style
business logic without opening Excel, LibreOffice, Google Sheets, or a browser
spreadsheet grid.

The recipe exposes one narrow smolagents `Tool`:

- run Bilig's public WorkPaper proof command from npm;
- edit a workbook input cell;
- recalculate a dependent formula;
- serialize and restore WorkPaper JSON;
- return a compact `verified: true` object that an agent can inspect.

No model key is needed for the smoke test. The script exercises the tool class
directly so the workbook boundary is proven before wiring it into a `CodeAgent`.

## Run

```sh
uv run --python 3.12 --with smolagents \
  python smolagents_workpaper_tool.py --output .tmp/smolagents-workpaper-proof.json
```

Expected top-level output:

```json
{
  "framework": "smolagents",
  "toolName": "verify_workpaper_formula_readback",
  "packageSpec": "@bilig/workpaper@latest",
  "verified": true
}
```

The full proof includes the edited cell, dependent formula cell, before/after
values, restore check, serialized document size, and adoption links.

## Why This Fits smolagents

`smolagents` agents write and execute Python code. That works well for Bilig
because workbook operations should be a small tool surface, not dozens of
spreadsheet UI actions. The agent gets one callable proof function and receives
structured evidence it can branch on:

```python
from smolagents import CodeAgent, InferenceClientModel
from smolagents_workpaper_tool import BiligWorkPaperFormulaProofTool

model = InferenceClientModel()
agent = CodeAgent(
    tools=[BiligWorkPaperFormulaProofTool()],
    model=model,
)
```

Ask the agent to call `verify_workpaper_formula_readback` before trusting a
formula-backed quote, payout, budget, or import-validation result.

## Boundary

This smoke test uses `@bilig/workpaper@latest` from npm and proves the
WorkPaper write/recalc/read/persist loop. It does not claim full desktop Excel
compatibility or mutate a user spreadsheet file.

For private workbook state, use Bilig's file-backed MCP server or build a
service route around `@bilig/workpaper` so the agent tool writes your own
WorkPaper JSON document.

Official smolagents references:

- <https://huggingface.co/docs/smolagents/main/en/tutorials/tools>
- <https://huggingface.co/docs/smolagents/main/en/guided_tour>
