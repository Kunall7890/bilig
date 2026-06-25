---
title: Browser Use WorkPaper formula tool
published: true
description: Give Browser Use agents a deterministic Bilig WorkPaper tool for spreadsheet formula readback instead of clicking through Excel, Google Sheets, or browser grids.
tags: browser-use, browser automation, spreadsheet-agent, workpaper, mcp
canonical_url: https://proompteng.github.io/bilig/browser-use-workpaper-formula-tool.html
image: /assets/github-social-preview.png
---

# Browser Use WorkPaper Formula Tool

Use this when a Browser Use agent can browse a page, extract quote or forecast
inputs, and fill forms, but the calculation itself is workbook-shaped. Browser
Use should own web navigation. Bilig should own workbook cells, formula
recalculation, JSON persistence, and readback proof.

That split avoids the bad path: asking a browser agent to open Excel, Google
Sheets, LibreOffice, or a browser grid, click cells, infer formulas from pixels,
and report success from a screenshot.

Official Browser Use references:

- <https://docs.browser-use.com/open-source/customize/tools/basics>
- <https://docs.browser-use.com/open-source/customize/tools/add>
- <https://docs.browser-use.com/open-source/customize/tools/response>
- <https://docs.browser-use.com/open-source/customize/agent/all-parameters>
- <https://docs.browser-use.com/open-source/customize/integrations/mcp-server>

## First Proof

Before wiring Browser Use, prove the published Bilig WorkPaper agent door:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

The useful invariants are:

- `schemaVersion: "bilig-evaluator.v1"`
- `door: "agent-mcp"`
- `verified: true`
- `editedCell: "Inputs!B3"`
- `dependentCell: "Summary!B3"`
- `before: 60000`
- `after: 96000`
- `afterRestore: 96000`
- `afterRestart: 96000`
- `restartReadbackMatchesAfter: true`

For a richer workbook, run:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario revenue-plan --json
```

That scenario verifies `SUM`, `SUMIF`, `XLOOKUP`, a `FILTER` spill, a named
expression, JSON persistence, and restart readback through the same MCP door.

## Custom Tool Shape

Browser Use supports custom Python tools through a `Tools()` registry and
`@tools.action(...)`. Return an `ActionResult` so the model receives a compact
proof object rather than a long process log.

```python
import json
import subprocess

from browser_use import ActionResult, Agent, Tools

tools = Tools()


@tools.action(
    description=(
        "Run Bilig WorkPaper formula readback. Use this for spreadsheet-style "
        "pricing, quote, forecast, or validation calculations instead of "
        "clicking Excel, Google Sheets, LibreOffice, or a browser grid."
    )
)
async def run_bilig_workpaper_formula_readback(scenario: str = "agent-mcp") -> ActionResult:
    args = [
        "npm",
        "exec",
        "--yes",
        "--package",
        "@bilig/workpaper@latest",
        "--",
        "bilig-evaluate",
        "--door",
        "agent-mcp",
        "--json",
    ]
    if scenario == "revenue-plan":
        args.extend(["--scenario", "revenue-plan"])

    completed = subprocess.run(args, check=True, capture_output=True, text=True)
    proof = json.loads(completed.stdout)

    return ActionResult(
        extracted_content=json.dumps(
            {
                "door": proof["door"],
                "verified": proof["verified"],
                "editedCell": proof["evidence"]["editedCell"],
                "dependentCell": proof["evidence"]["dependentCell"],
                "before": proof["evidence"].get("before"),
                "after": proof["evidence"].get("after"),
                "afterRestore": proof["evidence"].get("afterRestore"),
                "afterRestart": proof["evidence"].get("afterRestart"),
                "checks": proof["evidence"]["checks"],
            },
            indent=2,
        ),
        long_term_memory="Bilig WorkPaper returned verified formula readback and persisted-state proof.",
    )
```

Wire it into Browser Use:

```python
agent = Agent(
    task=(
        "Read the quote inputs from the current page. If the job needs "
        "spreadsheet formulas, call run_bilig_workpaper_formula_readback and "
        "return editedCell, before, after, afterRestore, afterRestart, checks, "
        "verified, and limitations. Do not use spreadsheet UI screenshots as "
        "formula truth."
    ),
    llm=llm,
    tools=tools,
)
```

For project-local workbook state, use the file-backed Bilig MCP server instead
of the evaluator:

```sh
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper ./.bilig/pricing.workpaper.json --init-demo-workpaper --writable
```

Expected tools:

- `list_sheets`
- `read_range`
- `read_cell`
- `set_cell_contents`
- `set_cell_contents_and_readback`
- `get_cell_display_value`
- `export_workpaper_document`
- `validate_formula`

If your agent host already uses MCP, Browser Use can run its own local MCP
server for browser automation, while Bilig runs a separate WorkPaper MCP server
for workbook state. Keep those responsibilities separate: browser tools inspect
web pages; WorkPaper tools compute formulas.

## Prompt For A Browser Agent

```text
You may use Browser Use to inspect pages and collect quote or forecast inputs.
Do not open Excel, Google Sheets, LibreOffice, or a browser spreadsheet grid for
formula truth. When the workflow becomes workbook-shaped, call the Bilig
WorkPaper formula tool. Return editedCell, dependentCell, before, after,
afterRestore or afterRestart, checks, verified, and limitations. A screenshot
or write-call status is not success.
```

## Duplicate-Safe External Lane

There is already an existing upstream Browser Use integration PR:
<https://github.com/browser-use/browser-use/pull/4909>.

Do not open a second Browser Use PR for the same Bilig example. If maintainers
ask for changes, update that PR in place. This Bilig-owned page exists so
agents and searchers can find the integration path even while upstream review
is pending.

The current upstream PR shape adds:

- `examples/integrations/bilig_workpaper/README.md`
- `examples/integrations/bilig_workpaper/bilig_workpaper_example.py`
- a no-key smoke mode
- proof that `Inputs!B3` edits change expected ARR from `60000` to `96000`
- restore/readback evidence with `verified: true`

## Related

- [Agent WorkPaper handoff](agent-adoption-kit.md)
- [Evaluate Bilig as an agent MCP workbook tool](eval-agent-mcp.md)
- [Why agents need workbook APIs](why-agents-need-workbook-apis.md)
- [Stop driving spreadsheets with screenshots](stop-driving-spreadsheets-with-screenshots.md)
- [MCP WorkPaper tool server](mcp-workpaper-tool-server.md)
- [Agent framework workbook tools](agent-framework-workbook-tools.md)
