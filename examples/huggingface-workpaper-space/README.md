---
title: Bilig WorkPaper MCP Readback
sdk: docker
app_port: 7860
license: mit
tags:
  - mcp-server
  - gradio
  - agents
  - ai-agents
  - spreadsheet
  - formulas
  - xlsx
  - nodejs
  - workpaper
---

# Bilig WorkPaper MCP Readback

This is a Docker Space template for one Bilig WorkPaper readback tool exposed
through Gradio and MCP.

The tool edits `Inputs!B3`, recalculates dependent formulas, serializes the
WorkPaper JSON, restores it, and returns the exact values it read back. It is
meant for agents and evaluators that need a no-key workbook API check before
they wire Bilig into a private workflow.

The template pins `@bilig/workpaper@0.164.8`, matching the current published
Bilig WorkPaper package line.

## Local Check

```sh
npm install --omit=dev --package-lock=false
uv run --python 3.12 --with 'gradio[mcp]>=6.0,<7' python app.py --check
```

Expected result:

```json
{
  "packageVersion": "0.164.8",
  "verified": true
}
```

## MCP Boundary

A deployed public Space should stay a small readback fixture. It should not
store user files or claim full desktop Excel compatibility. For private
persisted workbook state, run the file-backed WorkPaper MCP server:

```sh
npm exec --package @bilig/workpaper@latest -- \
  bilig-workpaper-mcp \
  --workpaper ./pricing.workpaper.json \
  --init-demo-workpaper \
  --writable
```
