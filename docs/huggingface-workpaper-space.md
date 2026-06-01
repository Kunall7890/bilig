---
title: Hugging Face WorkPaper MCP Space
description: Run a Bilig WorkPaper formula-readback fixture from a Hugging Face Gradio Space with MCP enabled.
published: true
tags: hugging-face, gradio, mcp, workpaper, spreadsheet-automation
canonical_url: https://proompteng.github.io/bilig/huggingface-workpaper-space.html
---

# Hugging Face WorkPaper MCP Space

Use this when an agent already works with Hugging Face Spaces, Gradio, or MCP
tools and needs a small WorkPaper readback fixture before wiring a private
workbook or service.

Live Space:
<https://huggingface.co/spaces/gregkonush/bilig-workpaper-mcp-readback>

The live Space runs the Docker template from this repo. The Gradio function has
MCP enabled, edits `Inputs!B3`, recalculates dependent formulas, exports
WorkPaper JSON, restores it, and returns the exact values it read back.

Source in this repo:

```text
examples/huggingface-workpaper-space
```

## Run Locally

```sh
cd examples/huggingface-workpaper-space
npm install --omit=dev --package-lock=false
uv run --python 3.12 --with 'gradio[mcp]>=6.0,<7' python app.py --check
```

The check calls the same Node script the Space uses and expects:

```json
{
  "editedCell": "Inputs!B3",
  "verified": true
}
```

## What The Tool Returns

The response includes:

- edited cell: `Inputs!B3`;
- formulas for `Summary!B2:B5`;
- before, after, and restored readback values;
- serialized WorkPaper JSON byte count;
- checks for formula output changes, restored readback, input persistence, and
  formula persistence.

For the default win rate `0.4`, `Summary!B3` reads back as `96000`.

## Deploy Your Own

Create a Hugging Face Docker Space and upload the files from
`examples/huggingface-workpaper-space`. The public Space is a no-key hosted
fixture for agent stacks that want to inspect a workbook API result without
spreadsheet UI. It is a discovery and evaluation artifact, not user-file
storage.

For private files, run the file-backed WorkPaper MCP server:

```sh
npm exec --package @bilig/workpaper@latest -- \
  bilig-workpaper-mcp \
  --workpaper ./pricing.workpaper.json \
  --init-demo-workpaper \
  --writable
```
