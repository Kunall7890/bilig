---
title: ChatGPT Apps WorkPaper MCP
published: true
description: Add Bilig as a ChatGPT Developer Mode remote MCP app for WorkPaper workbook readback proof.
tags: chatgpt apps, mcp, spreadsheet agent, workbook, workpaper
canonical_url: https://proompteng.github.io/bilig/chatgpt-apps-workpaper-mcp.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# ChatGPT Apps WorkPaper MCP

Use this path when a ChatGPT conversation needs workbook tools without opening
Excel, Google Sheets, LibreOffice, or a browser grid. Bilig exposes a public
Streamable HTTP MCP endpoint that ChatGPT Developer Mode can add as a remote MCP
app:

```text
https://bilig.proompteng.ai/mcp
```

This is a data/tool-only remote MCP app path. It does not claim a custom Apps
SDK component UI yet. The hosted endpoint is for no-key discovery and stateless
WorkPaper proof; use the local file-backed stdio server for private workbook
state that must persist across calls.

## What ChatGPT Gets

The hosted endpoint advertises the same eight WorkPaper MCP tools as the local
server:

- `list_sheets`
- `read_range`
- `read_cell`
- `set_cell_contents`
- `set_cell_contents_and_readback`
- `get_cell_display_value`
- `export_workpaper_document`
- `validate_formula`

It also publishes a server card at:

```text
https://bilig.proompteng.ai/.well-known/mcp/server-card.json
```

OpenAI documents remote MCP servers for ChatGPT apps and API integrations at
<https://developers.openai.com/api/docs/mcp>. The Apps SDK connection guide
covers Developer Mode setup, remote HTTPS MCP URLs, metadata refresh, and write
tool confirmations:
<https://developers.openai.com/apps-sdk/deploy/connect-chatgpt>.

## Add It In ChatGPT

1. In ChatGPT, enable Developer Mode from Settings -> Apps & Connectors -> Advanced settings.
2. Create a new app or connector from a remote MCP server.
3. Use a name like `Bilig WorkPaper`.
4. Set the connector URL to `https://bilig.proompteng.ai/mcp`.
5. Use no authentication for the hosted demo endpoint.
6. Create the app, then confirm ChatGPT lists the eight WorkPaper tools above.
7. In a new chat, attach the Bilig WorkPaper app from the composer tool picker.

When the tool list or descriptions change, refresh the app metadata from the
ChatGPT app settings before testing again.

## Copy-Paste Prompt

```text
Use the Bilig WorkPaper app's set_cell_contents_and_readback tool.
Set Inputs!B3 to =0.4, read Summary!A1:B4, and report Summary!B3 before,
after, and after restored readback. Do not use browsing, screenshots, Excel,
Google Sheets, LibreOffice, or a spreadsheet UI. If the app is not attached,
say that first instead of guessing.
```

The useful result is not "tool called". The useful result is computed cell
evidence:

```json
{
  "tool": "set_cell_contents_and_readback",
  "editedCell": "Inputs!B3",
  "readbackRange": "Summary!A1:B4",
  "before": { "expectedArr": 60000 },
  "after": { "expectedArr": 96000 },
  "restored": { "expectedArr": 96000 },
  "persistence": { "persisted": false },
  "restoredMatchesAfter": true
}
```

`persistence.persisted: false` is expected on the hosted endpoint because every
request gets a fresh demo WorkPaper. The endpoint still exports serialized
WorkPaper bytes and restores them inside the same tool proof. For real project
state, run the stdio server against a file:

```sh
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
```

## Terminal Checks

Check that ChatGPT browser origins can preflight the MCP endpoint:

```sh
curl -i -X OPTIONS https://bilig.proompteng.ai/mcp \
  -H 'Origin: https://chatgpt.com' \
  -H 'Access-Control-Request-Method: POST' \
  -H 'Access-Control-Request-Headers: accept, content-type, mcp-protocol-version'
```

Check the published server card:

```sh
curl -fsS https://bilig.proompteng.ai/.well-known/mcp/server-card.json |
  jq '{serverInfo, tools: [.tools[].name]}'
```

Run the no-key package evaluator before trusting the agent workflow:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

Run the repo-owned OpenAI Agents SDK hosted MCP smoke if you are validating from
a Bilig checkout:

```sh
pnpm --dir examples/headless-workpaper run agent:openai-agents-sdk-hosted-mcp
```

## Boundaries

Use the ChatGPT Developer Mode app for:

- quick remote MCP tool discovery;
- proof that ChatGPT can call WorkPaper reads and verified edits;
- explaining the WorkPaper contract to another agent or teammate.

Use local file-backed stdio for:

- private workbook content;
- workflows that must persist state after the chat ends;
- approvals, quotes, budgets, forecasts, or imports that should write a project
  WorkPaper JSON file.

Use an Apps SDK component resource later when the product needs a custom ChatGPT
iframe UI. The current public proof is the tool contract: edit one input, read
dependent formulas, export or restore the WorkPaper document, and return the
exact cells that changed.
