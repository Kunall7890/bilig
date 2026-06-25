---
title: Spreadsheet MCP server comparison
published: true
description: 'Compare spreadsheet MCP server choices for agents: hosted Excel control, spreadsheet workspaces, file-first XLSX tools, Google Sheets tools, workbook inspection, and Bilig WorkPaper formula readback.'
tags: mcp, model context protocol, spreadsheet, excel, agents, excel mcp server
canonical_url: https://proompteng.github.io/bilig/spreadsheet-mcp-server-comparison.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Spreadsheet MCP Server Comparison

Spreadsheet MCP servers are not one category. Some control a live Excel session.
Some import workbooks into a hosted spreadsheet workspace. Some edit `.xlsx`
files. Some are Google Sheets API wrappers. Some inspect workbooks for an agent
without writing anything. Bilig WorkPaper is narrower: a local formula-backed
workbook runtime that lets an agent write known input cells, recalculate, and
return structured readback.

Use this page when you are choosing an MCP tool surface for agent workflows that
touch spreadsheet-shaped business logic.

## Quick Decision Table

| Need                                                                                                      | Better starting point                                       |
| --------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------- |
| Control a live Excel session through an add-in, paired session, OAuth, or account-backed service          | Hosted or Excel-native MCP control layer                    |
| Import an `.xlsx` into a collaborative spreadsheet workspace with Python, SQL, charts, and agent access   | Hosted spreadsheet workspace MCP                            |
| Run an agent-authored script against Excel files for rendering, linting, calculation, and structured JSON | Spreadsheet CLI or API runtime                              |
| Read and write arbitrary `.xlsx` files with formatting, charts, and workbook layout                       | Excel-focused MCP server or an Office automation workflow   |
| Read and update Google Sheets through a live cloud spreadsheet                                            | Google Sheets MCP server                                    |
| Let an agent inspect workbook structure, formulas, and cached values without mutating files               | Read-only spreadsheet inspection MCP server                 |
| Mutate service-owned workbook inputs, recalculate formulas, verify before/after values, and persist JSON  | Bilig WorkPaper MCP                                         |
| Exact Excel compatibility across macros, pivots, charts, external links, and every function               | Excel, LibreOffice, Graph API, or a dedicated Excel runtime |

## Named Public Alternatives

Use the existing spreadsheet MCP ecosystem when the source of truth is already
somewhere else:

| Server or path                                                                                 | Best fit                                                                                                                            | Boundary to check before adopting                                                                                                        |
| ---------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------- |
| [Witan](https://www.witanlabs.com/)                                                            | Agents that can call a CLI, SDK, or API against Excel files for read, write, render, calculate, lint, and structured JSON workflows | Not positioned as a public MCP server in the inspected docs; the Witan API, cloud, or self-hosted runtime is part of proof               |
| [Cellium](https://cellium.dev/)                                                                | MCP clients that need a paired Excel Add-in/session control layer with structured cell operations                                   | Requires a Cellium account/API key and live Excel runtime pairing; do not treat it as a local no-key WorkPaper runtime                   |
| [xlsx-for-ai](https://xlsx-for-ai.dev/)                                                        | Hosted API plus npm/MCP client for reading, writing, validating, diffing, and redacting Excel files                                 | Non-fallback API calls are hosted; strict mode is a privacy/error-capture setting, not proof of fresh formula recalculation              |
| [Quadratic Excel MCP](https://www.quadratichq.com/ai/mcp/excel)                                | Hosted Quadratic workspace after importing `.xlsx`, with formulas, charts, Python, SQL, OAuth, and AI clients                       | Quadratic becomes the working spreadsheet surface and exports back to `.xlsx`; it is not a local file-only MCP server                    |
| [Google Sheets MCP](https://github.com/henilcalagiya/google-sheets-mcp)                        | Agents that need CRUD operations against live Google Sheets through a service account                                               | Requires Google Cloud, Sheets API, Drive API, and service-account setup                                                                  |
| [Univer MCP](https://github.com/dream-num/univer-mcp)                                          | Agents that operate a Univer spreadsheet runtime through an MCP session                                                             | Requires an API key and a running Univer instance; the repo labels plain-text mode experimental                                          |
| [GRID MCP](https://github.com/GRID-is/claude-mcp)                                              | Claude Desktop workflows against spreadsheets uploaded to GRID                                                                      | Requires a GRID account, uploaded workbook, and API key                                                                                  |
| [mort-lab Excel MCP](https://github.com/mort-lab/excel-mcp)                                    | Openpyxl-backed local `.xlsx` creation, editing, formatting, and formula authoring                                                  | Openpyxl writes formulas but does not calculate them; `data_only` values are cached workbook values unless another engine refreshed them |
| [negokaz Excel MCP Server](https://github.com/negokaz/excel-mcp-server)                        | Local Excel workbook editing, with Windows live-Excel mode for open workbooks                                                       | Live editing and screenshots are Windows Excel COM/OLE paths; formula writes are not full proof                                          |
| [haris-musa Excel MCP Server](https://github.com/haris-musa/excel-mcp-server)                  | Openpyxl-backed Excel file mutation over stdio or HTTP transports                                                                   | Formula validation and formula writing are not the same as recalculated dependent readback                                               |
| [SheetForge MCP](https://mcpservers.org/servers/iheldan/sheetforge-mcp)                        | Local-first workbook inspection, mutation, audit, diff, repair, formula-inspection, and layout-aware agent workflows                | Its docs explicitly say read tools do not recalculate Excel formulas                                                                     |
| [CData MCP Server for Microsoft Excel](https://cdn.cdata.com/help/RXK/mcp/pg_excelformula.htm) | Commercial Excel connector with configurable read-time formula recalculation                                                        | Check connector coverage, licensing, and the `Recalculate` setting before relying on results                                             |
| Excel file or SheetJS-style tooling                                                            | Creating, reading, or preserving `.xlsx` files                                                                                      | A file library can preserve formulas without recalculating fresh results in Node                                                         |
| Bilig WorkPaper MCP                                                                            | Local agent tools that own WorkPaper JSON and need write, recalculate, readback, restore                                            | Not a full Excel editor; use it when formula readback is the product                                                                     |

That split is useful for outreach too. Do not pitch Bilig as "another Google
Sheets MCP server," "another Excel file editor," or "a hosted Excel control
layer." Pitch it where the agent needs a local formula runtime and a
machine-checkable proof object after an edit.

## Host And Account Boundary

The MCP client is not the proof. GitHub Copilot agent mode, Claude Desktop,
Cursor, VS Code, Codex, ChatGPT Apps, and similar hosts can expose configured
MCP tools to an agent, but each host still has its own approval, policy, and
tool-enablement boundary. In managed Copilot environments, MCP can also be
controlled by organization or enterprise policy.

That means a spreadsheet MCP comparison has two separate questions:

- which host can call the tool; and
- what the spreadsheet tool proves after a write.

Bilig belongs in the second column. It does not claim that Copilot, Claude, or
ChatGPT verifies workbook math by itself. It supplies a workbook-specific MCP
tool path whose evidence can include the edited input, dependent formula
readback, exported or restored WorkPaper state, and `verified: true`.

Hosted spreadsheet tools can be the right choice when the account/session is the
product. Cellium is a live Excel-control layer with API-key and session-pairing
boundaries. Quadratic is a hosted spreadsheet workspace after import. xlsx-for-ai
is a hosted API/npm MCP path for Excel-file operations. Those are legitimate
choices when the workflow wants those boundaries. They are not the same as a
no-key local WorkPaper proof.

## CLI And API Runtime Boundary

Some spreadsheet-agent tools do not try to be MCP servers. Witan is the current
public example: its spreadsheet surface is a CLI, SDK, and API path around
commands such as `witan xlsx exec`, `render`, `calc`, and `lint`. That can be a
good fit when an agent can run scripts directly against Excel files and the
desired proof includes Witan's runtime, rendering, linting, or API deployment
boundary.

That is still a different product shape from an MCP server. MCP tool discovery,
host approval, transport configuration, and tool-call readback are part of the
adoption surface for Cursor, VS Code, Claude Desktop, Codex, ChatGPT Apps, and
other clients. If the workflow wants a scriptable Excel-file runtime, start with
the CLI/API tool. If the workflow wants an MCP client to discover tools, edit a
known WorkPaper input, read a dependent formula, export state, restore it, and
return `verified: true`, use Bilig WorkPaper MCP.

## Where Bilig Fits

The Bilig MCP server is for workflows where the workbook is the service model,
not merely a file attachment. The useful loop is:

1. load a WorkPaper JSON document or the built-in demo workbook;
2. list sheets or read a range;
3. write one input cell;
4. read the recalculated display value;
5. export or persist the updated WorkPaper document.

That makes it a fit for quote approvals, payout checks, budget alerts,
import-validation workbooks, and agent tools that need proof of what changed.

It is not a replacement for a full Excel file editor. It should not be sold as
one.

## Formula Recalculation Is The Split

The important question is not "does this MCP server work with spreadsheets?"
It is "can the agent trust a formula result immediately after it writes an
input?"

Many spreadsheet MCP servers are intentionally file-oriented. That is useful
when the job is report generation, workbook inspection, or careful `.xlsx`
mutation. It is not the same as a formula-runtime loop. For example, SheetForge
MCP documents that its read tools do not recalculate Excel formulas and instead
surface formula cells as formula text. Openpyxl-backed MCP servers can write a
formula string and read cached workbook values, but openpyxl itself does not
calculate formulas. Those are the right boundaries for file tools that should
not invent fresh values.

The same user pain shows up outside MCP. A long-running SheetJS issue asks
whether a formula value can be refreshed after changing an input cell, and an
ExcelJS discussion describes JSON-driven workbook edits where shared formulas
and calculated results only become trustworthy after opening and saving in a
spreadsheet application. Those threads are not Bilig marketing claims; they are
evidence that "write XLSX" and "trust a recalculated value in Node" are separate
requirements.

Bilig takes the opposite boundary for service-owned workbooks:

- the persisted artifact is WorkPaper JSON, not an opaque Excel cache;
- the agent writes a known input cell;
- formulas recalculate inside the runtime;
- the agent reads a display value or raw value after the edit;
- the updated WorkPaper document can be exported and restored for audit.

That makes the comparison less about "best spreadsheet MCP server" and more
about the source of truth. Use file-first MCP tools when Excel fidelity is the
product. Use Bilig WorkPaper MCP when recalculated readback is the product.

## Formula Boundary Checklist

Before an agent trusts a number after a write, classify the spreadsheet server
by the proof it can return:

| Boundary                                                                        | What the agent can safely claim                                                                            |
| ------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------- |
| A hosted tool controls a paired Excel session or imported spreadsheet workspace | The live service or workspace performed the operation; the account/session/configuration is part of proof. |
| A server writes formula text into an `.xlsx` file                               | The formula was authored; the result still needs a calculation engine.                                     |
| A server reads workbook values through an `.xlsx` file library                  | The value may be a cached value from the file unless recalculation is documented.                          |
| A server can drive live desktop Excel or a commercial connector engine          | The source must stay available and the engine/configuration must be part of proof.                         |
| A server returns before/after cells from its own workbook runtime               | The agent can cite the edited input and dependent readback from the same run.                              |

That is the practical reason Bilig's MCP smoke is deliberately boring: edit a
known input, recalculate dependent formulas, export or restore the WorkPaper
document, and return the readback fields. It is a smaller claim than "Excel
replacement," but it is the claim an automation system can actually verify.

## Verify The Bilig MCP Path

Install and list the packaged server:

```sh
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp
```

Run the maintained JSON-RPC transcript from a clone:

```sh
git clone --depth 1 https://github.com/proompteng/bilig.git
cd bilig
pnpm --dir examples/headless-workpaper install --ignore-workspace
pnpm --dir examples/headless-workpaper run agent:mcp-transcript
```

The transcript edits `Inputs!B3`, recalculates dependent formulas, serializes
the WorkPaper document, restores it, and verifies that the restored values match
the post-edit values.

For a persisted workbook file:

```sh
npm exec --package @bilig/workpaper@latest -- \
  bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
```

File-backed mode exposes tools such as `list_sheets`, `read_range`,
`set_cell_contents`, `set_cell_contents_and_readback`,
`get_cell_display_value`, `export_workpaper_document`, and `validate_formula`.

## What To Ask Before Choosing A Spreadsheet MCP Server

- Is the source of truth an Excel file, a Google Sheet, or service-owned
  workbook state?
- Does the agent need to write cells, or only inspect them?
- Is stale cached formula data acceptable, or must the tool recalculate before
  responding?
- Does the workflow need exact file fidelity, or only auditable formula
  readback?
- What artifact proves the agent's edit: a screenshot, a saved file, or
  machine-checkable before/after values?

If the answer is "the backend must trust a recalculated value before it returns
or persists anything," choose a formula runtime path and keep the MCP layer thin.

## Related Bilig Pages

- [MCP spreadsheet tool server for WorkPaper agents](mcp-workpaper-tool-server.md)
- [MCP client setup](mcp-client-setup.md)
- [MCP spreadsheet server directory status](mcp-spreadsheet-server-directory.md)
- [Agent spreadsheet tool call loop](agent-spreadsheet-tool-call-loop.md)
- [Why agents need workbook APIs](why-agents-need-workbook-apis.md)
- [Stop driving spreadsheets with screenshots](stop-driving-spreadsheets-with-screenshots.md)

## Public Directory References

- [SheetForge MCP](https://mcpservers.org/servers/iheldan/sheetforge-mcp)
- [Witan spreadsheet tools](https://www.witanlabs.com/)
- [Witan CLI docs](https://docs.witanlabs.com/products/spreadsheet/sdks/cli)
- [Cellium](https://cellium.dev/)
- [xlsx-for-ai](https://xlsx-for-ai.dev/)
- [Quadratic Excel MCP](https://www.quadratichq.com/ai/mcp/excel)
- [mort-lab Excel MCP](https://github.com/mort-lab/excel-mcp)
- [negokaz Excel MCP Server](https://github.com/negokaz/excel-mcp-server)
- [haris-musa Excel MCP Server](https://github.com/haris-musa/excel-mcp-server)
- [CData MCP Server for Microsoft Excel formulas](https://cdn.cdata.com/help/RXK/mcp/pg_excelformula.htm)
- [Excel file manipulation MCP](https://mcp.directory/servers/excel-file-manipulation)
- [Bilig WorkPaper MCP registry search](https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.proompteng%2Fbilig-workpaper)
- [Bilig WorkPaper on Glama](https://glama.ai/mcp/servers/proompteng/bilig)

If this is the MCP boundary you were looking for, keep the repository and
release feed nearby:
<https://github.com/proompteng/bilig>.
