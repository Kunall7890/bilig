# Bilig WorkPaper MCP Server

## Marketplace Listing

- **Name:** Bilig WorkPaper
- **Short description:** Formula-backed workbook tools for agents that need cell edits, recalculation, readback, and JSON persistence without driving Excel, LibreOffice, Google Sheets, or browser screenshots.
- **Category:** Developer Tools
- **Pricing:** Free
- **License:** MIT
- **Repository:** https://github.com/proompteng/bilig
- **Homepage:** https://proompteng.github.io/bilig/
- **Package:** https://www.npmjs.com/package/@bilig/workpaper
- **Local docs:** https://proompteng.github.io/bilig/mcp-workpaper-tool-server.html
- **Agent handbook:** https://proompteng.github.io/bilig/headless-workpaper-agent-handbook.html

## Tags

spreadsheet, workbook, formulas, xlsx, mcp, mcp-server, agents, ai-agents,
nodejs, typescript, productivity, automation, formula-recalculation,
spreadsheet-automation, workpaper

## What It Does

Bilig WorkPaper gives MCP clients a small workbook tool surface:

- list sheets before editing;
- read ranges and cells with formula and display-value context;
- validate formula syntax before writing formulas;
- write one cell at a time;
- read recalculated dependent outputs;
- export the WorkPaper JSON document for persistence or handoff.

The main use case is spreadsheet-shaped business logic inside agent tools,
backend services, workflow runners, tests, and local automation. It is useful
when an agent needs proof that a workbook edit changed the computed result and
survived persistence.

## Good Fits

- pricing and quote-approval workbooks;
- budget, forecast, commission, payout, and import-validation formulas;
- coding agents that need spreadsheet calculations but should not rely on
  screenshots as formula truth;
- local WorkPaper JSON files that an MCP client can read, edit, verify, and
  persist.

## Not A Fit

- manual spreadsheet editing as the main product;
- Office macros, add-ins, charts, pivots, or desktop Excel automation;
- arbitrary XLSX compatibility claims without a reduced workbook fixture;
- storing private workbook state in the hosted demo endpoint.

## Local Stdio Server

Requirements:

- Node.js 22 or newer;
- npm available on the same PATH as the MCP client;
- a writable project directory for the WorkPaper JSON file.

Command:

```sh
npm exec --package @bilig/workpaper@latest -- \
  bilig-workpaper-mcp \
  --workpaper ./pricing.workpaper.json \
  --init-demo-workpaper \
  --writable
```

MCP client config:

```json
{
  "mcpServers": {
    "bilig-workpaper": {
      "command": "npm",
      "args": [
        "exec",
        "--package",
        "@bilig/workpaper@latest",
        "--",
        "bilig-workpaper-mcp",
        "--workpaper",
        "./pricing.workpaper.json",
        "--init-demo-workpaper",
        "--writable"
      ]
    }
  }
}
```

`--init-demo-workpaper` only creates `pricing.workpaper.json` when it does not
exist. `--writable` is required for edit tools to persist changes back to the
file.

## Remote Smoke Endpoint

For a quick hosted smoke test:

```text
https://bilig.proompteng.ai/mcp
```

The remote endpoint is a stateless demo server. Use it to verify discovery,
tool calls, and formula readback. Use the local stdio server for private or
project-specific workbook files.

## Tools

- `list_sheets`
- `read_range`
- `read_cell`
- `set_cell_contents`
- `get_cell_display_value`
- `export_workpaper_document`
- `validate_formula`

## Resources

- `bilig://workpaper/manifest`
- `bilig://workpaper/agent-handoff`
- `bilig://workpaper/sheets`
- `bilig://workpaper/current-document`

## Prompts

- `edit_and_verify_workpaper`
- `debug_workpaper_formula`

## Verification

Run the package-owned challenge before listing or after package upgrades:

```sh
npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json
```

A passing run returns a proof object with:

- discovered tools, resources, and prompts;
- an edited input cell;
- a changed dependent formula cell;
- persisted WorkPaper JSON;
- restart readback matching the post-edit value;
- `verified: true`.

The proof should not rely on screenshots or a write-only success response.

## Security And Data Boundary

- The local stdio server reads and writes only the WorkPaper file passed with
  `--workpaper`.
- No API key is required for local usage.
- The hosted endpoint is unauthenticated and intended for demo workbooks, not
  private workbook storage.
- Agents should read before editing, validate formulas before writing them, read
  dependent outputs after recalculation, and export or persist the WorkPaper
  document before reporting success.

## Suggested Marketplace Summary

Bilig WorkPaper is a formula workbook MCP server for agents. It lets an MCP
client inspect sheets, write input cells, recalculate formulas, read the
dependent result, and persist WorkPaper JSON without controlling Excel or
trusting stale cached formula values. Use it when workbook logic belongs in a
service, workflow runner, local automation script, or coding-agent tool.
