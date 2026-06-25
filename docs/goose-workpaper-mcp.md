---
title: Goose WorkPaper MCP recipe
published: true
description: Use Goose recipes and MCP extensions with Bilig WorkPaper to edit workbook inputs, recalculate formulas, and verify readback without spreadsheet UI automation.
tags: goose, ai-agents, mcp, spreadsheet-agent, workpaper
canonical_url: https://proompteng.github.io/bilig/goose-workpaper-mcp.html
image: /assets/github-social-preview.png
---

# Goose WorkPaper MCP Recipe

Use this page when a Goose agent needs spreadsheet-style workbook tools but
should not drive Excel, LibreOffice, Google Sheets, or a browser grid. The
owned Bilig path is a Goose recipe that launches the file-backed
`bilig-workpaper-mcp` stdio server, then requires formula readback and persisted
WorkPaper evidence before the agent trusts the result.

Run the no-key evaluator first:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

The result must include `verified: true`, discovered tools, a changed input
cell, a dependent formula readback, exported WorkPaper JSON, and restore or
restart readback. A Goose run that only says a tool was called is not enough.

## Recipe

Copy or reference the checked example recipe:

```sh
python examples/goose-workpaper-mcp/scripts/check-goose-recipe.py
goose recipe validate examples/goose-workpaper-mcp/recipe.yaml
goose run --recipe examples/goose-workpaper-mcp/recipe.yaml --debug
```

The recipe uses Goose's stdio MCP extension shape:

```yaml
version: "1.0.0"
title: Bilig WorkPaper MCP proof
description: Edit workbook inputs through Bilig WorkPaper MCP and verify formula readback.
instructions: |
  Use Bilig WorkPaper MCP tools for workbook-shaped tasks. Do not drive Excel,
  LibreOffice, Google Sheets, or browser grids for formula proof. Read the
  relevant cells, write the requested input, read the dependent formula output,
  export WorkPaper JSON, restore or restart from persisted state, and report
  editedCell, before, after, afterRestore, persistedDocumentBytes, verified,
  and limitations.
extensions:
  - type: stdio
    name: bilig-workpaper
    cmd: npm
    args:
      - exec
      - --yes
      - --package
      - "@bilig/workpaper@latest"
      - --
      - bilig-workpaper-mcp
      - --workpaper
      - ./pricing.workpaper.json
      - --init-demo-workpaper
      - --writable
    timeout: 300
    description: File-backed Bilig WorkPaper MCP server for durable workbook readback proof.
    bundled: false
```

Goose recipe files should use the `.yaml` extension. Keep the WorkPaper path
inside the project when the agent must persist state across runs.

## One-Off CLI

For an interactive Goose session without a recipe file:

```sh
goose session --with-extension "npm exec --yes --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable"
```

Then ask Goose to perform the same proof:

```text
Use Bilig WorkPaper tools to read Inputs!B3 and Summary!B3, set Inputs!B3 to
0.4, verify Summary!B3 changes from 60000 to 96000, export the WorkPaper
document, restore or restart from the persisted file, and return verified=true
only if the restored readback still matches.
```

## Hosted MCP Smoke

The hosted endpoint is useful for no-key tool discovery and stateless smoke
tests:

```sh
goose run --with-streamable-http-extension "https://bilig.proompteng.ai/mcp" -t "List Bilig WorkPaper MCP tools and verify formula readback."
```

For a recipe, Goose's Streamable HTTP extension uses `uri`:

```yaml
extensions:
  - type: streamable_http
    name: bilig-workpaper-hosted
    description: Hosted Bilig WorkPaper MCP endpoint for no-key discovery and smoke tests.
    uri: https://bilig.proompteng.ai/mcp
    timeout: 300
```

Use hosted Streamable HTTP only for hosted discovery and stateless proof. Use
stdio when a team needs a durable local WorkPaper file.

## Proof Bar

Require the Goose transcript to show:

- `set_cell_contents_and_readback`, not just `set_cell_contents`;
- edited cell `Inputs!B3`;
- dependent cell `Summary!B3`;
- formula readback `60000 -> 96000`;
- exported or persisted WorkPaper JSON;
- restore or restart readback matching the post-edit value;
- `verified: true` plus explicit limitations.

Do not describe this as an official Goose integration, Goose endorsement, Excel
compatibility certificate, or guarantee that every workbook will run. It is a
copyable owned recipe for Goose users who already decided that MCP tools should
own workbook-shaped formula proof.

## Upstream Boundary

No upstream Goose PR or issue was opened for this tranche. The prior Goose MCP
catalog PR was closed while Goose paused new MCP server additions, so this page
keeps the work on Bilig-owned docs instead of creating duplicate outreach.

## Official Goose Docs Checked

- [Using extensions](https://goose-docs.ai/docs/getting-started/using-extensions/)
- [Recipe reference](https://goose-docs.ai/docs/guides/recipes/recipe-reference/)
- [Goose CLI commands](https://goose-docs.ai/docs/guides/goose-cli-commands/)
- [Running tasks](https://goose-docs.ai/docs/guides/running-tasks/)

## Related

- [Agent WorkPaper handoff](agent-adoption-kit.md)
- [Coding agent rule chooser](agent-rule-chooser.md)
- [MCP WorkPaper tool server](mcp-workpaper-tool-server.md)
- [Agent WorkPaper proof matrix](agent-proof-matrix.md)
- [MCP directory status](mcp-spreadsheet-server-directory.md)
