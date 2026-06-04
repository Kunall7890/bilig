# Goose WorkPaper MCP Recipe

This example gives Goose a local, file-backed Bilig WorkPaper MCP server.
Use it when a Goose agent needs to read workbook cells, edit inputs,
recalculate formulas, export WorkPaper JSON, and verify restore readback without
spreadsheet UI automation.

Run Bilig's no-key evaluator first:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

Validate the owned recipe:

```sh
python examples/goose-workpaper-mcp/scripts/check-goose-recipe.py
goose recipe validate examples/goose-workpaper-mcp/recipe.yaml
goose run --recipe examples/goose-workpaper-mcp/recipe.yaml --debug
```

The transcript should show `set_cell_contents_and_readback`, `Inputs!B3`,
`Summary!B3`, `60000`, `96000`, exported or persisted WorkPaper JSON, restore
or restart readback, and `verified: true`.

For one interactive session without a recipe file:

```sh
goose session --with-extension "npm exec --yes --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable"
```

The hosted endpoint can be used for no-key discovery and stateless smoke tests:

```sh
goose run --with-streamable-http-extension "https://bilig.proompteng.ai/mcp" -t "List Bilig WorkPaper MCP tools and verify formula readback."
```

Use the stdio recipe when durable file-backed WorkPaper state matters. Do not
describe this as an official Goose integration, Goose endorsement, Excel
compatibility certificate, or guarantee that every workbook will run.

No upstream Goose PR or issue was opened for this tranche.
