# Bilig Evaluator Examples

Run these no-key checks when an agent, service, or CI job needs a workbook
readback check.

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door workpaper-service --json
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- bilig-evaluate --door xlsx-cache --json
```

Every command returns `schemaVersion: "bilig-evaluator.v1"`, the selected
`door`, `evidence`, `verified: true`, limitations, and the source command
output. `bilig-agent-challenge`, `bilig-mcp-challenge`, `xlsx-cache-doctor`,
and `xlsx-recalc` remain available when a caller needs their full output.
