# Bilig Evaluator Proof Examples

Run these no-key checks when an agent, service, or CI evaluator needs one
production proof shape before choosing a deeper path.

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door workpaper-service --json
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- bilig-evaluate --door xlsx-cache --json
```

Every command returns `schemaVersion: "bilig-evaluator.v1"`, the selected
`door`, computed `evidence`, `verified: true`, limitations, and the underlying
source proof. The older `bilig-agent-challenge`, `bilig-mcp-challenge`,
`xlsx-cache-doctor`, and `xlsx-recalc` commands remain available when a caller
needs their full specialized output.
