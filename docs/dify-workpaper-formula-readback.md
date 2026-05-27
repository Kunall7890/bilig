# Dify WorkPaper Formula Readback

Bilig can be exposed to Dify as a small tool plugin: one tool writes a workbook
input cell through the hosted WorkPaper OpenAPI endpoint, recalculates dependent
formulas, and returns JSON proof that the computed output changed and the
WorkPaper document restores to the same value.

The plugin source artifact lives at:

```text
examples/dify-workpaper-formula-readback
```

It follows Dify's tool-plugin shape: `manifest.yaml`, `provider/*.yaml`, one
tool YAML file, and one Python implementation file.

## Tool

`forecast_formula_readback` calls:

```text
POST https://bilig.proompteng.ai/openapi/workpaper/set-cell-and-readback
```

The default provider `base_url` is the hosted no-key smoke endpoint:

```text
https://bilig.proompteng.ai/openapi/workpaper
```

Set it to your own Bilig app root or OpenAPI base URL when the workbook data is
private.

Example input:

```json
{
  "sheetName": "Inputs",
  "address": "B3",
  "value": 0.4,
  "readbackRange": "Summary!A1:B3"
}
```

Example output:

```json
{
  "verified": true,
  "editedCell": "Inputs!B3",
  "readbackRange": "Summary!A1:B3",
  "before": {
    "input": 0.25,
    "expectedCustomers": 5,
    "expectedArr": 60000
  },
  "after": {
    "input": 0.4,
    "expectedCustomers": 8,
    "expectedArr": 96000
  },
  "checks": {
    "readbackChanged": true,
    "restoredReadbackMatchesAfter": true,
    "persisted": false
  }
}
```

## Why This Exists

Dify should orchestrate the agent workflow. Bilig should own spreadsheet formula
state: write the input, recalculate, read the computed output, and return proof.

That avoids a spreadsheet UI dependency and gives the agent a compact, auditable
tool result. The hosted endpoint is request-local for public smoke tests; use a
self-hosted Bilig app when Dify needs a private or persistent workbook.

## Package

Dify documents plugin manifests and packaging through its CLI:

- Manifest: <https://docs.dify.ai/en/develop-plugin/features-and-specs/plugin-types/plugin-info-by-manifest>
- CLI: <https://docs.dify.ai/en/develop-plugin/getting-started/cli>
- Local package file: <https://docs.dify.ai/en/develop-plugin/publishing/marketplace-listing/release-by-file>

From the example directory:

```sh
python -m unittest discover -s tests
uv lock
dify plugin package .
```
