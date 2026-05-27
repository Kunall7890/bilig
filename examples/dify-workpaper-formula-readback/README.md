# Bilig WorkPaper Formula Readback For Dify

This is a standalone Dify tool-plugin source artifact. It gives a Dify agent or
workflow one focused spreadsheet tool: edit a forecast input cell through the
hosted Bilig WorkPaper OpenAPI endpoint, recalculate formulas, and return
verified readback.

It intentionally lives outside the Bilig pnpm workspace. Packaging this plugin
does not require changing Bilig's root `package.json` or `pnpm-lock.yaml`.

## Tool

`forecast_formula_readback` calls:

```text
POST https://bilig.proompteng.ai/openapi/workpaper/set-cell-and-readback
```

The provider defaults to the hosted no-key smoke endpoint:

```text
https://bilig.proompteng.ai/openapi/workpaper
```

Set `base_url` to your own Bilig app root or OpenAPI base URL when the workbook
data is private.

with:

```json
{
  "sheetName": "Inputs",
  "address": "B3",
  "value": 0.4,
  "readbackRange": "Summary!A1:B3"
}
```

The tool returns JSON proof:

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

## Package

Dify packages plugins with the Dify CLI:
<https://docs.dify.ai/en/develop-plugin/getting-started/cli>.

From this directory:

```sh
uv lock
dify plugin package .
```

Dify Marketplace submissions require the plugin source plus the generated
`.difypkg` file in a directory under `langgenius/dify-plugins`, and the Dify
plugin repository documents that review flow:
<https://github.com/langgenius/dify-plugins>.

## Local Checks

```sh
python3 -m py_compile main.py provider/bilig.py tools/bilig_openapi_client.py tools/forecast_formula_readback.py
PYTHONPATH=. python3 -m unittest discover -s tests
ruby -e 'require "yaml"; %w[manifest.yaml provider/bilig.yaml tools/forecast_formula_readback.yaml].each { |p| YAML.load_file(p) }'
```
