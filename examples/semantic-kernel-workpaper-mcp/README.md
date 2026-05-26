# Semantic Kernel WorkPaper MCP

Use this example when a Microsoft Semantic Kernel Python agent needs spreadsheet
formula tools through MCP. Semantic Kernel owns the plugin host. Bilig owns the
WorkPaper file, formula recalculation, JSON persistence, and readback proof.

The smoke test does not need an LLM key. It starts the Bilig WorkPaper MCP
stdio server, imports the MCP tools as a Semantic Kernel plugin, writes
`Inputs!B3` with a formula string, and verifies the recalculated readback after
JSON restore.

## Run

```sh
uv run --python 3.12 --with 'semantic-kernel[mcp]' \
  python examples/semantic-kernel-workpaper-mcp/semantic_kernel_workpaper_mcp.py \
  --local-source \
  --workpaper .tmp/pricing.workpaper.json \
  --output .tmp/semantic-kernel-workpaper-proof.json
```

Run that command from the repository root. After the published package release
catches up, omit `--local-source` to call `@bilig/workpaper@latest`.

Expected top-level output:

```json
{
  "framework": "semantic-kernel-mcp",
  "pluginName": "BiligWorkPaper",
  "verified": true
}
```

The example passes `=0.4` as a formula string because Semantic Kernel requires
MCP tool parameter schemas to use a single primitive `type`. Bilig still accepts
JSON number, boolean, and null values from MCP hosts that support union-typed
parameters.

## Local Source Smoke

From the repository root, this is the shortest local TypeScript server smoke:

```sh
uv run --python 3.12 --with 'semantic-kernel[mcp]' \
  python examples/semantic-kernel-workpaper-mcp/semantic_kernel_workpaper_mcp.py \
  --local-source \
  --workpaper .tmp/semantic-kernel-pricing.workpaper.json
```

## Boundary

This proves Semantic Kernel can import and call Bilig's MCP tools. It does not
claim desktop Excel compatibility, macro execution, external link refresh, or
mutation of arbitrary private workbooks without your own WorkPaper JSON file.

Official references:

- <https://learn.microsoft.com/en-us/semantic-kernel/concepts/plugins/adding-mcp-plugins>
- <https://devblogs.microsoft.com/agent-framework/semantic-kernel-adds-model-context-protocol-mcp-support-for-python/>
