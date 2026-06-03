# Google ADK WorkPaper MCP tools

This smoke test proves Google ADK's `McpToolset` can launch Bilig's file-backed
WorkPaper MCP server, edit a workbook input, read the dependent formula output,
export the document, and verify persisted readback without a model key.

Run from the repository root:

```sh
uv run --python 3.12 --with google-adk --with mcp \
  python examples/google-adk-workpaper-mcp/google_adk_workpaper_mcp.py \
  --output .tmp/google-adk-workpaper-proof.json
```

Expected result shape:

```json
{
  "framework": "google-adk",
  "toolset": "McpToolset",
  "packageSpec": "@bilig/workpaper@latest",
  "verified": true,
  "beforeExpectedArr": 60000,
  "afterExpectedArr": 96000,
  "restoredExpectedArr": 96000
}
```

The test uses the official ADK MCP imports:

```python
from google.adk.tools.mcp_tool import McpToolset
from google.adk.tools.mcp_tool.mcp_session_manager import StdioConnectionParams
from mcp import StdioServerParameters
```

`McpToolset` starts:

```sh
npm exec --yes --package @bilig/workpaper@latest -- \
  bilig-workpaper-mcp \
  --workpaper /tmp/pricing.workpaper.json \
  --init-demo-workpaper \
  --writable
```

The proof is valid only when `set_cell_contents_and_readback` changes
`Inputs!B3` to `0.4`, reads `Summary!A1:B4`, observes `Summary!B3` change from
`60000` to `96000`, persists the WorkPaper file, and confirms restored readback
matches the post-edit range.

Run the recipe guard after editing this example:

```sh
python examples/google-adk-workpaper-mcp/scripts/check-google-adk-recipe.py
```
