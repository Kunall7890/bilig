# Microsoft Agent Framework WorkPaper MCP tools

This recipe shows how a Microsoft Agent Framework agent can use Bilig's MCP
tools for workbook formula edits without spreadsheet UI automation.

Run the package-owned no-key proof first:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

Expected result shape:

```json
{
  "door": "agent-mcp",
  "verified": true,
  "editedCell": "Inputs!B3",
  "dependentCell": "Summary!B3",
  "beforeExpectedArr": 60000,
  "afterExpectedArr": 96000
}
```

Agent Framework's Python MCP docs expose local and hosted MCP tool wrappers:

```python
from agent_framework import MCPStdioTool, MCPStreamableHTTPTool
```

Local file-backed WorkPaper MCP server:

```python
stdio_tools = MCPStdioTool(
    name="Bilig WorkPaper",
    command="npm",
    args=[
        "exec",
        "--yes",
        "--package",
        "@bilig/workpaper@latest",
        "--",
        "bilig-workpaper-mcp",
        "--workpaper",
        "./pricing.workpaper.json",
        "--init-demo-workpaper",
        "--writable",
    ],
)
```

Hosted Streamable HTTP MCP endpoint for no-key smoke tests:

```python
remote_tools = MCPStreamableHTTPTool(
    name="Bilig hosted WorkPaper MCP",
    url="https://bilig.proompteng.ai/mcp",
)
```

The proof is valid only when the agent discovers
`set_cell_contents_and_readback`, changes `Inputs!B3` to `0.4`, observes
`Summary!B3` change from `60000` to `96000`, exports the WorkPaper document,
and confirms restored readback matches the post-edit value.

No upstream Microsoft Agent Framework PR or issue was opened for this recipe.

Run the recipe guard after editing this example:

```sh
python examples/microsoft-agent-framework-workpaper-mcp/scripts/check-microsoft-agent-framework-recipe.py
```
