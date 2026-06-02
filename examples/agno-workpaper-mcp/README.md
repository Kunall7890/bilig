# Agno WorkPaper MCP

Use this example when an Agno agent needs workbook formulas through MCP, not a
spreadsheet UI.

The default smoke test is keyless. It launches Bilig's published WorkPaper MCP
server over stdio, edits `Inputs!B3`, reads the calculated `Summary!B3` value,
persists the WorkPaper JSON, restores it, and writes a proof file.

```sh
uv run --python 3.12 --with agno --with mcp --with openai \
  python agno_workpaper_mcp.py --output .tmp/agno-workpaper-proof.json
```

Expected result:

```json
{
  "framework": "agno",
  "toolkit": "MCPTools",
  "packageSpec": "@bilig/workpaper@latest",
  "verified": true,
  "beforeExpectedArr": 60000,
  "afterExpectedArr": 96000
}
```

To ask an OpenAI-backed Agno agent to summarize the proof, set
`OPENAI_API_KEY` and add `--agent`.

```sh
uv run --python 3.12 --with agno --with mcp --with openai \
  python agno_workpaper_mcp.py --agent --output .tmp/agno-agent-proof.json
```

The agent path is optional. The smoke test itself does not need an LLM key.
