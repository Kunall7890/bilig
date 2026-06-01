from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
CLIENT = ROOT / "fastmcp_workpaper_client.py"
README = ROOT / "README.md"
MCP_JSON = ROOT / "mcp.json"

required_client_snippets = [
    "from fastmcp import Client",
    "DEFAULT_ENDPOINT = \"https://bilig.proompteng.ai/mcp\"",
    "DEFAULT_PACKAGE = \"@bilig/workpaper@latest\"",
    "EXPECTED_TOOLS = {",
    "\"set_cell_contents_and_readback\"",
    "\"export_workpaper_document\"",
    "def build_stdio_config(",
    "--transport",
    "\"The hosted endpoint is stateless",
]

required_readme_snippets = [
    "FastMCP owns the MCP client session. Bilig owns the WorkPaper formula tools",
    "uv run --python 3.12 --with 'fastmcp-slim[client]'",
    "python fastmcp_workpaper_client.py \\",
    "--transport stdio",
    "Summary!B3 = 96000",
    "`mcp.json`",
    "The default endpoint is intentionally stateless.",
    "https://gofastmcp.com/clients/client",
    "https://gofastmcp.com/clients/transports",
    "https://gofastmcp.com/community/showcase",
]

required_mcp_json_snippets = [
    "\"mcpServers\"",
    "\"bilig-workpaper\"",
    "\"transport\": \"stdio\"",
    "\"command\": \"npm\"",
    "\"@bilig/workpaper@latest\"",
    "\"bilig-workpaper-mcp\"",
    "\"--init-demo-workpaper\"",
    "\"--writable\"",
]


def require_in_file(path: Path, snippets: list[str], *, python: bool = False) -> None:
    content = path.read_text(encoding="utf-8")
    for snippet in snippets:
        if snippet not in content:
            raise SystemExit(f"{path} is missing required snippet: {snippet}")

    if python:
        compile(content, str(path), "exec")
    if path.suffix == ".json":
        json.loads(content)


def main() -> None:
    require_in_file(CLIENT, required_client_snippets, python=True)
    require_in_file(README, required_readme_snippets)
    require_in_file(MCP_JSON, required_mcp_json_snippets)
    print("FastMCP WorkPaper client example is wired.")


if __name__ == "__main__":
    main()
