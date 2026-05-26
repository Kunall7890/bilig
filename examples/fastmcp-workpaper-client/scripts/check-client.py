from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
CLIENT = ROOT / "fastmcp_workpaper_client.py"
README = ROOT / "README.md"

required_client_snippets = [
    "from fastmcp import Client",
    "DEFAULT_ENDPOINT = \"https://bilig.proompteng.ai/mcp\"",
    "EXPECTED_TOOLS = {",
    "\"set_cell_contents\"",
    "\"export_workpaper_document\"",
    "\"The hosted endpoint is stateless",
]

required_readme_snippets = [
    "FastMCP owns the MCP client session. Bilig owns the WorkPaper formula tools",
    "uv run --python 3.12 --with 'fastmcp-slim[client]'",
    "The default endpoint is intentionally stateless.",
    "https://gofastmcp.com/clients/client",
    "https://gofastmcp.com/community/showcase",
]


def require_in_file(path: Path, snippets: list[str], *, python: bool = False) -> None:
    content = path.read_text(encoding="utf-8")
    for snippet in snippets:
        if snippet not in content:
            raise SystemExit(f"{path} is missing required snippet: {snippet}")

    if python:
        compile(content, str(path), "exec")


def main() -> None:
    require_in_file(CLIENT, required_client_snippets, python=True)
    require_in_file(README, required_readme_snippets)
    print("FastMCP WorkPaper client example is wired.")


if __name__ == "__main__":
    main()
