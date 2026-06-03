from __future__ import annotations

from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def require_includes(path: Path, text: str) -> None:
    content = path.read_text(encoding="utf-8")
    if text not in content:
        raise SystemExit(f"{path.relative_to(ROOT)} is missing {text!r}")


def main() -> None:
    require_includes(ROOT / "README.md", "uv run --python 3.12 --with google-adk --with mcp")
    require_includes(ROOT / "README.md", '"verified": true')
    require_includes(ROOT / "google_adk_workpaper_mcp.py", "McpToolset")
    require_includes(ROOT / "google_adk_workpaper_mcp.py", "StdioConnectionParams")
    require_includes(ROOT / "google_adk_workpaper_mcp.py", "StdioServerParameters")
    require_includes(ROOT / "google_adk_workpaper_mcp.py", "tool_filter=MCP_TOOL_FILTER")
    require_includes(ROOT / "google_adk_workpaper_mcp.py", '"set_cell_contents_and_readback"')
    require_includes(ROOT / "google_adk_workpaper_mcp.py", '"value": 0.4')


if __name__ == "__main__":
    main()
