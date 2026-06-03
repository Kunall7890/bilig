from __future__ import annotations

from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
REPO_ROOT = ROOT.parents[1]


def require_includes(path: Path, text: str) -> None:
    content = path.read_text(encoding="utf-8")
    if text not in content:
        raise SystemExit(f"{path.relative_to(REPO_ROOT)} is missing {text!r}")


def main() -> None:
    require_includes(ROOT / "README.md", "Microsoft Agent Framework WorkPaper MCP tools")
    require_includes(ROOT / "README.md", "MCPStdioTool")
    require_includes(ROOT / "README.md", "MCPStreamableHTTPTool")
    require_includes(ROOT / "README.md", "@bilig/workpaper@latest")
    require_includes(ROOT / "README.md", "set_cell_contents_and_readback")
    require_includes(ROOT / "README.md", "Inputs!B3")
    require_includes(ROOT / "README.md", "Summary!B3")
    require_includes(ROOT / "README.md", "60000")
    require_includes(ROOT / "README.md", "96000")
    require_includes(ROOT / "README.md", "No upstream Microsoft Agent Framework PR or issue was opened")
    require_includes(REPO_ROOT / "docs" / "microsoft-agent-framework-workpaper-mcp.md", "MCPStdioTool")
    require_includes(REPO_ROOT / "docs" / "microsoft-agent-framework-workpaper-mcp.md", "MCPStreamableHTTPTool")
    require_includes(
        REPO_ROOT / "docs" / "microsoft-agent-framework-workpaper-mcp.md",
        "https://learn.microsoft.com/en-us/agent-framework/agents/tools/local-mcp-tools",
    )


if __name__ == "__main__":
    main()
