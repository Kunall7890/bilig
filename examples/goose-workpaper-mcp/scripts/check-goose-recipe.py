#!/usr/bin/env python3
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
REPO_ROOT = ROOT.parents[1]


def require_includes(path: Path, needle: str) -> None:
    content = path.read_text(encoding="utf-8")
    if needle not in content:
        raise SystemExit(f"{path} is missing expected text: {needle}")


def main() -> None:
    require_includes(ROOT / "README.md", "Goose WorkPaper MCP Recipe")
    require_includes(ROOT / "README.md", "goose recipe validate examples/goose-workpaper-mcp/recipe.yaml")
    require_includes(ROOT / "README.md", "set_cell_contents_and_readback")
    require_includes(ROOT / "README.md", "No upstream Goose PR or issue was opened")
    require_includes(ROOT / "recipe.yaml", 'version: "1.0.0"')
    require_includes(ROOT / "recipe.yaml", "type: stdio")
    require_includes(ROOT / "recipe.yaml", "name: bilig-workpaper")
    require_includes(ROOT / "recipe.yaml", "cmd: npm")
    require_includes(ROOT / "recipe.yaml", "- bilig-workpaper-mcp")
    require_includes(ROOT / "recipe.yaml", "- ./pricing.workpaper.json")
    require_includes(ROOT / "recipe.yaml", "set_cell_contents_and_readback")
    require_includes(REPO_ROOT / "docs" / "goose-workpaper-mcp.md", "Goose WorkPaper MCP Recipe")
    require_includes(REPO_ROOT / "docs" / "goose-workpaper-mcp.md", "https://goose-docs.ai/docs/guides/recipes/recipe-reference/")


if __name__ == "__main__":
    main()
