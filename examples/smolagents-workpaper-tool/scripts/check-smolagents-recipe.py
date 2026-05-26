from __future__ import annotations

from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def require_contains(path: Path, needle: str) -> None:
    text = path.read_text(encoding="utf-8")
    if needle not in text:
        raise SystemExit(f"{path} does not contain expected text: {needle}")


def main() -> None:
    require_contains(ROOT / "smolagents_workpaper_tool.py", "from smolagents import Tool")
    require_contains(ROOT / "smolagents_workpaper_tool.py", "class BiligWorkPaperFormulaProofTool(Tool)")
    require_contains(ROOT / "smolagents_workpaper_tool.py", "bilig-agent-challenge")
    require_contains(ROOT / "README.md", "uv run --python 3.12 --with smolagents")
    require_contains(ROOT / "README.md", "verify_workpaper_formula_readback")


if __name__ == "__main__":
    main()
