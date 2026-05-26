from __future__ import annotations

import argparse
import asyncio
import json
import tempfile
from pathlib import Path
from typing import Any

from semantic_kernel import Kernel
from semantic_kernel.connectors.mcp import MCPStdioPlugin


DEFAULT_PACKAGE_SPEC = "@bilig/workpaper@latest"
PLUGIN_NAME = "BiligWorkPaper"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Verify Bilig WorkPaper MCP tools through Microsoft Semantic Kernel.")
    parser.add_argument("--package", default=DEFAULT_PACKAGE_SPEC, help="Bilig npm package spec for the MCP server.")
    parser.add_argument("--workpaper", type=Path, help="WorkPaper JSON file to create and persist during the smoke test.")
    parser.add_argument("--output", type=Path, help="Optional JSON proof output path.")
    parser.add_argument(
        "--local-source",
        action="store_true",
        help="Run the local TypeScript MCP server from the repository root instead of the published npm package.",
    )
    return parser.parse_args()


def mcp_server_command(args: argparse.Namespace, workpaper: Path) -> tuple[str, list[str]]:
    server_args = [
        "--workpaper",
        str(workpaper),
        "--init-demo-workpaper",
        "--writable",
    ]
    if args.local_source:
        return "node", ["--import", "tsx", "packages/headless/src/work-paper-mcp-stdio-bin.ts", *server_args]
    return "npm", ["exec", "--yes", "--package", args.package, "--", "bilig-workpaper-mcp", *server_args]


def parse_single_json_text(contents: list[Any], label: str) -> dict[str, Any]:
    text = "\n".join(getattr(item, "text", "") for item in contents).strip()
    if not text:
        raise RuntimeError(f"{label} returned no text content")
    parsed = json.loads(text)
    if not isinstance(parsed, dict):
        raise RuntimeError(f"{label} did not return a JSON object: {text}")
    return parsed


async def run_smoke(args: argparse.Namespace) -> dict[str, Any]:
    workpaper = args.workpaper or Path(tempfile.mkdtemp()) / "pricing.workpaper.json"
    command, server_args = mcp_server_command(args, workpaper)

    async with MCPStdioPlugin(
        name=PLUGIN_NAME,
        description="Bilig WorkPaper spreadsheet formula tools",
        command=command,
        args=server_args,
        load_prompts=False,
        request_timeout=30,
    ) as plugin:
        kernel = Kernel()
        kernel.add_plugin(plugin)

        tool_names = sorted(
            name
            for name in dir(plugin)
            if name in {"list_sheets", "read_cell", "set_cell_contents", "get_cell_display_value", "validate_formula"}
        )
        if "set_cell_contents" not in tool_names:
            raise RuntimeError(f"Semantic Kernel did not import set_cell_contents. Loaded: {tool_names}")

        before = parse_single_json_text(
            await plugin.call_tool("read_cell", sheetName="Inputs", address="B3"),
            "read_cell before",
        )
        write = parse_single_json_text(
            await plugin.call_tool("set_cell_contents", sheetName="Inputs", address="B3", value="=0.4"),
            "set_cell_contents",
        )
        after = parse_single_json_text(
            await plugin.call_tool("read_cell", sheetName="Inputs", address="B3"),
            "read_cell after",
        )

    verified = (
        write.get("checks", {}).get("restoredMatchesAfter") is True
        and write.get("persistence", {}).get("persisted") is True
        and write.get("after", {}).get("value", {}).get("value") == 0.4
        and after.get("value", {}).get("value") == 0.4
    )
    if not verified:
        raise RuntimeError(f"Semantic Kernel WorkPaper proof failed: {json.dumps(write, indent=2)}")

    return {
        "framework": "semantic-kernel-mcp",
        "pluginName": PLUGIN_NAME,
        "packageSpec": "local-source" if args.local_source else args.package,
        "loadedTools": tool_names,
        "workpaper": str(workpaper),
        "editedCell": write.get("editedCell"),
        "before": {
            "serialized": before.get("serialized"),
            "displayValue": before.get("displayValue"),
        },
        "after": {
            "serialized": after.get("serialized"),
            "displayValue": after.get("displayValue"),
        },
        "persistence": write.get("persistence"),
        "checks": write.get("checks"),
        "verified": True,
        "boundary": [
            "No LLM key required for this smoke test.",
            "This proves Semantic Kernel MCP import and WorkPaper read/write/readback.",
            "It does not claim desktop Excel macro or external-link compatibility.",
        ],
    }


def main() -> None:
    args = parse_args()
    proof = asyncio.run(run_smoke(args))
    proof_json = json.dumps(proof, indent=2, sort_keys=True)

    if args.output is not None:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(f"{proof_json}\n", encoding="utf-8")

    print(proof_json)


if __name__ == "__main__":
    main()
