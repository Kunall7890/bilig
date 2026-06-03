from __future__ import annotations

import argparse
import asyncio
import json
import tempfile
from pathlib import Path
from typing import Any

from google.adk.tools.mcp_tool import McpToolset
from google.adk.tools.mcp_tool.mcp_session_manager import StdioConnectionParams
from mcp import StdioServerParameters


DEFAULT_PACKAGE_SPEC = "@bilig/workpaper@latest"
READBACK_RANGE = "Summary!A1:B4"
EXPECTED_ARR_PATH = (2, 1)
MCP_TOOL_FILTER = ["read_range", "set_cell_contents_and_readback", "export_workpaper_document"]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Verify Bilig WorkPaper MCP tools through Google ADK McpToolset.")
    parser.add_argument("--package", default=DEFAULT_PACKAGE_SPEC, help="Bilig npm package spec for the MCP server.")
    parser.add_argument("--workpaper", type=Path, help="WorkPaper JSON file to create and persist during the smoke test.")
    parser.add_argument("--output", type=Path, help="Optional JSON proof output path.")
    return parser.parse_args()


def build_toolset(package_spec: str, workpaper: Path) -> McpToolset:
    return McpToolset(
        connection_params=StdioConnectionParams(
            server_params=StdioServerParameters(
                command="npm",
                args=[
                    "exec",
                    "--yes",
                    "--package",
                    package_spec,
                    "--",
                    "bilig-workpaper-mcp",
                    "--workpaper",
                    str(workpaper),
                    "--init-demo-workpaper",
                    "--writable",
                ],
            ),
            timeout=30,
        ),
        tool_filter=MCP_TOOL_FILTER,
    )


def evaluated_cell_value(range_result: dict[str, Any], row_index: int, column_index: int) -> Any:
    values = range_result.get("values")
    if not isinstance(values, list):
        raise RuntimeError(f"readback result did not include evaluated rows: {json.dumps(range_result, indent=2)}")
    row = values[row_index]
    if not isinstance(row, list):
        raise RuntimeError(f"readback row {row_index} is not a list: {json.dumps(range_result, indent=2)}")
    cell = row[column_index]
    if not isinstance(cell, dict) or "value" not in cell:
        raise RuntimeError(f"readback cell {row_index},{column_index} did not include a value: {json.dumps(range_result, indent=2)}")
    return cell["value"]


def structured_tool_result(tool_name: str, result: Any) -> dict[str, Any]:
    if not isinstance(result, dict):
        raise RuntimeError(f"{tool_name} did not return an ADK result object: {result!r}")
    if result.get("isError") is True:
        raise RuntimeError(f"{tool_name} failed: {json.dumps(result, indent=2)}")

    structured = result.get("structuredContent")
    if isinstance(structured, dict):
        return structured

    content = result.get("content")
    if isinstance(content, list):
        for item in content:
            if isinstance(item, dict) and item.get("type") == "text" and isinstance(item.get("text"), str):
                parsed: Any = json.loads(item["text"])
                if isinstance(parsed, dict):
                    return parsed

    raise RuntimeError(f"{tool_name} did not return structured JSON content: {json.dumps(result, indent=2)}")


async def call_tool(tools_by_name: dict[str, Any], name: str, args: dict[str, Any]) -> dict[str, Any]:
    tool = tools_by_name.get(name)
    if tool is None:
        raise RuntimeError(f"ADK McpToolset did not load required tool {name!r}")
    result = await tool.run_async(args=args, tool_context=None)
    return structured_tool_result(name, result)


async def build_proof(args: argparse.Namespace) -> dict[str, Any]:
    workpaper = args.workpaper or Path(tempfile.mkdtemp()) / "pricing.workpaper.json"
    toolset = build_toolset(args.package, workpaper)
    try:
        tools = await toolset.get_tools()
        tools_by_name = {tool.name: tool for tool in tools}
        loaded_tools = sorted(tools_by_name)
        missing_tools = sorted(set(MCP_TOOL_FILTER) - set(loaded_tools))
        if missing_tools:
            raise RuntimeError(f"ADK McpToolset missed required Bilig tools: {missing_tools}")

        before = await call_tool(tools_by_name, "read_range", {"sheetName": "Summary", "range": READBACK_RANGE})
        write = await call_tool(
            tools_by_name,
            "set_cell_contents_and_readback",
            {
                "sheetName": "Inputs",
                "address": "B3",
                "value": 0.4,
                "readbackRange": READBACK_RANGE,
            },
        )
        exported = await call_tool(tools_by_name, "export_workpaper_document", {})
    finally:
        await toolset.close()

    after = write.get("afterReadback")
    restored = write.get("restoredReadback")
    checks = write.get("checks")
    persistence = write.get("persistence")
    if not isinstance(after, dict) or not isinstance(restored, dict):
        raise RuntimeError(f"set_cell_contents_and_readback did not return readback objects: {json.dumps(write, indent=2)}")
    if not isinstance(checks, dict) or not isinstance(persistence, dict):
        raise RuntimeError(f"set_cell_contents_and_readback did not return checks and persistence: {json.dumps(write, indent=2)}")

    before_expected_arr = evaluated_cell_value(before, *EXPECTED_ARR_PATH)
    after_expected_arr = evaluated_cell_value(after, *EXPECTED_ARR_PATH)
    restored_expected_arr = evaluated_cell_value(restored, *EXPECTED_ARR_PATH)
    persisted = persistence.get("persisted") is True
    restored_matches = checks.get("restoredReadbackMatchesAfter") is True
    verified = (
        before_expected_arr == 60000
        and after_expected_arr == 96000
        and restored_expected_arr == 96000
        and checks.get("readbackChanged") is True
        and persisted
        and restored_matches
    )
    if not verified:
        raise RuntimeError(
            "Google ADK WorkPaper MCP proof failed:\n"
            + json.dumps(
                {
                    "beforeExpectedArr": before_expected_arr,
                    "afterExpectedArr": after_expected_arr,
                    "restoredExpectedArr": restored_expected_arr,
                    "checks": checks,
                    "persistence": persistence,
                },
                indent=2,
            ),
        )

    return {
        "framework": "google-adk",
        "toolset": "McpToolset",
        "packageSpec": args.package,
        "workpaper": str(workpaper),
        "loadedTools": loaded_tools,
        "toolFilter": MCP_TOOL_FILTER,
        "editedCell": write.get("editedCell"),
        "readbackRange": READBACK_RANGE,
        "beforeExpectedArr": before_expected_arr,
        "afterExpectedArr": after_expected_arr,
        "restoredExpectedArr": restored_expected_arr,
        "persisted": persisted,
        "restoredReadbackMatchesAfter": restored_matches,
        "exportedDocumentBytes": len(json.dumps(exported.get("document", {}), separators=(",", ":"))),
        "verified": True,
        "boundary": [
            "No LLM key required for the smoke test.",
            "This proves Google ADK McpToolset can launch Bilig WorkPaper MCP and validate formula readback.",
            "It does not claim desktop Excel macro, pivot-table, or external-link compatibility.",
        ],
    }


def main() -> None:
    args = parse_args()
    proof = asyncio.run(build_proof(args))
    proof_json = json.dumps(proof, indent=2)

    if args.output is not None:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(f"{proof_json}\n", encoding="utf-8")

    print(proof_json)


if __name__ == "__main__":
    main()
