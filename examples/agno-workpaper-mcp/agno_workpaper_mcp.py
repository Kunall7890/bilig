from __future__ import annotations

import argparse
import asyncio
import json
import os
import tempfile
from pathlib import Path
from typing import Any

from agno.agent import Agent
from agno.models.openai import OpenAIChat
from agno.tools.function import Function, ToolResult
from agno.tools.mcp import MCPTools
from mcp import StdioServerParameters


DEFAULT_PACKAGE_SPEC = "@bilig/workpaper@latest"
READBACK_RANGE = "Summary!A1:B4"
EXPECTED_ARR_PATH = (2, 1)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Verify Bilig WorkPaper MCP tools through Agno MCPTools.")
    parser.add_argument("--package", default=DEFAULT_PACKAGE_SPEC, help="Bilig npm package spec for the MCP server.")
    parser.add_argument("--workpaper", type=Path, help="WorkPaper JSON file to create and persist during the smoke test.")
    parser.add_argument("--output", type=Path, help="Optional JSON proof output path.")
    parser.add_argument("--agent", action="store_true", help="Ask an OpenAI-backed Agno Agent to summarize the proof.")
    return parser.parse_args()


def mcp_server_params(package_spec: str, workpaper: Path) -> StdioServerParameters:
    return StdioServerParameters(
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
    )


async def call_json(function: Function, **kwargs: Any) -> dict[str, Any]:
    if function.entrypoint is None:
        raise RuntimeError(f"Agno function {function.name} has no entrypoint")
    result = await function.entrypoint(**kwargs)
    if not isinstance(result, ToolResult):
        raise RuntimeError(f"Agno function {function.name} returned {type(result).__name__}, not ToolResult")
    parsed = json.loads(result.content)
    if not isinstance(parsed, dict):
        raise RuntimeError(f"Agno function {function.name} did not return a JSON object: {result.content}")
    return parsed


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


async def build_proof(args: argparse.Namespace) -> dict[str, Any]:
    workpaper = args.workpaper or Path(tempfile.mkdtemp()) / "pricing.workpaper.json"
    server_params = mcp_server_params(args.package, workpaper)

    async with MCPTools(server_params=server_params, timeout_seconds=30) as tools:
        await tools.initialize()
        functions = tools.get_functions()
        loaded_tools = sorted(functions.keys())

        sheets = await call_json(functions["list_sheets"])
        before = await call_json(functions["read_range"], sheetName="Summary", range="A1:B4")
        write = await call_json(
            functions["set_cell_contents_and_readback"],
            sheetName="Inputs",
            address="B3",
            value=0.4,
            readbackRange=READBACK_RANGE,
        )
        after = write.get("afterReadback")
        restored = write.get("restoredReadback")
        exported = await call_json(functions["export_workpaper_document"])

    if not isinstance(after, dict) or not isinstance(restored, dict):
        raise RuntimeError(f"set_cell_contents_and_readback did not return readback objects: {json.dumps(write, indent=2)}")

    before_expected_arr = evaluated_cell_value(before, *EXPECTED_ARR_PATH)
    after_expected_arr = evaluated_cell_value(after, *EXPECTED_ARR_PATH)
    restored_expected_arr = evaluated_cell_value(restored, *EXPECTED_ARR_PATH)
    checks = write.get("checks")
    persistence = write.get("persistence")
    verified = (
        before_expected_arr == 60000
        and after_expected_arr == 96000
        and restored_expected_arr == 96000
        and isinstance(checks, dict)
        and checks.get("readbackChanged") is True
        and checks.get("restoredReadbackMatchesAfter") is True
        and isinstance(persistence, dict)
        and persistence.get("persisted") is True
    )
    if not verified:
        raise RuntimeError(
            "Agno WorkPaper MCP proof failed:\n"
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

    proof: dict[str, Any] = {
        "framework": "agno",
        "toolkit": "MCPTools",
        "packageSpec": args.package,
        "workpaper": str(workpaper),
        "loadedTools": loaded_tools,
        "sheets": sheets.get("sheets"),
        "editedCell": write.get("editedCell"),
        "readbackRange": READBACK_RANGE,
        "beforeExpectedArr": before_expected_arr,
        "afterExpectedArr": after_expected_arr,
        "restoredExpectedArr": restored_expected_arr,
        "persistence": persistence,
        "checks": checks,
        "exportedDocumentBytes": len(json.dumps(exported.get("document", {}), separators=(",", ":"))),
        "verified": True,
        "boundary": [
            "No LLM key required for the smoke test.",
            "This proves Agno MCPTools can launch Bilig WorkPaper MCP and verify workbook readback.",
            "It does not claim desktop Excel macro, pivot-table, or external-link compatibility.",
        ],
    }

    if args.agent:
        proof["agentSummary"] = run_agent_summary(proof)

    return proof


def run_agent_summary(proof: dict[str, Any]) -> str:
    if not os.environ.get("OPENAI_API_KEY"):
        raise RuntimeError("--agent requires OPENAI_API_KEY. Run without --agent for the no-key smoke test.")

    agent = Agent(
        model=OpenAIChat(id="gpt-4o-mini"),
        instructions=[
            "Summarize the workbook proof in one sentence.",
            "Only say it is verified when the JSON field verified is true.",
        ],
    )
    response = agent.run(json.dumps(proof, sort_keys=True))
    content = getattr(response, "content", response)
    return str(content)


def main() -> None:
    args = parse_args()
    proof = asyncio.run(build_proof(args))
    proof_json = json.dumps(proof, indent=2, sort_keys=True)

    if args.output is not None:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(f"{proof_json}\n", encoding="utf-8")

    print(proof_json)


if __name__ == "__main__":
    main()
