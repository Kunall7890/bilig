from __future__ import annotations

import argparse
import asyncio
import json
import tempfile
from pathlib import Path
from typing import Any

from fastmcp.client.transports import StdioTransport
from pydantic import BaseModel, ConfigDict, Field
from pydantic_ai.mcp import MCPToolset


DEFAULT_PACKAGE_SPEC = "@bilig/workpaper@latest"
READBACK_RANGE = "Summary!A1:B4"
EXPECTED_ARR_PATH = (2, 1)


class WorkPaperProof(BaseModel):
    model_config = ConfigDict(extra="forbid")

    framework: str = "pydantic-ai"
    toolset: str = "MCPToolset"
    package_spec: str = Field(alias="packageSpec")
    workpaper: str
    loaded_tools: list[str] = Field(alias="loadedTools")
    edited_cell: str = Field(alias="editedCell")
    readback_range: str = Field(alias="readbackRange")
    before_expected_arr: float = Field(alias="beforeExpectedArr")
    after_expected_arr: float = Field(alias="afterExpectedArr")
    restored_expected_arr: float = Field(alias="restoredExpectedArr")
    persisted: bool
    restored_readback_matches_after: bool = Field(alias="restoredReadbackMatchesAfter")
    exported_document_bytes: int = Field(alias="exportedDocumentBytes")
    verified: bool
    boundary: list[str]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Verify Bilig WorkPaper MCP tools through Pydantic AI MCPToolset.")
    parser.add_argument("--package", default=DEFAULT_PACKAGE_SPEC, help="Bilig npm package spec for the MCP server.")
    parser.add_argument("--workpaper", type=Path, help="WorkPaper JSON file to create and persist during the smoke test.")
    parser.add_argument("--output", type=Path, help="Optional JSON proof output path.")
    return parser.parse_args()


def transport_for(package_spec: str, workpaper: Path) -> StdioTransport:
    return StdioTransport(
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


async def build_proof(args: argparse.Namespace) -> WorkPaperProof:
    workpaper = args.workpaper or Path(tempfile.mkdtemp()) / "pricing.workpaper.json"
    toolset = MCPToolset(
        transport_for(args.package, workpaper),
        tool_error_behavior="error",
        read_timeout=30,
    )

    tools = await toolset.list_tools()
    loaded_tools = sorted(tool.name for tool in tools)
    before = await toolset.direct_call_tool("read_range", {"sheetName": "Summary", "range": "A1:B4"})
    write = await toolset.direct_call_tool(
        "set_cell_contents_and_readback",
        {
            "sheetName": "Inputs",
            "address": "B3",
            "value": 0.4,
            "readbackRange": READBACK_RANGE,
        },
    )
    exported = await toolset.direct_call_tool("export_workpaper_document", {})

    if not isinstance(before, dict) or not isinstance(write, dict) or not isinstance(exported, dict):
        raise RuntimeError("Bilig MCP tools did not return JSON objects")

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
            "Pydantic AI WorkPaper MCP proof failed:\n"
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

    proof = WorkPaperProof.model_validate(
        {
            "packageSpec": args.package,
            "workpaper": str(workpaper),
            "loadedTools": loaded_tools,
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
                "This proves Pydantic AI MCPToolset can launch Bilig WorkPaper MCP and validate typed readback.",
                "It does not claim desktop Excel macro, pivot-table, or external-link compatibility.",
            ],
        },
    )
    return proof


def main() -> None:
    args = parse_args()
    proof = asyncio.run(build_proof(args))
    proof_json = proof.model_dump_json(by_alias=True, indent=2)

    if args.output is not None:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(f"{proof_json}\n", encoding="utf-8")

    print(proof_json)


if __name__ == "__main__":
    main()
