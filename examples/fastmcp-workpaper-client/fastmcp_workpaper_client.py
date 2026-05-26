from __future__ import annotations

import argparse
import asyncio
import json
from pathlib import Path
from typing import Any

from fastmcp import Client

DEFAULT_ENDPOINT = "https://bilig.proompteng.ai/mcp"
EXPECTED_TOOLS = {
    "list_sheets",
    "read_range",
    "read_cell",
    "set_cell_contents",
    "get_cell_display_value",
    "export_workpaper_document",
    "validate_formula",
}


def jsonable(value: Any) -> Any:
    if value is None or isinstance(value, str | int | float | bool):
        return value
    if isinstance(value, list | tuple):
        return [jsonable(item) for item in value]
    if isinstance(value, dict):
        return {str(key): jsonable(item) for key, item in value.items()}
    if hasattr(value, "model_dump"):
        return jsonable(value.model_dump(mode="json", exclude_none=True))
    return str(value)


def tool_result_data(result: Any) -> Any:
    structured_content = getattr(result, "structured_content", None)
    if structured_content is not None:
        return jsonable(structured_content)

    data = getattr(result, "data", None)
    if data is not None:
        return jsonable(data)

    raise RuntimeError(f"tool result did not expose structured content: {result!r}")


def require(condition: bool, message: str) -> None:
    if not condition:
        raise RuntimeError(message)


async def build_proof(endpoint: str) -> dict[str, Any]:
    async with Client(endpoint) as client:
        tools = await client.list_tools()
        tool_names = sorted(tool.name for tool in tools)
        missing_tools = sorted(EXPECTED_TOOLS.difference(tool_names))
        require(not missing_tools, f"missing expected Bilig MCP tools: {', '.join(missing_tools)}")

        sheets = tool_result_data(await client.call_tool("list_sheets", {}))
        expected_customers = tool_result_data(
            await client.call_tool("read_cell", {"sheetName": "Summary", "address": "B2"}),
        )
        expected_arr = tool_result_data(
            await client.call_tool("read_cell", {"sheetName": "Summary", "address": "B3"}),
        )
        edit = tool_result_data(
            await client.call_tool("set_cell_contents", {"sheetName": "Inputs", "address": "B3", "value": 0.4}),
        )
        exported = tool_result_data(await client.call_tool("export_workpaper_document", {}))

    sheet_names = [sheet["name"] for sheet in sheets.get("sheets", []) if isinstance(sheet, dict) and "name" in sheet]
    checks = {
        "hasInputsSheet": "Inputs" in sheet_names,
        "hasSummarySheet": "Summary" in sheet_names,
        "expectedCustomersFormula": expected_customers.get("formula") == "=Inputs!B2*Inputs!B3",
        "expectedArrFormula": expected_arr.get("formula") == "=B2*Inputs!B4",
        "editedExpectedConversionCell": edit.get("editedCell") == "Inputs!B3",
        "newSerializedIsPointFour": edit.get("checks", {}).get("newSerialized") == 0.4,
        "restoredMatchesAfter": edit.get("checks", {}).get("restoredMatchesAfter") is True,
        "exportedDocumentHasBytes": exported.get("serializedBytes", 0) > 0,
    }
    verified = all(checks.values())
    require(verified, f"FastMCP WorkPaper proof failed checks: {json.dumps(checks, indent=2)}")

    return {
        "client": "fastmcp",
        "transport": "streamable-http",
        "endpoint": endpoint,
        "toolNames": tool_names,
        "sheets": sheet_names,
        "readback": {
            "expectedCustomersCell": "Summary!B2",
            "expectedCustomersFormula": expected_customers.get("formula"),
            "expectedCustomersDisplayValue": expected_customers.get("displayValue"),
            "expectedArrCell": "Summary!B3",
            "expectedArrFormula": expected_arr.get("formula"),
            "expectedArrDisplayValue": expected_arr.get("displayValue"),
            "editedCell": edit.get("editedCell"),
            "previousSerialized": edit.get("checks", {}).get("previousSerialized"),
            "newSerialized": edit.get("checks", {}).get("newSerialized"),
            "restoredMatchesAfter": edit.get("checks", {}).get("restoredMatchesAfter"),
            "persisted": edit.get("checks", {}).get("persisted"),
        },
        "exportedDocumentBytes": exported.get("serializedBytes"),
        "checks": checks,
        "verified": verified,
        "limitations": [
            "The hosted endpoint is stateless and does not mutate a shared workbook across calls.",
            "Use the local file-backed bilig-workpaper-mcp stdio server for private or persisted WorkPaper JSON files.",
        ],
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Smoke-test Bilig WorkPaper MCP tools with the FastMCP Python client.")
    parser.add_argument("--endpoint", default=DEFAULT_ENDPOINT, help="Streamable HTTP MCP endpoint to call.")
    parser.add_argument("--output", type=Path, help="Optional path for the JSON proof artifact.")
    return parser.parse_args()


async def async_main() -> None:
    args = parse_args()
    proof = await build_proof(args.endpoint)
    proof_json = json.dumps(proof, indent=2, sort_keys=True)

    if args.output is not None:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(f"{proof_json}\n", encoding="utf-8")

    print(proof_json)


def main() -> None:
    asyncio.run(async_main())


if __name__ == "__main__":
    main()
