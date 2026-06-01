from __future__ import annotations

import argparse
import asyncio
import json
from pathlib import Path
from typing import Any

from fastmcp import Client

DEFAULT_ENDPOINT = "https://bilig.proompteng.ai/mcp"
DEFAULT_WORKPAPER = Path(".tmp/fastmcp-pricing.workpaper.json")
DEFAULT_PACKAGE = "@bilig/workpaper@latest"
EXPECTED_TOOLS = {
    "list_sheets",
    "read_range",
    "read_cell",
    "set_cell_contents",
    "set_cell_contents_and_readback",
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


def numeric_cell_value(cell: dict[str, Any]) -> float | int | None:
    value = cell.get("value")
    if isinstance(value, dict):
        numeric = value.get("value")
        if isinstance(numeric, int | float):
            return numeric
    if isinstance(value, int | float):
        return value
    return None


def range_value(range_data: dict[str, Any], row_index: int, column_index: int) -> float | int | str | bool | None:
    values = range_data.get("values")
    if not isinstance(values, list) or row_index >= len(values):
        return None
    row = values[row_index]
    if not isinstance(row, list) or column_index >= len(row):
        return None
    cell = row[column_index]
    if isinstance(cell, dict):
        value = cell.get("value")
        if isinstance(value, str | int | float | bool):
            return value
    if cell is None or isinstance(cell, str | int | float | bool):
        return cell
    return None


def build_stdio_config(workpaper: Path, package_spec: str) -> dict[str, Any]:
    return {
        "mcpServers": {
            "bilig-workpaper": {
                "transport": "stdio",
                "command": "npm",
                "args": [
                    "exec",
                    "--yes",
                    "--package",
                    package_spec,
                    "--",
                    "bilig-workpaper-mcp",
                    "--workpaper",
                    str(workpaper.resolve()),
                    "--init-demo-workpaper",
                    "--writable",
                ],
                "env": {},
            }
        }
    }


async def require_expected_tools(client: Client[Any]) -> list[str]:
    tools = await client.list_tools()
    tool_names = sorted(tool.name for tool in tools)
    missing_tools = sorted(EXPECTED_TOOLS.difference(tool_names))
    require(not missing_tools, f"missing expected Bilig MCP tools: {', '.join(missing_tools)}")
    return tool_names


async def build_hosted_proof(endpoint: str) -> dict[str, Any]:
    async with Client(endpoint) as client:
        tool_names = await require_expected_tools(client)

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


async def build_stdio_proof(workpaper: Path, package_spec: str) -> dict[str, Any]:
    config = build_stdio_config(workpaper, package_spec)
    workpaper.parent.mkdir(parents=True, exist_ok=True)

    async with Client(config) as client:
        tool_names = await require_expected_tools(client)
        sheets = tool_result_data(await client.call_tool("list_sheets", {}))
        baseline = tool_result_data(
            await client.call_tool(
                "set_cell_contents_and_readback",
                {
                    "sheetName": "Inputs",
                    "address": "B3",
                    "value": 0.25,
                    "readbackRange": "Summary!A1:B3",
                },
            ),
        )
        edit = tool_result_data(
            await client.call_tool(
                "set_cell_contents_and_readback",
                {
                    "sheetName": "Inputs",
                    "address": "B3",
                    "value": 0.4,
                    "readbackRange": "Summary!A1:B3",
                },
            ),
        )
        exported = tool_result_data(await client.call_tool("export_workpaper_document", {}))

    async with Client(config) as client:
        persisted_input = tool_result_data(
            await client.call_tool("read_cell", {"sheetName": "Inputs", "address": "B3"}),
        )
        persisted_customers = tool_result_data(
            await client.call_tool("read_cell", {"sheetName": "Summary", "address": "B2"}),
        )
        persisted_arr = tool_result_data(await client.call_tool("read_cell", {"sheetName": "Summary", "address": "B3"}))

    sheet_names = [sheet["name"] for sheet in sheets.get("sheets", []) if isinstance(sheet, dict) and "name" in sheet]
    after_readback = edit.get("afterReadback", {})
    restored_readback = edit.get("restoredReadback", {})
    checks = {
        "hasInputsSheet": "Inputs" in sheet_names,
        "hasSummarySheet": "Summary" in sheet_names,
        "baselineResetToQuarter": baseline.get("checks", {}).get("newSerialized") == 0.25,
        "editedExpectedConversionCell": edit.get("editedCell") == "Inputs!B3",
        "readbackChanged": edit.get("checks", {}).get("readbackChanged") is True,
        "restoredReadbackMatchesAfter": edit.get("checks", {}).get("restoredReadbackMatchesAfter") is True,
        "persistedToWorkpaperFile": edit.get("checks", {}).get("persisted") is True,
        "newSerializedIsPointFour": edit.get("checks", {}).get("newSerialized") == 0.4,
        "expectedCustomersIsEight": range_value(after_readback, 1, 1) == 8,
        "expectedArrIsNinetySixThousand": range_value(after_readback, 2, 1) == 96000,
        "reopenedInputIsPointFour": numeric_cell_value(persisted_input) == 0.4,
        "reopenedCustomersIsEight": numeric_cell_value(persisted_customers) == 8,
        "reopenedArrIsNinetySixThousand": numeric_cell_value(persisted_arr) == 96000,
        "exportedDocumentHasBytes": exported.get("serializedBytes", 0) > 0,
    }
    verified = all(checks.values())
    require(verified, f"FastMCP local WorkPaper proof failed checks: {json.dumps(checks, indent=2)}")

    return {
        "client": "fastmcp",
        "transport": "stdio",
        "package": package_spec,
        "workpaper": str(workpaper),
        "toolNames": tool_names,
        "sheets": sheet_names,
        "readback": {
            "editedCell": edit.get("editedCell"),
            "readbackRange": edit.get("readbackRange"),
            "previousSerialized": edit.get("checks", {}).get("previousSerialized"),
            "newSerialized": edit.get("checks", {}).get("newSerialized"),
            "expectedCustomers": range_value(after_readback, 1, 1),
            "expectedArr": range_value(after_readback, 2, 1),
            "restoredExpectedCustomers": range_value(restored_readback, 1, 1),
            "restoredExpectedArr": range_value(restored_readback, 2, 1),
            "persisted": edit.get("checks", {}).get("persisted"),
            "reopenedInput": persisted_input.get("serialized"),
            "reopenedExpectedCustomers": persisted_customers.get("displayValue"),
            "reopenedExpectedArr": persisted_arr.get("displayValue"),
        },
        "exportedDocumentBytes": exported.get("serializedBytes"),
        "checks": checks,
        "verified": verified,
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Smoke-test Bilig WorkPaper MCP tools with the FastMCP Python client.")
    parser.add_argument(
        "--transport",
        choices=["hosted", "stdio"],
        default="hosted",
        help="Use the hosted endpoint or launch a local stdio MCP server.",
    )
    parser.add_argument(
        "--endpoint",
        default=DEFAULT_ENDPOINT,
        help="Streamable HTTP MCP endpoint to call in hosted mode.",
    )
    parser.add_argument("--workpaper", type=Path, default=DEFAULT_WORKPAPER, help="WorkPaper JSON path for stdio mode.")
    parser.add_argument("--package", default=DEFAULT_PACKAGE, help="npm package spec for stdio mode.")
    parser.add_argument("--output", type=Path, help="Optional path for the JSON proof artifact.")
    return parser.parse_args()


async def async_main() -> None:
    args = parse_args()
    if args.transport == "stdio":
        proof = await build_stdio_proof(args.workpaper, args.package)
    else:
        proof = await build_hosted_proof(args.endpoint)
    proof_json = json.dumps(proof, indent=2, sort_keys=True)

    if args.output is not None:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(f"{proof_json}\n", encoding="utf-8")

    print(proof_json)


def main() -> None:
    asyncio.run(async_main())


if __name__ == "__main__":
    main()
